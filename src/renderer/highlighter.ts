// Tokenize each line of the note into spans for the syntax overlay.
// Inputs: the raw text and the per-line evaluator results (for error markers
// and to know which identifiers are user variables).

import { LineResult, EXCEL_FUNCTIONS, isExcelFunctionName, CURRENCY_SYMBOLS } from './evaluator';
import type { Mode } from '../shared/types';

export type ActiveToken =
  | { type: 'var'; name: string }   // lowercased identifier
  | { type: 'lref'; line: number }; // 1-based line number

const RESERVED_WORDS = new Set([
  'pi', 'e', 'PI', 'E', 'tau', 'phi',
  'sin', 'cos', 'tan', 'asin', 'acos', 'atan', 'atan2',
  'log', 'log2', 'log10', 'ln', 'exp', 'sqrt', 'abs', 'round', 'floor', 'ceil',
  'min', 'max', 'sum', 'mean', 'median', 'mod',
  'true', 'false', 'null'
]);

const CURRENCY_CHAR_RE = new RegExp(`^[${CURRENCY_SYMBOLS}]$`);

export interface HighlightContext {
  knownVariables: Set<string>; // names defined elsewhere in the note
}

// Columns to hang-indent a wrapped line by: the width of its leading marker
// ("- ", "1. ", "  2) ", "# ", "## ") so continuation rows line up under the
// item / heading text, Word/Google-Docs style. Returns 0 when the line has no
// such marker. The editor is monospace, so 1 column == 1ch. The list arm mirrors
// parseListLine() in popup.ts and the header arm mirrors tokenizeLine()'s hMatch
// — keep them in sync. Applied identically to every render layer (.ed-line,
// .ov-line, find layer) so the caret never drifts from the colored text.
export function listIndentCols(line: string, bulletConcealed = false): number {
  // Markdown header: "# " … "###### " (must track tokenizeLine's hMatch).
  const h = /^(\s*)(#{1,6})(\s+)/.exec(line);
  if (h) return h[1].length + h[2].length + h[3].length;
  const m = /^(\s*)(?:([-*+])|(\d+)([.)]))(\s*)(.*)$/.exec(line);
  if (!m) return 0;
  const [, indent, bullet, num, delim, spaceAfter, content] = m;
  const hasContent = content.trim().length > 0;
  if (spaceAfter.length === 0 && !(num && !hasContent)) return 0;
  let cols = indent.length + (bullet ? bullet.length : num!.length + delim!.length) + spaceAfter.length;
  // Task checkbox after a bullet ("- [ ] ..."): hang the wrap under the text AFTER
  // the box, not under the box. The box renders at a fixed 3ch (== its raw "[ ]"
  // width), so counting the raw characters keeps the indent aligned. Mirrors the
  // task branch in parseListLine() (popup.ts).
  if (bullet) {
    const task = /^(\[[ xX]\])(\s*)/.exec(content);
    if (task) {
      cols += task[1].length + task[2].length;
      // A concealed task line collapses its "- " bullet+gap away (caret off the
      // line), so the wrap then hangs under the checkbox text — drop that width.
      if (bulletConcealed) cols -= bullet.length + spaceAfter.length;
    }
  }
  return cols;
}

// Column just after a task line's "]" — lead + "- " + "[ ]", i.e. the last caret
// position still counted as "on the checkbox marker". The "- " bullet reveals at or
// before this column, but NOT in the trailing space or the text after it. Returns
// -1 when the line isn't a task line.
export function taskMarkerEndCol(line: string): number {
  const m = /^(\s*)([-*+])(\s+)(\[[ xX]\])/.exec(line);
  if (!m) return -1;
  return m[1].length + m[2].length + m[3].length + m[4].length;
}

export function highlightNote(text: string, lineResults: LineResult[], lineModes: Mode[] = [], activeToken?: ActiveToken | null, caretLine = -1, caretCol = -1): string {
  const lines = text.split('\n');
  const knownVariables = new Set<string>();
  for (const r of lineResults) {
    if (r.varName) knownVariables.add(r.varName);
  }
  const ctx: HighlightContext = { knownVariables };

  return lines
    .map((line, i) => {
      const r = lineResults[i];
      const mode: Mode = lineModes[i] ?? 'text';
      const tokens = tokenizeLine(line, r, ctx, mode, activeToken, lineResults, i === caretLine, i);
      // Each line is its own block so layoutGutters can read per-line heights and
      // so each line wraps independently: math lines get .ov-math (reserving the
      // answer column via right padding) while text lines use the full width.
      // Obsidian-style conceal. Inline bold/italic/underline (and link) markers are
      // hidden by default on text lines and revealed per-SPAN by applyMarkerReveal()
      // — only while the caret is inside that span. A task line's "- " bullet reveals
      // only while the caret sits in its leading marker region.
      const isText = mode === 'text';
      const bulletConceal = isText && !(i === caretLine && caretCol >= 0 && caretCol <= taskMarkerEndCol(line));
      let cls = mode === 'math' ? 'ov-line ov-math' : 'ov-line';
      if (bulletConceal) cls += ' ov-bullet-conceal';
      const indent = listIndentCols(line, bulletConceal);
      const style = indent > 0 ? ` style="padding-left:${indent}ch;text-indent:-${indent}ch"` : '';
      return `<div class="${cls}"${style}>${tokens || '&#8203;'}</div>`;
    })
    .join('');
}

function tokenizeLine(line: string, r: LineResult | undefined, ctx: HighlightContext, mode: Mode, activeToken: ActiveToken | null | undefined, lineResults: LineResult[], isCaretLine: boolean, lineIndex: number): string {
  if (line.length === 0) return '';

  // `/no_dec_limit` / `/clear` directive line — render as a single styled
  // token so it's visually obvious it isn't part of the math.
  if (r?.kind === 'directive') {
    const m = /^(\s*)(\/\S+)(\s*)$/.exec(line);
    if (m) {
      return `${escapeHtml(m[1])}<span class="tk-directive">${escapeHtml(m[2])}</span>${escapeHtml(m[3])}`;
    }
  }

  // Markdown header
  const hMatch = /^(\s*)(#{1,6})(\s+)(.*)$/.exec(line);
  if (hMatch) {
    const [, lead, hashes, gap, rest] = hMatch;
    const cls = hashes.length === 1 ? 'md-h1' : hashes.length === 2 ? 'md-h2' : 'md-h3';
    return `${escapeHtml(lead)}<span class="md-h-marker">${escapeHtml(hashes)}</span>${escapeHtml(gap)}<span class="${cls}">${renderInlineText(rest, lead.length + hashes.length + gap.length)}</span>`;
  }

  // Obsidian task checkbox: keep the checkbox marker at exactly 3ch wide so the
  // overlay remains aligned with the editable raw `[ ]` / `[x]` text.
  const taskMatch = /^(\s*)([-*+])(\s+)(\[[ xX]\])(\s*)(.*)$/.exec(line);
  if (taskMatch) {
    const [, lead, mark, gap, box, afterBox, rest] = taskMatch;
    const checked = /\[[xX]\]/.test(box);
    const taskCls = checked ? 'md-task-box checked' : 'md-task-box';
    const textCls = checked ? 'md-task-text checked' : 'md-task-text';
    const inner = mode === 'math' ? tokenizeMath(rest, r, ctx, activeToken) : renderTextContent(rest, lineResults, isCaretLine);
    return `${escapeHtml(lead)}<span class="md-task-bullet"><span class="md-bullet">${escapeHtml(mark)}</span>${escapeHtml(gap)}</span>` +
      `<span class="${taskCls}" data-task-line="${lineIndex}" aria-hidden="true"><span class="md-task-square"></span></span>` +
      `${escapeHtml(afterBox)}<span class="${textCls}">${inner}</span>`;
  }

  // Markdown bullet
  const bMatch = /^(\s*)([-*])(\s+)(.*)$/.exec(line);
  if (bMatch) {
    const [, lead, mark, gap, rest] = bMatch;
    const inner = mode === 'math' ? tokenizeMath(rest, r, ctx, activeToken) : renderTextContent(rest, lineResults, isCaretLine);
    return `${escapeHtml(lead)}<span class="md-bullet">${escapeHtml(mark)}</span>${escapeHtml(gap)}${inner}`;
  }

  if (mode === 'text') {
    return renderTextContent(line, lineResults, isCaretLine);
  }
  return tokenizeMath(line, r, ctx, activeToken);
}

// Render a text line's content, replacing inline references — written as
// ""L<n>"" — with the live result of that math line. On the line the caret is
// on (or when the target has no value yet) we keep the raw ""L<n>"" so it stays
// editable; otherwise we show the result, which updates on every render.
function renderTextContent(raw: string, lineResults: LineResult[], isCaretLine: boolean): string {
  const refRe = /""[lL](\d+)""/g;
  let out = '';
  let last = 0;
  let m: RegExpExecArray | null;
  while ((m = refRe.exec(raw)) !== null) {
    out += renderInlineText(raw.slice(last, m.index), last);
    out += renderRefToken(m[0], parseInt(m[1], 10), lineResults, isCaretLine);
    last = m.index + m[0].length;
  }
  out += renderInlineText(raw.slice(last), last);
  return out;
}

function renderRefToken(rawToken: string, lineNum: number, lineResults: LineResult[], isCaretLine: boolean): string {
  const res = lineResults[lineNum - 1];
  const hasVal = !!res && !res.error && (res.numeric !== undefined || res.stringValue !== undefined);
  if (isCaretLine || !hasVal) {
    // Raw ""L#"" — editable when the caret is here, or a visible placeholder
    // when the referenced line has no value yet.
    const cls = hasVal ? 'tk-textref-raw' : 'tk-textref-raw tk-textref-unresolved';
    return `<span class="${cls}">${escapeHtml(rawToken)}</span>`;
  }
  const display = res!.display ?? res!.stringValue ?? String(res!.numeric);
  return `<span class="tk-textref" title="from L${lineNum}">${escapeHtml(display)}</span>`;
}

// Tokenize a math-bearing line. Recognises numbers, identifiers, operators,
// parens, %, bps/bp, L<digit> line refs, L<a>:L<b> ranges, currency symbols,
// and `x` as a multiplication operator.
function tokenizeMath(line: string, r: LineResult | undefined, ctx: HighlightContext, activeToken?: ActiveToken | null): string {
  if (line.length === 0) return '';
  const out: string[] = [];
  // Order matters: the `%%` / `..` display-format marker BEFORE % and numbers, and
  // the line range BEFORE plain L<n> ref. The marker pattern mirrors the evaluator's
  // splitDisplayFormat: a standalone (whitespace-delimited) `%%`/`..` ANYWHERE on the
  // line, so styling and evaluation agree on what counts as a marker.
  const tokenRe = /(\s+)|((?<!\S)(?:%%|\.\.)(?=\s|$))|(L\d+\s*:\s*L\d+)|([A-Za-z_][A-Za-z0-9_]*)|([0-9][0-9,]*(?:\.[0-9]+)?|\.[0-9]+)|(%)|(:)|([+\-*/^=])|(\()|(\))|(.)/gi;
  let m: RegExpExecArray | null;
  while ((m = tokenRe.exec(line))) {
    if (m[1]) {
      out.push(escapeHtml(m[1]));
    } else if (m[2]) {
      // `%%` / `..` result-format marker — display-only, so style it muted.
      out.push(`<span class="tk-fmt">${escapeHtml(m[2])}</span>`);
    } else if (m[3]) {
      // L<a>:L<b> range
      out.push(`<span class="tk-lrange">${escapeHtml(m[3])}</span>`);
    } else if (m[4]) {
      const ident = m[4];
      const identLow = ident.toLowerCase();
      // Reserved-x error: render the bare `x` in the assignment with a
      // strong reserved-marker style instead of variable blue.
      if (/^x$/i.test(ident) && r?.errorKind === 'reserved-x') {
        out.push(`<span class="tk-reserved-x">${escapeHtml(ident)}</span>`);
      } else if (/^x$/i.test(ident)) {
        // Standalone x is a multiplication operator.
        out.push(`<span class="tk-op">${escapeHtml(ident)}</span>`);
      } else if (/^L\d+$/i.test(ident)) {
        const lineNum = parseInt(ident.slice(1));
        const isActive = activeToken?.type === 'lref' && activeToken.line === lineNum;
        const cls = isActive ? 'tk-lref tk-hl-ref' : 'tk-lref';
        out.push(`<span class="${cls}">${ident}</span>`);
      } else if (/^bps?$/i.test(ident) && /[0-9.]\s*$/.test(line.slice(0, m.index))) {
        // `bp`/`bps` is the basis-points unit ONLY when it directly follows a
        // number (e.g. `50bps`, `50 bps`) — that mirrors the evaluator, which
        // converts to basis points only when a number precedes. A standalone `bp`
        // (e.g. `bp = 4`, `f + bp`) is a normal identifier, so fall through to the
        // variable logic below and let a defined `bp` render in variable blue.
        out.push(`<span class="tk-bps">${ident}</span>`);
      } else if (isExcelFunctionName(ident) && isFollowedByParen(line, tokenRe.lastIndex)) {
        out.push(`<span class="tk-excel">${escapeHtml(ident)}</span>`);
      } else if (isExcelFunctionName(ident) && r?.errorKind === 'reserved-excel') {
        out.push(`<span class="tk-excel">${escapeHtml(ident)}</span>`);
      } else if (RESERVED_WORDS.has(identLow)) {
        out.push(`<span class="tk-fn">${ident}</span>`);
      } else if (ctx.knownVariables.has(identLow) || identLow === r?.varName) {
        const isActive = activeToken?.type === 'var' && activeToken.name === identLow;
        const cls = isActive ? 'tk-var tk-hl-ref' : 'tk-var';
        out.push(`<span class="${cls}">${ident}</span>`);
      } else {
        // Unknown identifier in a math-looking line; tag as variable but don't
        // call it an error — the user might be defining it elsewhere.
        const isActive = activeToken?.type === 'var' && activeToken.name === identLow;
        const cls = isActive ? 'tk-var tk-hl-ref' : 'tk-var';
        out.push(`<span class="${cls}">${ident}</span>`);
      }
    } else if (m[5]) {
      out.push(`<span class="tk-num">${m[5]}</span>`);
    } else if (m[6]) {
      out.push(`<span class="tk-pct">%</span>`);
    } else if (m[7]) {
      // Bare colon outside a range — just punctuation.
      out.push(`<span class="tk-op">:</span>`);
    } else if (m[8]) {
      out.push(`<span class="tk-op">${escapeHtml(m[8])}</span>`);
    } else if (m[9]) {
      out.push(`<span class="tk-paren">(</span>`);
    } else if (m[10]) {
      out.push(`<span class="tk-paren">)</span>`);
    } else if (m[11]) {
      const ch = m[11];
      if (CURRENCY_CHAR_RE.test(ch)) {
        out.push(`<span class="tk-currency">${escapeHtml(ch)}</span>`);
      } else {
        out.push(escapeHtml(ch));
      }
    }
  }
  const html = out.join('');
  // Errors are surfaced in the result column (see resultGutter rendering), not
  // by red-highlighting/underlining the line in the editor.
  return highlightInlineMarkdownPreserveSpans(html);
}

function isFollowedByParen(line: string, idx: number): boolean {
  // Skip whitespace, then check for `(`.
  while (idx < line.length && /\s/.test(line[idx])) idx++;
  return line[idx] === '(';
}

// Pass over tokenized HTML and apply **bold** and *italic* on text content.
// We avoid mangling tag attributes by skipping over `<…>` tags.
function highlightInlineMarkdownPreserveSpans(html: string): string {
  // Inline markdown only matters in plain text segments. Walk segment-by-segment.
  const parts: string[] = [];
  const segRe = /(<[^>]+>)|([^<]+)/g;
  let m: RegExpExecArray | null;
  while ((m = segRe.exec(html))) {
    if (m[1]) parts.push(m[1]);
    else if (m[2]) parts.push(highlightInlineMarkdown(m[2]));
  }
  return parts.join('');
}

function highlightInlineMarkdown(s: string): string {
  // **bold**
  s = s.replace(/\*\*([^*\n]+)\*\*/g, '<span class="md-marker">**</span><span class="md-bold">$1</span><span class="md-marker">**</span>');
  // __underline__
  s = s.replace(/__([^_\n]+)__/g, '<span class="md-marker">__</span><span class="md-underline">$1</span><span class="md-marker">__</span>');
  // *italic* (avoid matching after ** which we already replaced into spans)
  s = s.replace(/(^|[^*<>])\*([^*\n<>]+)\*(?!\*)/g,
    '$1<span class="md-marker">*</span><span class="md-italic">$2</span><span class="md-marker">*</span>');
  return s;
}

// ---- inline spans (bold/italic/underline), parsed left-to-right on RAW text so
// each formatted run knows its column range. That range drives per-SPAN conceal:
// a span's markers reveal only while the caret sits inside THAT span, not anywhere
// on the line. Both render layers (overlay .md-marker, editor .ed-mk) tag their
// markers with the same [data-cs, data-ce] so they conceal/reveal in lockstep. ----
export interface InlineSpan {
  type: 'bold' | 'italic' | 'underline' | 'link';
  outerStart: number;   // column of the first marker char
  openEnd: number;      // column where the visible content starts
  closeStart: number;   // column where the visible content ends
  outerEnd: number;     // column just past the last marker char
  url?: string;         // for links: the [text](url) target
}
export function parseInlineSpans(raw: string): InlineSpan[] {
  const spans: InlineSpan[] = [];
  let i = 0;
  while (i < raw.length) {
    // [text](url) link — the visible part is `text`; `[` and `](url)` are markers
    if (raw[i] === '[') {
      const rb = raw.indexOf(']', i + 1);
      if (rb > i + 1 && raw[rb + 1] === '(') {
        const rp = raw.indexOf(')', rb + 2);
        if (rp > rb + 2 && raw.slice(rb + 2, rp).indexOf(' ') === -1) {
          spans.push({ type: 'link', outerStart: i, openEnd: i + 1, closeStart: rb, outerEnd: rp + 1, url: raw.slice(rb + 2, rp) });
          i = rp + 1; continue;
        }
      }
    }
    // **bold** — content has no '*' (matches the /\*\*[^*\n]+\*\*/ rule)
    if (raw.startsWith('**', i)) {
      const close = raw.indexOf('**', i + 2);
      if (close > i + 2 && raw.slice(i + 2, close).indexOf('*') === -1) {
        spans.push({ type: 'bold', outerStart: i, openEnd: i + 2, closeStart: close, outerEnd: close + 2 });
        i = close + 2; continue;
      }
    }
    // __underline__ — content has no '_'
    if (raw.startsWith('__', i)) {
      const close = raw.indexOf('__', i + 2);
      if (close > i + 2 && raw.slice(i + 2, close).indexOf('_') === -1) {
        spans.push({ type: 'underline', outerStart: i, openEnd: i + 2, closeStart: close, outerEnd: close + 2 });
        i = close + 2; continue;
      }
    }
    // *italic* — a lone '*' (not part of '**'), content has no '*'
    if (raw[i] === '*' && raw[i + 1] !== '*' && raw[i - 1] !== '*') {
      const close = raw.indexOf('*', i + 1);
      if (close > i + 1 && raw[close + 1] !== '*' && raw.slice(i + 1, close).indexOf('*') === -1) {
        spans.push({ type: 'italic', outerStart: i, openEnd: i + 1, closeStart: close, outerEnd: close + 1 });
        i = close + 1; continue;
      }
    }
    i++;
  }
  return spans;
}

// "Open in new tab" icon (box with an arrow leaving the top-right corner).
const LINK_OPEN_SVG = '<svg viewBox="0 0 24 24" width="11" height="11" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M14 3h7v7"/><path d="M10 14 21 3"/><path d="M21 14v5a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5"/></svg>';

// Render a raw text run for the OVERLAY: colored formatting + markers tagged with
// their span's column range (baseCol is the run's start column in the full line).
function renderInlineText(raw: string, baseCol: number): string {
  const spans = parseInlineSpans(raw);
  if (spans.length === 0) return highlightInlineMarkdown(escapeHtml(raw));
  let out = '';
  let pos = 0;
  for (const sp of spans) {
    out += escapeHtml(raw.slice(pos, sp.outerStart));
    const attr = ` data-cs="${baseCol + sp.outerStart}" data-ce="${baseCol + sp.outerEnd}"`;
    const open = escapeHtml(raw.slice(sp.outerStart, sp.openEnd));
    const content = escapeHtml(raw.slice(sp.openEnd, sp.closeStart));
    const close = escapeHtml(raw.slice(sp.closeStart, sp.outerEnd));
    if (sp.type === 'link') {
      // Link text carries its target (for the hover tooltip / Ctrl+Click), plus a
      // little "open in new tab" icon. The icon is absolutely positioned so it adds
      // no flow width (keeping the editor + overlay aligned); a plain click on it
      // opens the URL directly (hit-tested in the editor — see popup.ts).
      const href = escapeAttr(sp.url ?? '');
      out += `<span class="md-marker"${attr}>${open}</span>` +
        `<span class="md-link" data-href="${href}">${content}` +
        `<span class="md-link-open" data-href="${href}" aria-hidden="true">${LINK_OPEN_SVG}</span></span>` +
        `<span class="md-marker"${attr}>${close}</span>`;
    } else {
      const cls = sp.type === 'bold' ? 'md-bold' : sp.type === 'underline' ? 'md-underline' : 'md-italic';
      out += `<span class="md-marker"${attr}>${open}</span><span class="${cls}">${content}</span><span class="md-marker"${attr}>${close}</span>`;
    }
    pos = sp.outerEnd;
  }
  out += escapeHtml(raw.slice(pos));
  return out;
}

// Wrap ONLY the inline-markdown markers (** __ *) of a raw line in
// <span class="ed-mk">, leaving the formatted text as plain escaped text. The
// editor (transparent edit layer) renders lines through this so those markers can
// be collapsed (concealed) in lockstep with the overlay's <span class="md-marker">
// — same patterns/order as highlightInlineMarkdown means identical marker
// positions, so the two layers stay character-aligned (monospace: bold/italic are
// the same width as plain). Lines with no markers come back as plain escaped text,
// i.e. a single text node, exactly as before.
export function wrapInlineMarkers(line: string): string {
  // A task line also conceals its "- " bullet (an unfocused task reads as just the
  // checkbox). Wrap that prefix in .ed-mk, keeping the nesting lead visible, then
  // wrap the inline markers in the rest (offset by the prefix width so the tagged
  // marker columns stay in the full-line coordinate space).
  const task = /^(\s*)([-*+])(\s+)(\[[ xX]\].*)$/.exec(line);
  if (task) {
    const [, lead, mark, gap, rest] = task;
    return escapeHtml(lead) + `<span class="ed-task-mk">${escapeHtml(mark + gap)}</span>` +
      wrapInlineSpans(rest, lead.length + mark.length + gap.length);
  }
  return wrapInlineSpans(line, 0);
}
// Editor layer: the SAME spans/columns as renderInlineText, but plain (transparent)
// text plus the .ed-mk markers — so both layers collapse/reveal identical characters
// and the caret never drifts. Each marker carries its span's [data-cs, data-ce].
function wrapInlineSpans(line: string, baseCol: number): string {
  const spans = parseInlineSpans(line);
  if (spans.length === 0) return escapeHtml(line);
  let out = '';
  let pos = 0;
  for (const sp of spans) {
    out += escapeHtml(line.slice(pos, sp.outerStart));
    const attr = ` data-cs="${baseCol + sp.outerStart}" data-ce="${baseCol + sp.outerEnd}"`;
    out += `<span class="ed-mk"${attr}>${escapeHtml(line.slice(sp.outerStart, sp.openEnd))}</span>` +
           escapeHtml(line.slice(sp.openEnd, sp.closeStart)) +
           `<span class="ed-mk"${attr}>${escapeHtml(line.slice(sp.closeStart, sp.outerEnd))}</span>`;
    pos = sp.outerEnd;
  }
  out += escapeHtml(line.slice(pos));
  return out;
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}
function escapeAttr(s: string): string {
  return escapeHtml(s).replace(/"/g, '&quot;');
}

// Re-export so popup.ts can read the function list (e.g. for the slash menu's
// help tooltip if needed). Keeps import chain shallow.
export { EXCEL_FUNCTIONS };
