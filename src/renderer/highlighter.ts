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

export function highlightNote(text: string, lineResults: LineResult[], mode: Mode = 'math', activeToken?: ActiveToken | null): string {
  const lines = text.split('\n');
  const knownVariables = new Set<string>();
  for (const r of lineResults) {
    if (r.varName) knownVariables.add(r.varName);
  }
  const ctx: HighlightContext = { knownVariables };

  return lines
    .map((line, i) => {
      const r = lineResults[i];
      const tokens = tokenizeLine(line, r, ctx, mode, activeToken);
      // Each line wrapped in its own block element so layoutGutters can read
      // per-line rendered heights (handles wrap accurately because the overlay
      // shares width/font/wrap rules with the editor).
      return `<div class="ov-line">${tokens || '&#8203;'}</div>`;
    })
    .join('');
}

function tokenizeLine(line: string, r: LineResult | undefined, ctx: HighlightContext, mode: Mode, activeToken?: ActiveToken | null): string {
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
    return `${escapeHtml(lead)}<span class="md-h-marker">${escapeHtml(hashes)}</span>${escapeHtml(gap)}<span class="${cls}">${highlightInlineMarkdown(escapeHtml(rest))}</span>`;
  }

  // Markdown bullet
  const bMatch = /^(\s*)([-*])(\s+)(.*)$/.exec(line);
  if (bMatch) {
    const [, lead, mark, gap, rest] = bMatch;
    const inner = mode === 'math' ? tokenizeMath(rest, r, ctx, activeToken) : highlightInlineMarkdown(escapeHtml(rest));
    return `${escapeHtml(lead)}<span class="md-bullet">${escapeHtml(mark)}</span>${escapeHtml(gap)}${inner}`;
  }

  if (mode === 'text') {
    return highlightInlineMarkdown(escapeHtml(line));
  }
  return tokenizeMath(line, r, ctx, activeToken);
}

// Tokenize a math-bearing line. Recognises numbers, identifiers, operators,
// parens, %, bps/bp, L<digit> line refs, L<a>:L<b> ranges, currency symbols,
// and `x` as a multiplication operator.
function tokenizeMath(line: string, r: LineResult | undefined, ctx: HighlightContext, activeToken?: ActiveToken | null): string {
  if (line.length === 0) return '';
  const out: string[] = [];
  // Order matters: line range BEFORE plain L<n> ref.
  const tokenRe = /(\s+)|(L\d+\s*:\s*L\d+)|([A-Za-z_][A-Za-z0-9_]*)|([0-9][0-9,]*(?:\.[0-9]+)?|\.[0-9]+)|(%)|(:)|([+\-*/^=])|(\()|(\))|(.)/gi;
  let m: RegExpExecArray | null;
  while ((m = tokenRe.exec(line))) {
    if (m[1]) {
      out.push(escapeHtml(m[1]));
    } else if (m[2]) {
      // L<a>:L<b> range
      out.push(`<span class="tk-lrange">${escapeHtml(m[2])}</span>`);
    } else if (m[3]) {
      const ident = m[3];
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
      } else if (/^bps?$/i.test(ident)) {
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
    } else if (m[4]) {
      out.push(`<span class="tk-num">${m[4]}</span>`);
    } else if (m[5]) {
      out.push(`<span class="tk-pct">%</span>`);
    } else if (m[6]) {
      // Bare colon outside a range — just punctuation.
      out.push(`<span class="tk-op">:</span>`);
    } else if (m[7]) {
      out.push(`<span class="tk-op">${escapeHtml(m[7])}</span>`);
    } else if (m[8]) {
      out.push(`<span class="tk-paren">(</span>`);
    } else if (m[9]) {
      out.push(`<span class="tk-paren">)</span>`);
    } else if (m[10]) {
      const ch = m[10];
      if (CURRENCY_CHAR_RE.test(ch)) {
        out.push(`<span class="tk-currency">${escapeHtml(ch)}</span>`);
      } else {
        out.push(escapeHtml(ch));
      }
    }
  }
  let html = out.join('');
  if (r?.error && r.errorKind !== 'reserved-x' && r.errorKind !== 'reserved-excel'
      && r.errorKind !== 'reserved-name') {
    html = `<span class="tk-error">${html}</span>`;
  }
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
  // *italic* (avoid matching after ** which we already replaced into spans)
  s = s.replace(/(^|[^*<>])\*([^*\n<>]+)\*(?!\*)/g,
    '$1<span class="md-marker">*</span><span class="md-italic">$2</span><span class="md-marker">*</span>');
  return s;
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

// Re-export so popup.ts can read the function list (e.g. for the slash menu's
// help tooltip if needed). Keeps import chain shallow.
export { EXCEL_FUNCTIONS };
