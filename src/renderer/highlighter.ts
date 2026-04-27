// Tokenize each line of the note into spans for the syntax overlay.
// Inputs: the raw text and the per-line evaluator results (for error markers
// and to know which identifiers are user variables).

import { LineResult } from './evaluator';
import type { Mode } from '../shared/types';

const RESERVED_WORDS = new Set([
  'pi', 'e', 'PI', 'E', 'tau', 'phi',
  'sin', 'cos', 'tan', 'asin', 'acos', 'atan', 'atan2',
  'log', 'log2', 'log10', 'ln', 'exp', 'sqrt', 'abs', 'round', 'floor', 'ceil',
  'min', 'max', 'sum', 'mean', 'median', 'mod',
  'true', 'false', 'null'
]);

export interface HighlightContext {
  knownVariables: Set<string>; // names defined elsewhere in the note
}

export function highlightNote(text: string, lineResults: LineResult[], mode: Mode = 'math'): string {
  const lines = text.split('\n');
  const knownVariables = new Set<string>();
  for (const r of lineResults) {
    if (r.varName) knownVariables.add(r.varName);
  }
  const ctx: HighlightContext = { knownVariables };

  return lines
    .map((line, i) => {
      const r = lineResults[i];
      const tokens = tokenizeLine(line, r, ctx, mode);
      // Need a trailing space so a final empty line still has measurable height.
      return tokens || '&#8203;';
    })
    .join('\n');
}

function tokenizeLine(line: string, r: LineResult | undefined, ctx: HighlightContext, mode: Mode): string {
  if (line.length === 0) return '';

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
    const inner = mode === 'math' ? tokenizeMath(rest, r, ctx) : highlightInlineMarkdown(escapeHtml(rest));
    return `${escapeHtml(lead)}<span class="md-bullet">${escapeHtml(mark)}</span>${escapeHtml(gap)}${inner}`;
  }

  if (mode === 'text') {
    return highlightInlineMarkdown(escapeHtml(line));
  }
  return tokenizeMath(line, r, ctx);
}

// Tokenize a math-bearing line. Recognises numbers, identifiers, operators,
// parens, %, bps/bp, and L<digit> line refs.
function tokenizeMath(line: string, r: LineResult | undefined, ctx: HighlightContext): string {
  if (line.length === 0) return '';
  const out: string[] = [];
  const tokenRe = /(\s+)|([A-Za-z_][A-Za-z0-9_]*)|([0-9][0-9,]*(?:\.[0-9]+)?|\.[0-9]+)|(%)|([+\-*/^=])|(\()|(\))|(.)/g;
  let m: RegExpExecArray | null;
  while ((m = tokenRe.exec(line))) {
    if (m[1]) {
      out.push(escapeHtml(m[1]));
    } else if (m[2]) {
      const ident = m[2];
      if (/^L\d+$/.test(ident)) {
        out.push(`<span class="tk-lref">${ident}</span>`);
      } else if (/^bps?$/i.test(ident)) {
        out.push(`<span class="tk-bps">${ident}</span>`);
      } else if (RESERVED_WORDS.has(ident)) {
        out.push(`<span class="tk-fn">${ident}</span>`);
      } else if (ctx.knownVariables.has(ident) || ident === r?.varName) {
        out.push(`<span class="tk-var">${ident}</span>`);
      } else {
        // Unknown identifier in a math-looking line; tag as variable but don't
        // call it an error — the user might be defining it elsewhere.
        out.push(`<span class="tk-var">${ident}</span>`);
      }
    } else if (m[3]) {
      out.push(`<span class="tk-num">${m[3]}</span>`);
    } else if (m[4]) {
      out.push(`<span class="tk-pct">%</span>`);
    } else if (m[5]) {
      out.push(`<span class="tk-op">${escapeHtml(m[5])}</span>`);
    } else if (m[6]) {
      out.push(`<span class="tk-paren">(</span>`);
    } else if (m[7]) {
      out.push(`<span class="tk-paren">)</span>`);
    } else if (m[8]) {
      out.push(escapeHtml(m[8]));
    }
  }
  let html = out.join('');
  if (r?.error) {
    html = `<span class="tk-error">${html}</span>`;
  }
  return highlightInlineMarkdownPreserveSpans(html);
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
