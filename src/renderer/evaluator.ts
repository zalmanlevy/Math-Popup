// Evaluator: takes the full multi-line note, returns a per-line result array.
// Handles:
//   - inline `var = expr` definitions
//   - line refs (L1, L2, ...) that resolve to that line's numeric result
//   - custom suffixes from settings (e.g. `m` -> 1_000_000) when not handled
//     by the editor's eager replacement
//   - `%` and `bps`/`bp` with the "relative to LHS" rule for + / -
//   - markdown lines (#, ##, -) — header lines are skipped, bullet lines are
//     evaluated on their content portion
//   - PEMDAS via mathjs

import { create, all, type MathJsInstance } from 'mathjs';
import { Suffix } from '../shared/types';
import { formatResult, stripNumberCommas } from './formatter';

const math: MathJsInstance = create(all, { number: 'number' });

export type LineKind = 'blank' | 'header' | 'bullet' | 'assignment' | 'expression' | 'text';

export interface LineResult {
  index: number;       // zero-based
  kind: LineKind;
  raw: string;         // original line text
  display?: string;    // formatted result for the gutter ("" if none)
  numeric?: number;    // numeric result if available (for line refs / copy)
  error?: string;      // short error message if eval failed (only set when no
                       //   stale fallback is available)
  varName?: string;    // assignment target, if any
  stale?: boolean;     // true when display/numeric is the previous render's
                       //   value (e.g. user is mid-typing an incomplete expr)
}

const HEADER_RE = /^(\s*)(#{1,6})\s+(.*)$/;
const BULLET_RE = /^(\s*)([-*])\s+(.*)$/;
const ASSIGN_RE = /^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.+?)\s*$/;
const LINE_REF_RE = /\bL(\d+)\b/gi;
// Identifier that does NOT start with the letter L followed only by digits
// (so we don't match L1, L2). We'll handle refs in a separate pass.

interface PreprocessCtx {
  scope: Record<string, number>;
  results: LineResult[];   // results so far (line refs use these)
  suffixes: Suffix[];
  previous: LineResult[];  // previous render's results (for stale fallback)
  decimals: number;
}

export function evaluateNote(
  text: string,
  suffixes: Suffix[],
  previous: LineResult[] = [],
  decimals = 2
): LineResult[] {
  const lines = text.split('\n');
  const results: LineResult[] = [];
  const scope: Record<string, number> = {};

  for (let i = 0; i < lines.length; i++) {
    const raw = lines[i];
    let r = evaluateLine(raw, i, { scope, results, suffixes, previous, decimals });

    // Sticky last-good-value: while the user is mid-edit, an expression line
    // may temporarily fail to parse. Carry over the previous render's numeric
    // value (marked as stale) instead of flashing "err" red.
    if ((r.kind === 'expression' || r.kind === 'assignment' || r.kind === 'bullet') &&
        (r.error || r.numeric === undefined)) {
      const prev = previous[i];
      if (prev && prev.numeric !== undefined && isFinite(prev.numeric)) {
        r = {
          ...r,
          numeric: prev.numeric,
          display: prev.display,
          error: undefined,
          stale: true
        };
      }
    }

    results.push(r);
    if (r.varName !== undefined && r.numeric !== undefined && isFinite(r.numeric) && !r.stale) {
      scope[r.varName] = r.numeric;
    } else if (r.varName !== undefined && r.numeric !== undefined && isFinite(r.numeric) && r.stale) {
      // Even stale assignments can satisfy downstream identifier lookups
      // so dependent lines don't all flash empty.
      scope[r.varName] = r.numeric;
    }
  }
  return results;
}

function evaluateLine(raw: string, index: number, ctx: PreprocessCtx): LineResult {
  const trimmed = raw.trim();
  if (trimmed === '') {
    return { index, kind: 'blank', raw, display: '' };
  }

  // Header lines: no math, no result.
  const hdr = HEADER_RE.exec(raw);
  if (hdr) return { index, kind: 'header', raw, display: '' };

  // Bullet lines: evaluate the content portion.
  const blt = BULLET_RE.exec(raw);
  if (blt) {
    const content = blt[3];
    const sub = evaluateLine(content, index, ctx);
    return { ...sub, kind: 'bullet', raw };
  }

  // Assignment: `name = expr`
  const asn = ASSIGN_RE.exec(raw);
  if (asn) {
    const name = asn[1];
    const exprText = asn[2];
    if (isReservedName(name)) {
      // Skip evaluation, treat as plain text — avoids redefining L1, pi, etc.
      return tryEvalExpression(raw, index, ctx);
    }
    const evaluated = computeExpression(exprText, index, ctx);
    if (evaluated.error) {
      return { index, kind: 'assignment', raw, varName: name, error: evaluated.error, display: '' };
    }
    return {
      index,
      kind: 'assignment',
      raw,
      varName: name,
      numeric: evaluated.value,
      display: evaluated.value === undefined ? '' : formatResult(evaluated.value, ctx.decimals)
    };
  }

  return tryEvalExpression(raw, index, ctx);
}

function tryEvalExpression(raw: string, index: number, ctx: PreprocessCtx): LineResult {
  // If the line has no digits AND no defined identifier in scope, treat as text.
  const looksLikeMath = /[0-9]|[+\-*/^()]|%|\bL\d+\b/.test(raw) || hasKnownIdentifier(raw, ctx);
  if (!looksLikeMath) {
    return { index, kind: 'text', raw, display: '' };
  }
  const result = computeExpression(raw, index, ctx);
  if (result.error) {
    return { index, kind: 'expression', raw, error: result.error, display: '' };
  }
  if (result.value === undefined || isNaN(result.value)) {
    return { index, kind: 'text', raw, display: '' };
  }
  return {
    index,
    kind: 'expression',
    raw,
    numeric: result.value,
    display: formatResult(result.value, ctx.decimals)
  };
}

function hasKnownIdentifier(raw: string, ctx: PreprocessCtx): boolean {
  const ids = raw.match(/[A-Za-z_][A-Za-z0-9_]*/g) ?? [];
  return ids.some(id => Object.prototype.hasOwnProperty.call(ctx.scope, id));
}

function isReservedName(name: string): boolean {
  return /^L\d+$/i.test(name) || ['pi', 'e', 'PI', 'E', 'true', 'false', 'null'].includes(name);
}

function computeExpression(
  exprText: string,
  index: number,
  ctx: PreprocessCtx
): { value?: number; error?: string } {
  try {
    let s = exprText;

    // 1. Strip number commas (1,000 -> 1000).
    s = stripNumberCommas(s);

    // 2. Replace line refs L<n>.
    s = s.replace(LINE_REF_RE, (m, n) => {
      const refIdx = Number(n) - 1;
      if (refIdx < 0 || refIdx >= ctx.results.length) return 'NaN';
      const ref = ctx.results[refIdx];
      if (ref.numeric === undefined || !isFinite(ref.numeric)) return 'NaN';
      return `(${ref.numeric})`;
    });

    // 3. Apply percentage / bps preprocessor.
    s = preprocessPercentages(s);

    // 4. Replace user suffixes (e.g. `5m` -> `(5*1000000)`). Editor usually
    //    expands these eagerly into commas; this is a fallback for unexpanded
    //    text or when expandSuffixesInEditor is off.
    s = applySuffixes(s, ctx.suffixes);

    if (s.trim() === '') return {};
    const value = math.evaluate(s, { ...ctx.scope });
    if (typeof value === 'number') return { value };
    if (typeof value === 'object' && value !== null && 'valueOf' in value) {
      const v = (value as { valueOf(): unknown }).valueOf();
      if (typeof v === 'number') return { value: v };
    }
    return {};
  } catch (e: unknown) {
    const msg = e instanceof Error ? e.message : String(e);
    return { error: shortError(msg) };
  }
}

function preprocessPercentages(s: string): string {
  // bps and bp first (longer suffix), with leading +/- as relative to LHS.
  s = s.replace(/([+\-])\s*([0-9]*\.?[0-9]+)\s*bps?\b/gi, (_m, op, n) =>
    op === '+' ? `*(1+${n}/10000)` : `*(1-${n}/10000)`);
  s = s.replace(/([0-9]*\.?[0-9]+)\s*bps?\b/gi, (_m, n) => `(${n}/10000)`);

  // Percentages: relative-to-LHS for +/-, otherwise literal /100.
  s = s.replace(/([+\-])\s*([0-9]*\.?[0-9]+)\s*%/g, (_m, op, n) =>
    op === '+' ? `*(1+${n}/100)` : `*(1-${n}/100)`);
  s = s.replace(/([0-9]*\.?[0-9]+)\s*%/g, (_m, n) => `(${n}/100)`);
  return s;
}

function applySuffixes(s: string, suffixes: Suffix[]): string {
  if (!suffixes.length) return s;
  // Sort longest-symbol-first to avoid `mb` being matched as `m` then `b`.
  const sorted = [...suffixes].sort((a, b) => b.symbol.length - a.symbol.length);
  for (const suf of sorted) {
    const flags = suf.caseSensitive ? 'g' : 'gi';
    // Match a number directly followed by the suffix, with right-side word boundary.
    const escaped = escapeRegex(suf.symbol);
    const re = new RegExp(`([0-9]*\\.?[0-9]+)${escaped}(?![A-Za-z0-9_])`, flags);
    s = s.replace(re, (_m, n) => `(${n}*${suf.multiplier})`);
  }
  return s;
}

function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function shortError(msg: string): string {
  // mathjs error messages are sometimes long; show just the first sentence.
  const idx = msg.indexOf('\n');
  const first = idx >= 0 ? msg.slice(0, idx) : msg;
  return first.length > 80 ? first.slice(0, 77) + '…' : first;
}
