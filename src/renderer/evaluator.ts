// Evaluator: takes the full multi-line note, returns a per-line result array.
// Handles:
//   - inline `var = expr` definitions
//   - line refs (L1, L2, ...) that resolve to that line's numeric result
//   - line ranges (L1:L4) used inside aggregate functions
//   - custom suffixes from settings (e.g. `m` -> 1_000_000) when not handled
//     by the editor's eager replacement
//   - `%` and `bps`/`bp` with the "relative to LHS" rule for + / -
//   - markdown lines (#, ##, -) — header lines are skipped, bullet lines are
//     evaluated on their content portion
//   - PEMDAS via mathjs
//   - x as a multiplication operator
//   - money symbols ($€£¥...) ignored when evaluating
//   - Excel-style aggregate functions (SUM, AVERAGE, MEAN, MAX, MIN, COUNT,
//     MEDIAN, ROUND, CEIL, FLOOR, ABS, IF, TODAY, NOW, SQRT)

import { create, all, type MathJsInstance } from 'mathjs';
import { Suffix } from '../shared/types';
import { formatResult, stripNumberCommas } from './formatter';

const math: MathJsInstance = create(all, { number: 'number' });

// Register a few Excel-style helpers that aren't built-ins. Names are stored
// lowercase here — `computeExpression` lowercases Excel function names in the
// user's expression before passing to mathjs so SUM/Sum/sum all work.
// Walk arrays and mathjs Matrix-like objects, collecting finite numbers.
function flattenNumbers(args: unknown[]): number[] {
  const nums: number[] = [];
  const walk = (v: unknown) => {
    if (Array.isArray(v)) {
      v.forEach(walk);
    } else if (v && typeof v === 'object' &&
               typeof (v as { toArray?: unknown }).toArray === 'function') {
      try { walk((v as { toArray: () => unknown[] }).toArray()); } catch { /* ignore */ }
    } else if (typeof v === 'number' && isFinite(v)) {
      nums.push(v);
    }
  };
  args.forEach(walk);
  return nums;
}

math.import({
  today: () => Math.floor(Date.now() / 86_400_000),  // days since Unix epoch
  now: () => Date.now() / 1000,                       // seconds since Unix epoch
  // average / mean: empty input returns NaN (so the line shows no result)
  // instead of 0 — averaging nothing isn't 0, and we don't want an artificial
  // 0 to leak in from an empty L<a>:L<b> range.
  average: (...args: unknown[]) => {
    const nums = flattenNumbers(args);
    if (nums.length === 0) return NaN;
    return nums.reduce((a, b) => a + b, 0) / nums.length;
  },
  mean: (...args: unknown[]) => {
    const nums = flattenNumbers(args);
    if (nums.length === 0) return NaN;
    return nums.reduce((a, b) => a + b, 0) / nums.length;
  },
  // sum overridden so an empty L<a>:L<b> range (which expands to `[]`) still
  // returns 0 rather than throwing. Works for any mix of numbers / arrays.
  sum: (...args: unknown[]) => {
    const nums = flattenNumbers(args);
    return nums.reduce((a, b) => a + b, 0);
  },
  count: (...args: unknown[]) => flattenNumbers(args).length,
  // Excel-style IF: defaults to "TRUE"/"FALSE" string output when the user
  // omits the value-if-true / value-if-false arguments. Treats undefined,
  // 0, NaN, and empty string as falsy; everything else as truthy.
  if: (cond: unknown, a: unknown, b: unknown) => {
    const aval = a === undefined ? 'TRUE' : a;
    const bval = b === undefined ? 'FALSE' : b;
    const truthy =
      cond !== undefined &&
      cond !== false &&
      cond !== 0 &&
      cond !== '' &&
      !(typeof cond === 'number' && isNaN(cond));
    return truthy ? aval : bval;
  }
}, { override: true });

export type LineKind = 'blank' | 'header' | 'bullet' | 'assignment' | 'expression' | 'text' | 'directive';

export interface LineResult {
  index: number;       // zero-based
  kind: LineKind;
  raw: string;         // original line text
  display?: string;    // formatted result for the gutter ("" if none)
  numeric?: number;    // numeric result if available (for line refs / copy)
  stringValue?: string; // string result (e.g. IF returning "TRUE"/"FALSE")
  error?: string;      // short error message if eval failed (only set when no
                       //   stale fallback is available)
  errorKind?: 'reserved-x' | 'reserved-excel' | 'reserved-name' | 'unquoted-string' | 'duplicate-var' | 'general';
  errorTooltip?: string;  // longer message shown on hover for special errors
  varName?: string;    // assignment target, if any
  stale?: boolean;     // true when display/numeric is the previous render's
                       //   value (e.g. user is mid-typing an incomplete expr)
}

const HEADER_RE = /^(\s*)(#{1,6})\s+(.*)$/;
const BULLET_RE = /^(\s*)([-*])\s+(.*)$/;
const ASSIGN_RE = /^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.+?)\s*$/;
const LINE_REF_RE = /\bL(\d+)\b/gi;
const LINE_RANGE_RE = /\bL(\d+)\s*:\s*L(\d+)\b/gi;
export const DIRECTIVE_NO_DEC_LIMIT_RE = /^\s*\/no_dec_limit\s*$/i;
export const DIRECTIVE_CLEAR_RE = /^\s*\/clear\s*$/i;
// Cap used for lines below a `/no_dec_limit` directive. BA II Plus-style:
// show however many decimals the value has, but never more than this.
export const NO_DEC_LIMIT_CAP = 6;

// Money / currency symbols stripped silently before evaluation.
export const CURRENCY_SYMBOLS = '$€£¥₹₽¢₩₪₫₴₸₺฿';
export const CURRENCY_RE = new RegExp(`[${CURRENCY_SYMBOLS}]`, 'g');

// Excel-style aggregate / spreadsheet functions. Listed in upper-case for the
// help tooltip — the actual matching is case-insensitive. Names listed here:
//   - cannot be assigned (`sum = 5` -> reserved-excel error)
//   - the `x` rule from `x` is handled separately
//   - must be followed by `(` to be recognised as a call (matches Excel)
export const EXCEL_FUNCTIONS = [
  'SUM', 'AVERAGE', 'MEAN', 'MAX', 'MIN', 'COUNT', 'MEDIAN',
  'ROUND', 'CEIL', 'FLOOR', 'ABS', 'IF', 'TODAY', 'NOW', 'SQRT'
] as const;
const EXCEL_FUNCTION_SET = new Set(EXCEL_FUNCTIONS.map(s => s.toLowerCase()));
export const EXCEL_FORMULA_TOOLTIP =
  'This name is reserved as an Excel-style formula and cannot be used as a variable. ' +
  'Available formulas: SUM, AVERAGE, MEAN, MAX, MIN, COUNT, MEDIAN, ROUND, CEIL, FLOOR, ' +
  'ABS, IF, TODAY, NOW, SQRT. Use them like Excel: e.g. SUM(L1:L4) or ROUND(L1, 2).';
export const X_RESERVED_TOOLTIP =
  'The letter "x" is reserved as a multiplication operator (like "*"), so it cannot be used as a variable.';
export const UNQUOTED_STRING_TOOLTIP =
  'Add Quotations — wrap text values in quotes (e.g. "YES" instead of YES). The only words that work without quotes are TRUE and FALSE.';
export const RESERVED_NAME_TOOLTIP =
  'This name is reserved (line refs like L1/L2, constants pi/e, or true/false/null) and cannot be used as a variable.';
export const DUPLICATE_VAR_TOOLTIP =
  'This variable is already defined on a previous line. Rename one of them to avoid conflicts.';

interface PreprocessCtx {
  scope: Record<string, number>;
  results: LineResult[];   // results so far (line refs use these)
  snapshot?: LineResult[]; // results from previous iteration (forward-ref lookup)
  suffixes: Suffix[];
  previous: LineResult[];  // previous render's results (for stale fallback)
  decimals: number;
  definedInThisPass?: Set<string>; // variable names assigned so far in this pass
}

export function isExcelFunctionName(name: string): boolean {
  return EXCEL_FUNCTION_SET.has(name.toLowerCase());
}

export function isReservedXName(name: string): boolean {
  return name.toLowerCase() === 'x';
}

export function evaluateNote(
  text: string,
  suffixes: Suffix[],
  previous: LineResult[] = [],
  decimals = 2
): LineResult[] {
  const lines = text.split('\n');
  // Iterative evaluation so forward `L<n>` references resolve. Each pass uses
  // the previous pass's full result array as a snapshot for forward lookups.
  // Two passes cover most cases; three handles chains of forward refs.
  const MAX_ITER = 3;
  let results: LineResult[] = previous.slice(0, lines.length);
  for (let iter = 0; iter < MAX_ITER; iter++) {
    const snapshot = results;
    const next = evaluateOnePass(lines, suffixes, snapshot, previous, decimals);
    if (resultsConverged(next, results)) {
      results = next;
      break;
    }
    results = next;
  }
  return results;
}

function evaluateOnePass(
  lines: string[],
  suffixes: Suffix[],
  snapshot: LineResult[],
  previous: LineResult[],
  decimals: number
): LineResult[] {
  const results: LineResult[] = [];
  const scope: Record<string, number> = {};
  const definedInThisPass = new Set<string>();
  let currentDecimals = decimals;

  // Pre-seed scope from the previous pass so forward variable references
  // (a variable used above where it's defined) resolve on pass 2+.
  // Lines evaluated top-to-bottom will overwrite these as they're reached,
  // so the last definition still wins.
  for (const r of snapshot) {
    if (r.varName !== undefined && r.numeric !== undefined && isFinite(r.numeric)) {
      scope[r.varName] = r.numeric;
    }
  }

  for (let i = 0; i < lines.length; i++) {
    const raw = lines[i];

    // `/no_dec_limit` directive: from this line down, switch to a 6-decimal
    // cap (overrides the user's settings.decimals). The directive line itself
    // produces no result.
    if (DIRECTIVE_NO_DEC_LIMIT_RE.test(raw)) {
      results.push({ index: i, kind: 'directive', raw, display: '' });
      currentDecimals = NO_DEC_LIMIT_CAP;
      continue;
    }

    // `/clear` directive: visible-only marker. The actual editor clear is
    // handled in the renderer (popup.ts) when the slash menu fires it. If a
    // stray /clear line slips into evaluation it's just a no-result directive.
    if (DIRECTIVE_CLEAR_RE.test(raw)) {
      results.push({ index: i, kind: 'directive', raw, display: '' });
      continue;
    }

    let r = evaluateLine(raw, i, { scope, results, snapshot, suffixes, previous, decimals: currentDecimals, definedInThisPass });

    // Sticky last-good-value: while the user is mid-edit, an expression line
    // may temporarily fail to parse. Carry over the previous render's numeric
    // value (marked as stale) instead of flashing "err" red. Skip for the
    // special reserved-name errors — those are user-facing intentional errors
    // and we don't want them to ever look successful.
    const isReservedErr = r.errorKind === 'reserved-x' || r.errorKind === 'reserved-excel'
      || r.errorKind === 'reserved-name' || r.errorKind === 'unquoted-string'
      || r.errorKind === 'duplicate-var';
    const hasResult = r.numeric !== undefined || r.stringValue !== undefined;
    if (!isReservedErr &&
        (r.kind === 'expression' || r.kind === 'assignment' || r.kind === 'bullet') &&
        (r.error || !hasResult)) {
      const prev = previous[i];
      const prevHas = prev && (prev.numeric !== undefined && isFinite(prev.numeric)
                               || prev.stringValue !== undefined);
      if (prevHas && prev) {
        r = {
          ...r,
          numeric: prev.numeric,
          stringValue: prev.stringValue,
          display: prev.display,
          error: undefined,
          stale: true
        };
      }
    }

    results.push(r);
    if (r.varName !== undefined && r.numeric !== undefined && isFinite(r.numeric)) {
      scope[r.varName] = r.numeric;
    }
  }
  return results;
}

function resultsConverged(a: LineResult[], b: LineResult[]): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    if (a[i].numeric !== b[i].numeric) return false;
    if ((a[i].stringValue ?? '') !== (b[i].stringValue ?? '')) return false;
    if ((a[i].error ?? '') !== (b[i].error ?? '')) return false;
  }
  return true;
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
    const name = asn[1].toLowerCase();
    const exprText = asn[2];

    // Reserved-name guards. These render specially in the result gutter as
    // a clickable-looking "N/A" / "Excel Formula" with a hover tooltip.
    if (isReservedXName(name)) {
      return {
        index, kind: 'assignment', raw, varName: name,
        display: '',
        error: 'N/A',
        errorKind: 'reserved-x',
        errorTooltip: X_RESERVED_TOOLTIP
      };
    }
    if (isExcelFunctionName(name)) {
      return {
        index, kind: 'assignment', raw, varName: name,
        display: '',
        error: 'Excel Formula',
        errorKind: 'reserved-excel',
        errorTooltip: EXCEL_FORMULA_TOOLTIP
      };
    }
    if (isReservedName(name)) {
      // Reserved names like L1, pi, e, true, false, null can't be assigned.
      // Show a clear error rather than silently falling through to expression
      // evaluation (which would convert `=` to `==` and return 0).
      return {
        index, kind: 'assignment', raw, varName: name,
        display: '',
        error: 'N/A',
        errorKind: 'reserved-name',
        errorTooltip: RESERVED_NAME_TOOLTIP
      };
    }
    if (ctx.definedInThisPass?.has(name)) {
      return {
        index, kind: 'assignment', raw, varName: name,
        display: '',
        error: 'Duplicate',
        errorKind: 'duplicate-var',
        errorTooltip: DUPLICATE_VAR_TOOLTIP
      };
    }
    ctx.definedInThisPass?.add(name);
    const evaluated = computeExpression(exprText, index, ctx);
    if (evaluated.error) {
      return {
        index, kind: 'assignment', raw, varName: name,
        error: evaluated.error,
        errorKind: evaluated.errorKind,
        errorTooltip: evaluated.errorTooltip,
        display: ''
      };
    }
    if (evaluated.stringValue !== undefined) {
      return {
        index, kind: 'assignment', raw, varName: name,
        stringValue: evaluated.stringValue,
        display: evaluated.stringValue
      };
    }
    // NaN (e.g. AVERAGE over an empty range) shows as no result rather than
    // the literal string "NaN", and is NOT bound into scope.
    const v = evaluated.value;
    if (v === undefined || isNaN(v)) {
      return { index, kind: 'assignment', raw, varName: name, display: '' };
    }
    return {
      index,
      kind: 'assignment',
      raw,
      varName: name,
      numeric: v,
      display: formatResult(v, ctx.decimals)
    };
  }

  return tryEvalExpression(raw, index, ctx);
}

function tryEvalExpression(raw: string, index: number, ctx: PreprocessCtx): LineResult {
  // Strip currency / x / commas first when sniffing for math-like tokens, so
  // a line like "$50 x 2" is recognized.
  const sniff = raw.replace(CURRENCY_RE, '');
  const looksLikeMath = /(?<![A-Za-z_])[0-9]|[+\-*/^()]|%|\bL\d+\b/.test(sniff) ||
    hasKnownIdentifier(sniff, ctx) ||
    hasExcelCall(sniff);
  if (!looksLikeMath) {
    return { index, kind: 'text', raw, display: '' };
  }
  const result = computeExpression(raw, index, ctx);
  if (result.error) {
    return {
      index, kind: 'expression', raw,
      error: result.error,
      errorKind: result.errorKind,
      errorTooltip: result.errorTooltip,
      display: ''
    };
  }
  if (result.stringValue !== undefined) {
    return {
      index, kind: 'expression', raw,
      stringValue: result.stringValue,
      display: result.stringValue
    };
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
  return ids.some(id => Object.prototype.hasOwnProperty.call(ctx.scope, id.toLowerCase()));
}

function hasExcelCall(raw: string): boolean {
  // Excel-style functions only count when followed by `(`.
  const re = /([A-Za-z_][A-Za-z0-9_]*)\s*\(/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(raw)) !== null) {
    if (isExcelFunctionName(m[1])) return true;
  }
  return false;
}

function isReservedName(name: string): boolean {
  // name is already lowercased at the call site
  return /^l\d+$/.test(name) || ['pi', 'e', 'true', 'false', 'null'].includes(name);
}

function computeExpression(
  exprText: string,
  index: number,
  ctx: PreprocessCtx
): {
  value?: number;
  stringValue?: string;
  error?: string;
  errorKind?: LineResult['errorKind'];
  errorTooltip?: string;
} {
  try {
    let s = exprText;

    // 0a. Strip an Excel-style leading `=` prefix so `=SUM(L1:L4)` is treated
    //     identically to `SUM(L1:L4)`. The line-level `name = expr`
    //     assignment was already split out by ASSIGN_RE before this is
    //     reached, so any leading `=` here is the formula prefix from Excel
    //     muscle memory, not an assignment operator.
    s = s.replace(/^\s*=\s*(?=\S)/, '');

    // 0. Strip currency symbols entirely. They contribute no value.
    s = s.replace(CURRENCY_RE, '');

    // 1. Strip number commas (1,000 -> 1000).
    s = stripNumberCommas(s);

    // 2. Replace L<a>:L<b> RANGES with the comma-separated values of each
    //    line that has a numeric result. Lines that are empty / blank /
    //    erroring are SKIPPED (matches Excel: SUM over a range with empty
    //    cells just sums the populated ones). If the entire range is empty,
    //    we emit `[]` so the (overridden) sum / mean / count helpers see an
    //    empty array — sum -> 0, count -> 0, mean -> NaN.
    s = s.replace(LINE_RANGE_RE, (_m, a, b) => {
      let lo = Number(a);
      let hi = Number(b);
      if (lo > hi) [lo, hi] = [hi, lo];
      const values: string[] = [];
      for (let n = lo; n <= hi; n++) {
        const refIdx = n - 1;
        const ref = (refIdx < ctx.results.length ? ctx.results[refIdx] : undefined) ??
                    (ctx.snapshot && refIdx < ctx.snapshot.length ? ctx.snapshot[refIdx] : undefined);
        if (ref && ref.numeric !== undefined && isFinite(ref.numeric)) {
          values.push(`(${ref.numeric})`);
        }
      }
      return values.length ? values.join(',') : '[]';
    });

    // 3. Replace line refs L<n>. Backward refs come from the current pass's
    //    accumulated results; forward refs fall back to the prior-pass
    //    snapshot so a line can reference one below it (resolves on iter 2+).
    s = s.replace(LINE_REF_RE, (_m, n) => {
      const refIdx = Number(n) - 1;
      if (refIdx < 0) return 'NaN';
      if (refIdx < ctx.results.length) {
        const ref = ctx.results[refIdx];
        if (ref && ref.numeric !== undefined && isFinite(ref.numeric)) {
          return `(${ref.numeric})`;
        }
      }
      if (ctx.snapshot && refIdx < ctx.snapshot.length) {
        const ref = ctx.snapshot[refIdx];
        if (ref && ref.numeric !== undefined && isFinite(ref.numeric)) {
          return `(${ref.numeric})`;
        }
      }
      return 'NaN';
    });

    // 4. Replace standalone `x` (not part of an identifier) with `*`. Done
    //    before percentages/suffixes so "5x10%" -> "5*10%".
    s = s.replace(/(^|[^A-Za-z0-9_])x(?![A-Za-z0-9_])/gi, (_m, lead) => `${lead}*`);

    // 4b. Normalise Excel function names to lower-case so user input like
    //     SUM(...) reaches mathjs as sum(...). Only when followed by `(`.
    s = s.replace(/\b([A-Za-z_][A-Za-z0-9_]*)\s*(?=\()/g, (m, name: string) => {
      return isExcelFunctionName(name) ? m.replace(name, name.toLowerCase()) : m;
    });

    // 4c. Bare TRUE / FALSE (uppercase, e.g. inside IF()) are treated as the
    //     string literals "TRUE" / "FALSE" so the user doesn't have to type
    //     quotes inside an IF call. Lowercase true/false stays as the mathjs
    //     boolean literals — different intent.
    s = s.replace(/\b(TRUE|FALSE)\b/g, (_m, w: string) => `"${w}"`);

    // 4d. A single `=` inside an expression means equality (Excel / English
    //     usage). mathjs uses `==` for that. Don't touch `==`, `!=`, `<=`,
    //     `>=`. The line-level `name = expr` assignment was already split out
    //     by ASSIGN_RE before we got here, so any remaining `=` is a
    //     comparison operator.
    s = s.replace(/(?<![=<>!])=(?!=)/g, '==');

    // 4e. Normalise all remaining identifiers to lower-case so variable names
    //     are case-insensitive (D0 and d0, Required and required, etc.).
    //     By this point: L<n> refs are replaced (step 3), x/X is `*` (step 4),
    //     Excel names are lowercased (step 4b), TRUE/FALSE are quoted (step 4c).
    //     We skip content inside double-quoted strings to preserve "TRUE"/"FALSE".
    s = s.replace(/"[^"]*"|\b([A-Za-z_][A-Za-z0-9_]*)\b/g, (m, ident) =>
      ident !== undefined ? m.toLowerCase() : m);

    // 5. Apply percentage / bps preprocessor.
    s = preprocessPercentages(s);

    // 6. Replace user suffixes (e.g. `5m` -> `(5*1000000)`). Editor usually
    //    expands these eagerly into commas; this is a fallback for unexpanded
    //    text or when expandSuffixesInEditor is off.
    s = applySuffixes(s, ctx.suffixes);

    if (s.trim() === '') return {};
    const value = math.evaluate(s, { ...ctx.scope });
    if (typeof value === 'number') return { value };
    if (typeof value === 'string') return { stringValue: value };
    if (typeof value === 'boolean') return { value: value ? 1 : 0 };
    if (typeof value === 'object' && value !== null && 'valueOf' in value) {
      const v = (value as { valueOf(): unknown }).valueOf();
      if (typeof v === 'number') return { value: v };
      if (typeof v === 'string') return { stringValue: v };
      if (typeof v === 'boolean') return { value: v ? 1 : 0 };
    }
    return {};
  } catch (e: unknown) {
    const msg = e instanceof Error ? e.message : String(e);
    // mathjs throws "Undefined symbol X" when the user wrote a bare word that
    // isn't a number, function, or known variable. The most common cause is
    // forgetting to put quotes around a text value (e.g. `IF(cond, YES, NO)`
    // instead of `IF(cond, "YES", "NO")`). Show a friendlier "err" with a
    // hover tooltip that explains the fix. TRUE/FALSE are exempt from this
    // (rewritten to "TRUE"/"FALSE" earlier in this function).
    if (/Undefined symbol/i.test(msg)) {
      return {
        error: 'err',
        errorKind: 'unquoted-string',
        errorTooltip: UNQUOTED_STRING_TOOLTIP
      };
    }
    return { error: shortError(msg) };
  }
}

function preprocessPercentages(s: string): string {
  // bps and bp first (longer suffix), with leading +/- as relative to LHS.
  // The LHS-relative form requires a value-producing token before the +/- (a
  // digit, identifier, or closing paren/bracket). Without that guard, a
  // standalone `-10%` at the start of a line would expand to `*(1-10/100)`
  // (a syntax error) instead of evaluating to `-0.1`.
  s = s.replace(/([\w)\]]\s*)([+\-])\s*([0-9]*\.?[0-9]+)\s*bps?\b/gi, (_m, lhs, op, n) =>
    op === '+' ? `${lhs}*(1+${n}/10000)` : `${lhs}*(1-${n}/10000)`);
  s = s.replace(/([0-9]*\.?[0-9]+)\s*bps?\b/gi, (_m, n) => `(${n}/10000)`);

  // Percentages: relative-to-LHS for +/-, otherwise literal /100.
  s = s.replace(/([\w)\]]\s*)([+\-])\s*([0-9]*\.?[0-9]+)\s*%/g, (_m, lhs, op, n) =>
    op === '+' ? `${lhs}*(1+${n}/100)` : `${lhs}*(1-${n}/100)`);
  s = s.replace(/([0-9]*\.?[0-9]+)\s*%/g, (_m, n) => `(${n}/100)`);
  return s;
}

function applySuffixes(s: string, suffixes: Suffix[]): string {
  if (!suffixes.length) return s;
  // Sort longest-symbol-first to avoid `mb` being matched as `m` then `b`.
  const sorted = [...suffixes].sort((a, b) => b.symbol.length - a.symbol.length);
  for (const suf of sorted) {
    const flags = suf.caseSensitive ? 'g' : 'gi';
    // Match a number directly followed by the suffix, with right-side word
    // boundary. Also exclude `.` so a malformed input like `5k.5` doesn't get
    // partially expanded to `5000.5` — we leave it alone instead.
    const escaped = escapeRegex(suf.symbol);
    const re = new RegExp(`([0-9]*\\.?[0-9]+)${escaped}(?![A-Za-z0-9_.])`, flags);
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

// Evaluate an arbitrary sub-expression (e.g. text the user dragged over within
// a line) using the scope visible at `lineIndex`. Returns the numeric result or
// undefined if the text cannot be evaluated or is not numeric.
export function evaluateSelectedText(
  selectedText: string,
  lineResults: LineResult[],
  lineIndex: number,
  suffixes: Suffix[],
  decimals: number
): number | undefined {
  const scope: Record<string, number> = {};
  for (let i = 0; i <= lineIndex; i++) {
    const r = lineResults[i];
    if (r && r.varName !== undefined && r.numeric !== undefined && isFinite(r.numeric)) {
      scope[r.varName] = r.numeric;
    }
  }
  const ctx: PreprocessCtx = {
    scope,
    results: lineResults.slice(0, lineIndex + 1),
    suffixes,
    previous: [],
    decimals
  };
  const result = computeExpression(selectedText, lineIndex, ctx);
  if (result.error || result.value === undefined || isNaN(result.value)) return undefined;
  return result.value;
}
