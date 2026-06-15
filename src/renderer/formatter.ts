// Number formatting helpers — formatting for the result gutter and for the
// in-editor "comma-ize" auto-format on space.

// `maxDecimals` is the cap, not a fixed length:
//   - integer result -> no decimal point shown
//   - 0.5 with cap=2 -> "0.5"
//   - 100/3 with cap=2 -> "33.33"
//   - 0.0001 with cap=2 -> "0" (rounded to zero is acceptable; bump the cap)
export function formatResult(n: number, maxDecimals: number): string {
  if (!isFinite(n)) return n.toString();
  if (Math.abs(n) >= 1e21) return n.toString();
  const cap = Math.max(0, Math.floor(maxDecimals));
  // Round to the cap, then strip trailing zeros AFTER the decimal point only
  // (and the dot itself if all fractional digits get stripped).
  const rounded = n.toFixed(cap);
  const trimmed = cap > 0
    ? rounded.replace(/(\.\d*?)0+$/, '$1').replace(/\.$/, '')
    : rounded;
  const [intPart, decPart] = trimmed.split('.');
  const withCommas = formatWithCommas(intPart);
  return decPart !== undefined ? `${withCommas}.${decPart}` : withCommas;
}

export function formatWithCommas(intStr: string): string {
  const negative = intStr.startsWith('-');
  const digits = negative ? intStr.slice(1) : intStr;
  if (!/^\d+$/.test(digits)) return intStr;
  const out = digits.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return negative ? `-${out}` : out;
}

// Strip thousands-separator commas so they're purely visual — their position in
// a number doesn't matter (1,234, 10,000,99, 10,000, all collapse to plain
// digits). Paren-aware so function arguments survive:
//   - At the TOP level a comma right after a digit can only be a (visual)
//     thousands separator, so it's always dropped — wherever it sits.
//   - INSIDE a function call's parens a comma after a digit might separate
//     arguments, so we only drop it when it's clearly a thousands group
//     (followed by 3+ digits). That keeps min(1,2), round(3.14,2), min(5,-3),
//     while still collapsing sum(1,000,000, 2,000) -> sum(1000000, 2000).
// Used by the evaluator preprocessor.
export function stripNumberCommas(s: string): string {
  let out = '';
  let depth = 0;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '(') depth++;
    else if (ch === ')') { if (depth > 0) depth--; }
    if (ch === ',' && i > 0 && /\d/.test(s[i - 1])) {
      if (depth === 0) continue;                    // top level: always visual → drop
      if (/^\d{3}/.test(s.slice(i + 1))) continue;  // inside (): thousands group → drop
      // otherwise inside (): argument separator → keep
    }
    out += ch;
  }
  return out;
}
