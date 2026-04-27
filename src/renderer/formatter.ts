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

// Strip commas inside number-like sequences. Used by the evaluator preprocessor.
export function stripNumberCommas(s: string): string {
  return s.replace(/(\d),(?=\d{3}(?!\d))/g, '$1');
}
