import { evaluateNote, LineResult } from './evaluator';
import { highlightNote } from './highlighter';
import { formatWithCommas } from './formatter';
import type { Mode, Settings, Suffix } from '../shared/types';

const editor = document.getElementById('editor') as HTMLTextAreaElement;
const overlay = document.getElementById('syntax-overlay') as HTMLPreElement;
const measure = document.getElementById('measure') as HTMLDivElement;
const lineGutter = document.getElementById('line-gutter') as HTMLDivElement;
const resultGutter = document.getElementById('result-gutter') as HTMLDivElement;
const status = document.getElementById('status-msg') as HTMLSpanElement;
const closeBtn = document.getElementById('close-window') as HTMLButtonElement;
const settingsBtn = document.getElementById('open-settings') as HTMLButtonElement;
const copyMdBtn = document.getElementById('copy-md') as HTMLButtonElement;
const modeMathBtn = document.getElementById('mode-math') as HTMLButtonElement;
const modeTextBtn = document.getElementById('mode-text') as HTMLButtonElement;

let settings: Settings;
let lastResults: LineResult[] = [];
let saveTimer: number | null = null;

async function init() {
  settings = await window.mathPopup.getSettings();
  editor.value = settings.noteContent ?? '';
  applyMode(settings.mode);
  bindEvents();
  render();
  editor.focus();
}

function bindEvents() {
  editor.addEventListener('input', onInput);
  editor.addEventListener('scroll', syncScroll);
  editor.addEventListener('keydown', onKeyDown);
  window.addEventListener('resize', () => render());

  closeBtn.addEventListener('click', () => window.mathPopup.hidePopup());
  settingsBtn.addEventListener('click', () => window.mathPopup.openSettings());
  copyMdBtn.addEventListener('click', copyAsMarkdown);

  modeMathBtn.addEventListener('click', () => setMode('math'));
  modeTextBtn.addEventListener('click', () => setMode('text'));

  // Listen for settings changes pushed via a polling refresh-on-focus.
  window.addEventListener('focus', async () => {
    settings = await window.mathPopup.getSettings();
    applyMode(settings.mode);
    render();
  });
}

function setMode(mode: Mode) {
  if (settings.mode === mode) return;
  settings.mode = mode;
  applyMode(mode);
  window.mathPopup.setSettings({ mode });
  render();
  editor.focus();
}

function applyMode(mode: Mode) {
  document.body.classList.toggle('text-mode', mode === 'text');
  modeMathBtn.classList.toggle('active', mode === 'math');
  modeMathBtn.setAttribute('aria-selected', mode === 'math' ? 'true' : 'false');
  modeTextBtn.classList.toggle('active', mode === 'text');
  modeTextBtn.setAttribute('aria-selected', mode === 'text' ? 'true' : 'false');
}

function onInput() {
  scheduleSave();
  render();
  ensureCaretLineVisible();
}

// Browsers only auto-scroll a textarea enough to make the caret pixel visible,
// not the whole line. That leaves the bottom of the cursor's line (and its
// gutter row) clipped when typing near the bottom of the editor. Scroll the
// editor so the entire line containing the caret is in view.
function ensureCaretLineVisible() {
  if (editor.selectionStart !== editor.selectionEnd) return;
  const caret = editor.selectionStart;
  const text = editor.value;
  const lineIndex = (text.slice(0, caret).match(/\n/g) || []).length;

  const editorStyle = getComputedStyle(editor);
  const padTop = parseFloat(editorStyle.paddingTop) || 0;
  const lineHeight = parseFloat(editorStyle.lineHeight) || 22;

  // Approximate line position (no wrap). Sufficient for short note lines.
  const lineTop = padTop + lineIndex * lineHeight;
  const lineBot = lineTop + lineHeight;

  const viewTop = editor.scrollTop;
  const viewBot = viewTop + editor.clientHeight;

  if (lineBot > viewBot) {
    editor.scrollTop = lineBot - editor.clientHeight;
  } else if (lineTop < viewTop) {
    editor.scrollTop = lineTop;
  }
}

function scheduleSave() {
  if (saveTimer) window.clearTimeout(saveTimer);
  saveTimer = window.setTimeout(() => {
    window.mathPopup.setSettings({ noteContent: editor.value });
  }, 250);
}

function onKeyDown(e: KeyboardEvent) {
  // Smart Tab: when the caret sits inside a number, jump past the number to
  // the trailing space (inserting one if missing) and re-run auto-format so
  // any commas are corrected. Math mode only.
  if (settings.mode === 'math' && e.key === 'Tab' &&
      !e.ctrlKey && !e.shiftKey && !e.altKey && !e.metaKey) {
    e.preventDefault();
    handleSmartTab();
    return;
  }

  // Auto-format / suffix-expansion triggers: space, operators, Enter, comma
  if (settings.mode === 'math' && shouldAutoFormatOnKey(e)) {
    queueMicrotask(() => maybeAutoFormat(e.key));
  }

  // Ctrl+Shift+C: copy current line's result (math mode only)
  if (settings.mode === 'math' && e.ctrlKey && e.shiftKey && (e.key === 'C' || e.key === 'c')) {
    e.preventDefault();
    copyCurrentLineResult();
  }
  // Ctrl+Shift+M: copy whole note as markdown with results inline
  if (e.ctrlKey && e.shiftKey && (e.key === 'M' || e.key === 'm')) {
    e.preventDefault();
    copyAsMarkdown();
  }
  // Esc: hide window
  if (e.key === 'Escape') {
    window.mathPopup.hidePopup();
  }
}

function shouldAutoFormatOnKey(e: KeyboardEvent): boolean {
  if (e.ctrlKey || e.metaKey || e.altKey) return false;
  return e.key === ' ' || e.key === 'Enter' || /[+\-*/^=()]/.test(e.key);
}

// ---- smart Tab ----
// If the caret sits inside (or at either edge of) a number, jump to just past
// the number, ensure there's a trailing space, then run the auto-format pass.
// Returns true if it handled the keystroke.
function handleSmartTab(): boolean {
  const caret = editor.selectionStart;
  if (caret !== editor.selectionEnd) return false;
  const text = editor.value;
  const isNumChar = (c: string | undefined) => c !== undefined && /[\d,.]/.test(c);

  let start = caret;
  while (start > 0 && isNumChar(text[start - 1])) start--;
  let end = caret;
  while (end < text.length && isNumChar(text[end])) end++;

  // Must contain at least one digit (commas / dots alone don't count).
  if (start === end || !/\d/.test(text.slice(start, end))) return false;

  let newText: string;
  let newCaret: number;
  if (text[end] === ' ') {
    newText = text;
    newCaret = end + 1;
  } else {
    newText = text.slice(0, end) + ' ' + text.slice(end);
    newCaret = end + 1;
  }
  editor.value = newText;
  editor.selectionStart = editor.selectionEnd = newCaret;
  // Re-use the regular auto-format pipeline: it formats the line up to the
  // caret (now sitting after the inserted space) which will recomma the number.
  maybeAutoFormat(' ');
  // maybeAutoFormat early-returns (no render) when no formatting was needed,
  // but we still inserted a space — ensure overlay + save catch up.
  scheduleSave();
  render();
  return true;
}

// ---- auto-format current line ----
// Strategy: split the line at the caret, format the LEFT half only, and
// re-join. The trigger character (space / operator / Enter) was just typed
// and lives at the end of the LEFT half — keeping format scoped to the left
// guarantees the trigger character lands right after the formatted text and
// the caret follows it.
function maybeAutoFormat(_triggerKey: string) {
  if (!settings.autoFormatNumbers && !settings.expandSuffixesInEditor) return;
  const caret = editor.selectionStart;
  if (caret !== editor.selectionEnd) return; // skip when there's a selection
  const text = editor.value;

  const lineStart = text.lastIndexOf('\n', caret - 1) + 1;
  let lineEnd = text.indexOf('\n', caret);
  if (lineEnd === -1) lineEnd = text.length;

  const before = text.slice(lineStart, caret);
  const after = text.slice(caret, lineEnd);

  const formattedBefore = formatLineForEditor(before, settings.suffixes, {
    autoFormatNumbers: settings.autoFormatNumbers,
    expandSuffixes: settings.expandSuffixesInEditor
  }).text;

  if (formattedBefore === before) return;

  const newLine = formattedBefore + after;
  const newCaretAbs = lineStart + formattedBefore.length;

  const head = text.slice(0, lineStart);
  const tail = text.slice(lineEnd);
  editor.value = head + newLine + tail;
  editor.selectionStart = editor.selectionEnd = newCaretAbs;
  scheduleSave();
  render();
}

interface FmtOpts { autoFormatNumbers: boolean; expandSuffixes: boolean; }

export function formatLineForEditor(line: string, suffixes: Suffix[], opts: FmtOpts): { text: string } {
  // Don't format markdown header lines.
  if (/^\s*#{1,6}\s+/.test(line)) return { text: line };

  let out = line;

  // 0. Bare decimals get a leading zero: ".123" -> "0.123" when preceded by
  //    whitespace, an operator, opening paren, or the start of the line.
  out = out.replace(/(^|[\s+\-*/^=(,])\.(\d)/g, (_m, lead, d) => `${lead}0.${d}`);

  // 1. Expand custom suffixes that come right after a number: "1m" -> "1000000".
  if (opts.expandSuffixes && suffixes.length) {
    const sorted = [...suffixes].sort((a, b) => b.symbol.length - a.symbol.length);
    for (const suf of sorted) {
      const flags = suf.caseSensitive ? 'g' : 'gi';
      const escaped = suf.symbol.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      // Only expand when the suffix is a STANDALONE token (no letter or digit
      // immediately after) so we don't munge identifiers like "max".
      const re = new RegExp(`(^|[^A-Za-z0-9_])([0-9][0-9,]*(?:\\.[0-9]+)?)${escaped}(?![A-Za-z0-9_])`, flags);
      out = out.replace(re, (_m, lead, num) => {
        const cleaned = num.replace(/,/g, '');
        const value = Number(cleaned) * suf.multiplier;
        if (!isFinite(value)) return _m;
        return `${lead}${formatNumberForEditor(value, opts.autoFormatNumbers)}`;
      });
    }
  }

  // 2. Re-comma-ize bare integers and decimals (4+ digits in integer part).
  if (opts.autoFormatNumbers) {
    out = out.replace(/(^|[^A-Za-z0-9_,.])(-?\d{4,}(?:\.\d+)?)(?![\d.])/g, (_m, lead, num) => {
      return `${lead}${commifyNumber(num)}`;
    });
    // Re-comma-ize numbers that already contain commas in case they were edited.
    out = out.replace(/(^|[^A-Za-z0-9_.])(-?\d{1,3}(?:,\d{3})*\d*(?:\.\d+)?)(?![\d.])/g,
      (_m, lead, num) => `${lead}${commifyNumber(num.replace(/,/g, ''))}`);
  }

  return { text: out };
}

function commifyNumber(numStr: string): string {
  const negative = numStr.startsWith('-');
  const body = negative ? numStr.slice(1) : numStr;
  const [intPart, decPart] = body.split('.');
  if (!/^\d+$/.test(intPart)) return numStr;
  const withCommas = intPart.length >= 4 ? formatWithCommas(intPart) : intPart;
  const out = decPart !== undefined ? `${withCommas}.${decPart}` : withCommas;
  return negative ? `-${out}` : out;
}

function formatNumberForEditor(n: number, useCommas: boolean): string {
  if (Number.isInteger(n) && Math.abs(n) < 1e21) {
    return useCommas ? formatWithCommas(n.toString()) : n.toString();
  }
  const s = n.toString();
  if (!useCommas) return s;
  const [intPart, decPart] = s.split('.');
  return decPart !== undefined ? `${formatWithCommas(intPart)}.${decPart}` : formatWithCommas(intPart);
}

// ---- render pipeline ----
function render() {
  if (settings.mode === 'math') {
    // Pass the previous render's results so the evaluator can carry over
    // last-good values for lines that are temporarily mid-edit.
    lastResults = evaluateNote(editor.value, settings.suffixes, lastResults, settings.decimals);
  } else {
    // Text mode: no evaluation, no results.
    lastResults = editor.value.split('\n').map((raw, i) => ({
      index: i,
      kind: 'text' as const,
      raw,
      display: ''
    }));
  }
  overlay.innerHTML = highlightNote(editor.value, lastResults, settings.mode);
  syncScroll();
  layoutGutters();
  updateStatus();
}

function syncScroll() {
  overlay.scrollTop = editor.scrollTop;
  overlay.scrollLeft = editor.scrollLeft;
  // Keep gutters in vertical sync with the editor's scroll.
  lineGutter.scrollTop = editor.scrollTop;
  resultGutter.scrollTop = editor.scrollTop;
}

function layoutGutters() {
  const lines = editor.value.split('\n');
  // (heights are read from the overlay's rendered children below)
  const editorStyle = getComputedStyle(editor);

  const lineHeight = parseFloat(editorStyle.lineHeight) || 22;
  const heights: number[] = [];

  // Read per-line heights from the overlay's rendered children. The overlay
  // shares font/padding/width/wrap rules with the editor, so each .ov-line's
  // offsetHeight matches the textarea's visual line height (including wrap).
  const overlayLines = overlay.querySelectorAll<HTMLElement>('.ov-line');
  for (let i = 0; i < lines.length; i++) {
    const text = lines[i].length === 0 ? ' ' : lines[i];
    const el = overlayLines[i];
    const h = el ? el.offsetHeight : lineHeight;
    heights.push(Math.max(lineHeight, h));
  }

  // Build line-number column
  const padTop = parseFloat(editorStyle.paddingTop) || 0;
  const padBot = parseFloat(editorStyle.paddingBottom) || 0;
  lineGutter.style.paddingTop = padTop + 'px';
  lineGutter.style.paddingBottom = padBot + 'px';
  resultGutter.style.paddingTop = padTop + 'px';
  resultGutter.style.paddingBottom = padBot + 'px';

  // Size the line-number column to fit the widest label (e.g. "L1234").
  // Measure the actual rendered width of the largest label. Use individual
  // font properties (not the `font` shorthand) because Chromium returns an
  // empty string for getComputedStyle().font when the font is set via
  // individual longhand properties.
  const probe = document.createElement('span');
  probe.style.position = 'absolute';
  probe.style.visibility = 'hidden';
  probe.style.whiteSpace = 'pre';
  probe.style.fontFamily = editorStyle.fontFamily;
  probe.style.fontSize = editorStyle.fontSize;
  probe.style.fontWeight = editorStyle.fontWeight;
  probe.style.letterSpacing = editorStyle.letterSpacing;
  probe.textContent = `L${lines.length}`;
  document.body.appendChild(probe);
  const labelWidth = probe.offsetWidth;
  probe.remove();
  const gutterStyle = getComputedStyle(lineGutter);
  const gutterPadX = (parseFloat(gutterStyle.paddingLeft) || 0) +
                     (parseFloat(gutterStyle.paddingRight) || 0);
  lineGutter.style.minWidth = Math.ceil(labelWidth + gutterPadX + 6) + 'px';

  lineGutter.innerHTML = heights
    .map((h, i) => `<div class="row" style="height:${h}px">L${i + 1}</div>`)
    .join('');

  // Build result column
  resultGutter.innerHTML = heights
    .map((h, i) => {
      const r = lastResults[i];
      if (!r) return `<div class="row empty" style="height:${h}px"></div>`;
      if (r.error) return `<div class="row error" style="height:${h}px" title="${escapeAttr(r.error)}">err</div>`;
      const txt = r.display ?? '';
      let cls = 'row';
      if (txt === '') cls = 'row empty';
      else if (r.stale) cls = 'row stale';
      return `<div class="${cls}" style="height:${h}px">${escapeHtml(txt)}</div>`;
    })
    .join('');
}

function updateStatus() {
  if (settings.mode === 'text') {
    status.textContent = 'Text mode';
    status.className = 'status-msg';
    return;
  }
  const errs = lastResults.filter(r => r.error).length;
  if (errs === 0) {
    status.textContent = 'Ready';
    status.className = 'status-msg ok';
  } else {
    status.textContent = `${errs} error${errs > 1 ? 's' : ''}`;
    status.className = 'status-msg err';
  }
}

// ---- copy actions ----
function copyCurrentLineResult() {
  const caret = editor.selectionStart;
  const text = editor.value;
  const lineStart = text.lastIndexOf('\n', caret - 1) + 1;
  const lineNumber = text.slice(0, lineStart).split('\n').length - 1; // 0-based
  const r = lastResults[lineNumber];
  if (!r || r.numeric === undefined) {
    flashStatus('No result on this line', true);
    return;
  }
  window.mathPopup.copyText(String(r.numeric));
  flashStatus(`Copied ${r.display ?? r.numeric}`);
}

function copyAsMarkdown() {
  const lines = editor.value.split('\n');
  const out: string[] = [];
  for (let i = 0; i < lines.length; i++) {
    const r = lastResults[i];
    const raw = lines[i];
    if (!r || r.display === undefined || r.display === '' || r.kind === 'header' || r.kind === 'blank' || r.kind === 'text') {
      out.push(raw);
      continue;
    }
    if (r.error) {
      out.push(`${raw}    \`error: ${r.error}\``);
      continue;
    }
    out.push(`${raw}    \`= ${r.display}\``);
  }
  window.mathPopup.copyText(out.join('\n'));
  flashStatus('Copied as markdown');
}

let flashTimer: number | null = null;
function flashStatus(msg: string, isErr = false) {
  status.textContent = msg;
  status.className = isErr ? 'status-msg err' : 'status-msg ok';
  if (flashTimer) window.clearTimeout(flashTimer);
  flashTimer = window.setTimeout(updateStatus, 1400);
}

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function escapeAttr(s: string): string {
  return escapeHtml(s).replace(/"/g, '&quot;');
}

init();
