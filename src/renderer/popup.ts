import { evaluateNote, evaluateSelectedText, LineResult, EXCEL_FORMULA_TOOLTIP, X_RESERVED_TOOLTIP, UNQUOTED_STRING_TOOLTIP, RESERVED_NAME_TOOLTIP, DUPLICATE_VAR_TOOLTIP, isExcelFunctionName } from './evaluator';
import { highlightNote, listIndentCols, wrapInlineMarkers, ActiveToken } from './highlighter';
import { formatWithCommas, formatResult } from './formatter';
import type { Mode, Page, Settings, Suffix, ThemePref } from '../shared/types';
import { ZOOM_MIN, ZOOM_MAX, ZOOM_STEP } from '../shared/types';

const editor = document.getElementById('editor') as HTMLTextAreaElement;
const overlay = document.getElementById('syntax-overlay') as HTMLPreElement;
const measure = document.getElementById('measure') as HTMLDivElement;
const lineGutter = document.getElementById('line-gutter') as HTMLDivElement;
const resultGutter = document.getElementById('result-gutter') as HTMLDivElement;
const resultOverlay = document.getElementById('result-overlay') as HTMLDivElement;
const status = document.getElementById('status-msg') as HTMLSpanElement;
const closeBtn = document.getElementById('close-window') as HTMLButtonElement;
const settingsBtn = document.getElementById('open-settings') as HTMLButtonElement;
const pinBtn = document.getElementById('toggle-pin') as HTMLButtonElement;
const helpBtn = document.getElementById('open-help') as HTMLButtonElement;
const varsBtn = document.getElementById('show-vars') as HTMLButtonElement;
const varsPopup = document.getElementById('vars-popup') as HTMLDivElement;
const pageIndicator = document.getElementById('page-indicator') as HTMLSpanElement;
const zoomIndicator = document.getElementById('zoom-indicator') as HTMLSpanElement;
const cmdMenu = document.getElementById('cmd-menu') as HTMLDivElement;
const hoverTooltip = document.getElementById('hover-tooltip') as HTMLDivElement;
const findLayer = document.getElementById('find-layer') as HTMLPreElement;
const findBar = document.getElementById('find-bar') as HTMLDivElement;
const findInput = document.getElementById('find-input') as HTMLInputElement;
const findCount = document.getElementById('find-count') as HTMLSpanElement;
const findPrevBtn = document.getElementById('find-prev') as HTMLButtonElement;
const findNextBtn = document.getElementById('find-next') as HTMLButtonElement;
const findCloseBtn = document.getElementById('find-close') as HTMLButtonElement;

// ============================================================
// Contenteditable editor engine (per-line blocks)
// ------------------------------------------------------------
// The editor is a contenteditable <div> whose children are one ".ed-line" block
// per logical line. Each block carries its OWN right padding (math lines reserve
// the answer column; text lines use the full width) — that per-line padding is
// how each line wraps independently, which a single <textarea> can't do. We keep
// a plain string as the source of truth (edValue) and expose a textarea-shaped
// API (value, selectionStart/End, setSelectionRange, setRangeText) on the div so
// the rest of the app — undo, find, the slash menu, list logic, ~170 call sites —
// keeps working unchanged.
// ============================================================
const ed = editor as unknown as HTMLDivElement;
let edValue = '';
let edCaretStart = 0;
let edCaretEnd = 0;
let edComposing = false;
// Scroll position captured just BEFORE the browser applies a native edit. The
// browser auto-scrolls a focused contenteditable to keep the caret visible on
// every edit (notably Backspace near the bottom), which reads as a jump. We
// restore this in onEditorInput and let ensureCaretLineVisible make the only
// deliberate scroll adjustment.
let scrollBeforeInput = 0;

function edFirstTextNode(el: HTMLElement): Text | null {
  for (const c of Array.from(el.childNodes)) if (c.nodeType === Node.TEXT_NODE) return c as Text;
  return null;
}

// Build the editor DOM: one .ed-line per line. The per-line math/text class is
// applied by applyEditorLineModes() once lineModes are known.
function buildEditorDOM(text: string) {
  const lines = text.split('\n');
  let html = '';
  for (let i = 0; i < lines.length; i++) {
    html += `<div class="ed-line">${lines[i].length ? wrapInlineMarkers(lines[i]) : '<br>'}</div>`;
  }
  // Replacing innerHTML resets the element's scrollTop to 0; preserve it so the
  // view doesn't jump to the top on every keystroke when the note is scrolled.
  const st = ed.scrollTop;
  ed.innerHTML = html;
  ed.scrollTop = st;
  edValue = text;
  applyEditorLineModes();
}

// Conceal inline-markdown markers (.ed-mk) on every text line except the one the
// caret sits on — the editor half of the Obsidian-style reveal. Toggles a class
// only (no DOM rebuild), so it's safe to call on a bare caret move. Mirrors the
// overlay's .ov-conceal (see highlightNote) so both layers collapse the same
// characters and stay aligned.
function applyEditorConceal(caretLine: number) {
  const divs = ed.children;
  for (let i = 0; i < divs.length; i++) {
    (divs[i] as HTMLElement).classList.toggle('ed-conceal', lineModeAt(i) === 'text' && i !== caretLine);
  }
}

// Toggle the per-line math class on existing .ed-line blocks, hang-indent list
// lines, and conceal inline markers off the caret line. Cheap and does not rebuild
// text, so it never disturbs the caret.
function applyEditorLineModes() {
  const divs = ed.children;
  const lines = editor.value.split('\n');
  const caretLine = document.activeElement === editor ? caretLineIndex() : -1;
  for (let i = 0; i < divs.length; i++) {
    const el = divs[i] as HTMLElement;
    const isMath = lineModeAt(i) === 'math';
    el.classList.toggle('ed-math', isMath);
    el.classList.toggle('ed-conceal', !isMath && i !== caretLine);
    // Mirror the overlay's hanging indent (see listIndentCols / highlightNote)
    // so a wrapped bullet/number line's continuation rows sit under the item
    // text. This MUST match .ov-line exactly or the caret drifts off the text.
    const cols = listIndentCols(lines[i] ?? '');
    if (cols > 0) {
      el.style.paddingLeft = `${cols}ch`;
      el.style.textIndent = `-${cols}ch`;
    } else if (el.style.paddingLeft || el.style.textIndent) {
      el.style.paddingLeft = '';
      el.style.textIndent = '';
    }
  }
}

// Gather a block's text and, if the selection focus is inside it, its local
// offset. Tolerates the shallow DOM the browser leaves after one native edit.
function edGatherLine(root: Node, focusNode: Node | null, focusOffset: number): { text: string; caret: number } {
  let text = '';
  let caret = -1;
  const walk = (n: Node) => {
    for (const c of Array.from(n.childNodes)) {
      if (c.nodeType === Node.TEXT_NODE) {
        if (c === focusNode) caret = text.length + Math.min(focusOffset, (c as Text).length);
        text += (c as Text).data;
      } else if (c.nodeName === 'BR') {
        if (c === focusNode) caret = text.length;
      } else if (c.nodeType === Node.ELEMENT_NODE) {
        if (c === focusNode) caret = text.length + (focusOffset === 0 ? 0 : (c.textContent ?? '').length);
        walk(c);
      }
    }
  };
  walk(root);
  return { text, caret };
}

// Read the (possibly browser-mutated) editor DOM back into a flat string + caret
// offset. Top-level blocks are the lines, joined by "\n".
function readEditorDOM(focusNode: Node | null, focusOffset: number): { text: string; caret: number } {
  const blocks = Array.from(ed.children) as HTMLElement[];
  if (blocks.length === 0) {
    const t = ed.textContent ?? '';
    return { text: t, caret: Math.min(focusOffset, t.length) };
  }
  let text = '';
  let caret = -1;
  for (let i = 0; i < blocks.length; i++) {
    if (i > 0) text += '\n';
    const lineStart = text.length;
    const block = blocks[i];
    if (block === focusNode) caret = lineStart + (focusOffset === 0 ? 0 : (block.textContent ?? '').length);
    const g = edGatherLine(block, focusNode, focusOffset);
    if (g.caret >= 0) caret = lineStart + g.caret;
    text += g.text;
  }
  if (caret < 0) caret = text.length;
  return { text, caret };
}

// Map a within-block offset to a DOM point, walking ALL of the block's text nodes
// in order. A line may now hold several text nodes (the marker spans .ed-mk used
// for conceal), not just one — so we can't assume a single first text node.
function edPointInBlock(block: HTMLElement, local: number): { node: Node; offset: number } {
  const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
  let acc = 0;
  let last: Text | null = null;
  for (let tn = walker.nextNode() as Text | null; tn; tn = walker.nextNode() as Text | null) {
    if (local <= acc + tn.length) return { node: tn, offset: local - acc };
    acc += tn.length;
    last = tn;
  }
  if (last) return { node: last, offset: last.length };
  return { node: block, offset: 0 };   // empty line (<br> placeholder)
}

// Map a flat offset to a DOM point inside the CURRENT (clean) editor structure.
function edOffsetToPoint(offset: number): { node: Node; offset: number } {
  const blocks = Array.from(ed.children) as HTMLElement[];
  let acc = 0;
  for (let i = 0; i < blocks.length; i++) {
    const len = (blocks[i].textContent ?? '').length;
    if (offset <= acc + len) return edPointInBlock(blocks[i], offset - acc);
    acc += len + 1;   // + the "\n" after this line
  }
  const last = blocks[blocks.length - 1];
  if (last) return edPointInBlock(last, (last.textContent ?? '').length);
  return { node: ed, offset: 0 };
}

function edClamp(n: number): number {
  return Math.max(0, Math.min(n, edValue.length));
}

function edSetSelection(start: number, end: number) {
  const sel = window.getSelection();
  if (!sel) return;
  const s = edOffsetToPoint(edClamp(start));
  const e = edOffsetToPoint(edClamp(end));
  const range = document.createRange();
  range.setStart(s.node, s.offset);
  try { range.setEnd(e.node, e.offset); } catch { range.setEnd(s.node, s.offset); }
  sel.removeAllRanges();
  sel.addRange(range);
  edCaretStart = edClamp(start);
  edCaretEnd = edClamp(end);
}

// Refresh the cached offsets from the live selection (after native caret moves:
// clicks, arrows, drag-selection).
function edRefreshSelectionCache() {
  const sel = window.getSelection();
  if (!sel || sel.rangeCount === 0 || !sel.anchorNode || !ed.contains(sel.anchorNode)) return;
  const a = readEditorDOM(sel.anchorNode, sel.anchorOffset).caret;
  const f = readEditorDOM(sel.focusNode, sel.focusOffset).caret;
  edCaretStart = Math.min(a, f);
  edCaretEnd = Math.max(a, f);
}

// The 'input' handler: adopt the natively-edited DOM as the new value, normalize
// it back to clean .ed-line blocks, restore the caret, then run the existing
// input pipeline (which reads editor.value / .selectionStart through the shim).
function onEditorInput() {
  if (edComposing) return;   // mid-IME: wait for compositionend
  const sel = window.getSelection();
  const fNode = sel && sel.rangeCount ? sel.focusNode : null;
  const fOff = sel ? sel.focusOffset : 0;
  const read = readEditorDOM(fNode, fOff);
  edValue = read.text;
  // Update the caret cache BEFORE syncLineModes: it reads caretLineIndex() to
  // decide which line is newly inserted (so Enter inherits the right mode). If
  // the cache still held the pre-edit caret, a new line would inherit from the
  // wrong neighbor — e.g. pressing Enter under a math line would flip the line
  // you were on back to math.
  edCaretStart = edCaretEnd = read.caret;
  syncLineModes();              // align lineModes to the new text before rebuilding
  buildEditorDOM(read.text);    // normalize to clean blocks (+ math classes)
  edSetSelection(read.caret, read.caret);
  // Undo the browser's native edit-scroll: snap back to the pre-edit position so
  // a plain edit (e.g. Backspace) doesn't shift the view. ensureCaretLineVisible
  // (run inside onInput) is then the sole authority on scrolling — it nudges only
  // when the caret line genuinely falls outside the viewport.
  editor.scrollTop = scrollBeforeInput;
  onInput();                    // existing pipeline: refs, undo, render(), menus
}

// Expose a <textarea>-compatible surface on the div so existing code is unchanged.
Object.defineProperties(ed, {
  value: {
    get() { return edValue; },
    set(v: string) { buildEditorDOM(v == null ? '' : String(v)); },
    configurable: true,
  },
  selectionStart: {
    get() { return edCaretStart; },
    set(v: number) { edSetSelection(v, Math.max(v, edCaretEnd)); },
    configurable: true,
  },
  selectionEnd: {
    get() { return edCaretEnd; },
    set(v: number) { edSetSelection(Math.min(edCaretStart, v), v); },
    configurable: true,
  },
});
(ed as any).setSelectionRange = (s: number, e: number) => edSetSelection(s, e ?? s);
(ed as any).setRangeText = (replacement: string, start?: number, end?: number, selectMode?: string) => {
  const s = start ?? edCaretStart;
  const e = end ?? edCaretEnd;
  buildEditorDOM(edValue.slice(0, s) + replacement + edValue.slice(e));
  if (selectMode === 'select') edSetSelection(s, s + replacement.length);
  else if (selectMode === 'start') edSetSelection(s, s);
  else edSetSelection(s + replacement.length, s + replacement.length);
};

let settings: Settings;
let closedPages: Page[] = [];
let pages: Page[] = [];
let activePageId: string = '';
// True immediately after a tab is closed (cleared on any edit/navigation) so
// that an immediate Ctrl+Z restores the just-closed tab instead of undoing text.
let justClosedTab = false;
// Id of the chip currently being drag-reordered, if any.
let dragSrcId: string | null = null;
let overflowHideTimer: number | null = null;
// When true, the overflow dropdown is in drag-to-reorder mode (shows all tabs).
let reorderMode = false;
let lastResults: LineResult[] = [];
// Per-line mode for the active page (parallel to editor lines) + the text
// snapshot it's aligned to. syncLineModes() in render() keeps it in step with
// every edit, so we don't have to touch each editor.value assignment site.
let lineModes: Mode[] = [];
let lineModesText = '';
let saveTimer: number | null = null;
// Pending single-click line-ref insert, held briefly so a double-click can
// pre-empt it (double-click toggles the line's math/text instead).
let pendingRefClick: { line: number; timer: number } | null = null;
// The line index whose ref-display was last rendered "raw" (caret on it), so we
// only re-render on caret moves that actually change it.
let lastCaretLine = -1;
let activeToken: ActiveToken | null = null;
const tabsBtn = document.getElementById('tabs-btn') as HTMLButtonElement;
const archiveBtn = document.getElementById('archive-btn') as HTMLButtonElement;
const tabBar = document.getElementById('tab-bar') as HTMLElement;
const tabStrip = document.getElementById('tab-strip') as HTMLElement;
const tabAddBtn = document.getElementById('tab-add-btn') as HTMLButtonElement;
const overflowBtn = document.getElementById('tab-overflow-btn') as HTMLButtonElement;
const overflowPopup = document.getElementById('tab-overflow-popup') as HTMLDivElement;
const contextMenu = document.getElementById('tab-context-menu') as HTMLDivElement;
const archivePopup = document.getElementById('archive-popup') as HTMLDivElement;

const COPY_ICON_HTML = `<span class="copy-icon"><svg class="copy-svg" viewBox="0 0 12 12" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="0.75" width="7.25" height="7.25" rx="1.25"/><rect x="0.75" y="4" width="7.25" height="7.25" rx="1.25"/></svg></span>`;
const COPIED_ICON_HTML = `<svg class="copy-svg" viewBox="0 0 12 12" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2 6l3 3.5 5-7"/></svg>`;
// Snapshot of editor.value from the end of the last input/auto-format pass,
// used to detect whole-line inserts/deletes so L<n> refs can shift to follow
// their target lines.
let previousText = '';

async function init() {
  settings = await window.mathPopup.getSettings();
  pages = settings.pages || [];
  closedPages = settings.closedPages || [];
  activePageId = settings.activePageId || '';
  if (pages.length === 0) {
    const id = Date.now().toString();
    pages.push({ id, title: 'Page 1', content: settings.noteContent ?? '', mode: 'text', lineModes: ['text'] });
    activePageId = id;
  }
  const activePage = pages.find(p => p.id === activePageId) || pages[0];
  activePageId = activePage.id;

  editor.value = activePage.content;
  previousText = editor.value;
  applyTheme(settings.theme);
  loadPageModes(activePage);
  applyAlwaysOnTop(settings.alwaysOnTop);
  applyZoom(settings.zoom ?? 1, { silent: true });
  bindEvents();
  // Re-render syntax overlay if the system theme flips while the app is open.
  window.mathPopup.onThemeChanged(() => render());
  updatePageIndicator();
  // Restore the tab bar's expanded/collapsed state from last session.
  if (settings.tabBarOpen) showTabBar();
  render();
  editor.focus();
}

// ============================================================
// Zoom
// ============================================================
let currentZoom = 1;
let zoomFlashTimer: number | null = null;
let zoomSaveTimer: number | null = null;

function clampZoom(z: number): number {
  if (!Number.isFinite(z)) return 1;
  if (z < ZOOM_MIN) return ZOOM_MIN;
  if (z > ZOOM_MAX) return ZOOM_MAX;
  // Snap to one decimal place to avoid floating-point drift across many steps.
  return Math.round(z * 10) / 10;
}

function applyZoom(factor: number, opts: { silent?: boolean } = {}) {
  const next = clampZoom(factor);
  currentZoom = next;
  window.mathPopup.setZoomFactor(next);
  updateZoomIndicator(!opts.silent);
  // Persist (debounced) so it survives restart. Skip during the initial
  // hydrate so we don't echo the value back to disk for no reason.
  if (!opts.silent) {
    if (zoomSaveTimer) window.clearTimeout(zoomSaveTimer);
    zoomSaveTimer = window.setTimeout(() => {
      window.mathPopup.setSettings({ zoom: currentZoom });
    }, 250);
  }
}

function updateZoomIndicator(flash: boolean) {
  const pct = Math.round(currentZoom * 100);
  zoomIndicator.textContent = `${pct}%`;
  // Visible whenever not at 100%, OR briefly after a change.
  const notDefault = pct !== 100;
  if (flash) {
    zoomIndicator.hidden = false;
    zoomIndicator.classList.add('flash');
    if (zoomFlashTimer) window.clearTimeout(zoomFlashTimer);
    zoomFlashTimer = window.setTimeout(() => {
      zoomIndicator.classList.remove('flash');
      // Hide only if we're back at 100% — otherwise stay visible (no flash).
      if (Math.round(currentZoom * 100) === 100) zoomIndicator.hidden = true;
    }, 1200);
  } else {
    zoomIndicator.hidden = !notDefault;
    zoomIndicator.classList.remove('flash');
  }
}

function zoomBy(deltaSteps: number) {
  applyZoom(currentZoom + deltaSteps * ZOOM_STEP);
}

function resetZoom() {
  applyZoom(settings?.zoomDefault ?? 1);
}

function updatePageIndicator() {
  const activePage = pages.find(p => p.id === activePageId);
  if (activePage) pageIndicator.textContent = activePage.title || 'Page';
  // Keep the inline tab bar's active chip in sync (no-op while it's closed).
  refreshTabBar();
}

function applyTheme(theme: ThemePref) {
  document.documentElement.setAttribute('data-theme', theme);
}

function bindEvents() {
  editor.addEventListener('input', onEditorInput);
  // Capture the scroll position before the browser mutates the DOM (and before
  // its native "keep caret visible" scroll), so onEditorInput can neutralize that
  // jump. Fires for typing, deletion, and paste alike.
  editor.addEventListener('beforeinput', () => { scrollBeforeInput = editor.scrollTop; });
  // IME: don't normalize the DOM mid-composition (it would cancel the IME);
  // process once on compositionend.
  editor.addEventListener('compositionstart', () => { edComposing = true; });
  editor.addEventListener('compositionend', () => { edComposing = false; onEditorInput(); });
  editor.addEventListener('scroll', syncScroll);
  editor.addEventListener('keydown', onKeyDown);
  editor.addEventListener('blur', () => {
    // Defer so click on the menu can take effect.
    setTimeout(() => {
      if (!cmdMenu.contains(document.activeElement)) hideMenu();
    }, 100);
    hideSignatureTooltip();
  });
  editor.addEventListener('click', () => { edRefreshSelectionCache(); updateMenuFromCaret(true); updateActiveToken(); });
  // Right-click in the editor: switch the selected line(s) — or the caret line —
  // between math and text. (Electron shows no native editor menu, so nothing is
  // lost by handling this.)
  editor.addEventListener('contextmenu', (e) => {
    e.preventDefault();
    showLineModeMenu(e.clientX, e.clientY);
  });
  // Re-render when the caret moves to a different line, so the per-line ref
  // display (raw "" L# "" vs the live result), the active-answer highlight, and
  // the current-line number highlight all stay correct. Refresh the cached
  // offsets first, since the browser moved the caret natively.
  document.addEventListener('selectionchange', () => {
    if (document.activeElement !== editor) return;
    edRefreshSelectionCache();
    if (caretLineIndex() !== lastCaretLine) refreshCaretLineDisplay();
  });
  // Line-number gutter: click a number to drop its L-ref at the caret. Using
  // mousedown + preventDefault keeps the textarea focused and its caret intact
  // (the standard "toolbar button inserts into a focused field" trick), so the
  // ref lands exactly where the user was typing.
  lineGutter.addEventListener('mousedown', (e) => {
    const row = (e.target as HTMLElement).closest<HTMLElement>('.row[data-line]');
    if (!row) return;
    e.preventDefault();
    const n = parseInt(row.dataset.line || '', 10);
    if (!Number.isFinite(n)) return;
    // Double-click a number → toggle that line between math and text. Single
    // click → insert its L-ref at the caret (delayed a moment so a double-click
    // can pre-empt the insert).
    if (pendingRefClick && pendingRefClick.line === n) {
      window.clearTimeout(pendingRefClick.timer);
      pendingRefClick = null;
      toggleLineMode(n - 1);
      return;
    }
    if (pendingRefClick) {            // a click on a different line commits the previous one first
      window.clearTimeout(pendingRefClick.timer);
      const prev = pendingRefClick.line;
      pendingRefClick = null;
      insertLineRefAtCaret(prev);
    }
    const line = n;
    const timer = window.setTimeout(() => { pendingRefClick = null; insertLineRefAtCaret(line); }, 250);
    pendingRefClick = { line, timer };
  });

  // Find (Ctrl/Cmd+F). Window-level so it opens whatever currently has focus;
  // F3 / Shift+F3 cycle matches. Escape is handled by the capture-phase Escape
  // listener above (find takes priority over closing the window).
  window.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && !e.altKey && !e.shiftKey && (e.key === 'f' || e.key === 'F')) {
      e.preventDefault();
      // Toggle: if you're already typing in the find box, close it; otherwise
      // open it (or pull focus back into it) and select the text.
      if (findActive && document.activeElement === findInput) closeFind();
      else openFind();
    } else if (e.key === 'F3') {
      e.preventDefault();
      if (findActive) findNext(e.shiftKey ? -1 : 1); else openFind();
    }
  });
  findInput.addEventListener('input', () => { autosizeFindInput(); runFind(findInput.value); });
  findInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); findNext(e.shiftKey ? -1 : 1); }
  });
  findPrevBtn.addEventListener('click', () => { findNext(-1); findInput.focus(); });
  findNextBtn.addEventListener('click', () => { findNext(1); findInput.focus(); });
  findCloseBtn.addEventListener('click', () => closeFind());

  window.addEventListener('resize', () => render());

  closeBtn.addEventListener('click', () => window.mathPopup.hidePopup());
  settingsBtn.addEventListener('click', () => window.mathPopup.openSettings());
  pinBtn.addEventListener('click', toggleAlwaysOnTop);
  helpBtn.addEventListener('click', () => window.mathPopup.openHelp());

  // Zoom: Ctrl+wheel and Ctrl+Plus/Minus/0. Window-level so it works whether
  // the editor or any button has focus.
  window.addEventListener('wheel', (e) => {
    if (!e.ctrlKey) return;
    e.preventDefault();
    zoomBy(e.deltaY < 0 ? 1 : -1);
  }, { passive: false });
  window.addEventListener('keydown', (e) => {
    if (!(e.ctrlKey || e.metaKey) || e.altKey) return;
    // Plus: handles both '+' (shifted) and '=' (same key), and numpad add.
    if (e.key === '+' || e.key === '=' || e.code === 'NumpadAdd') {
      e.preventDefault();
      zoomBy(1);
    } else if (e.key === '-' || e.key === '_' || e.code === 'NumpadSubtract') {
      e.preventDefault();
      zoomBy(-1);
    } else if (e.key === '0' || e.code === 'Numpad0') {
      e.preventDefault();
      resetZoom();
    }
  });

  // Tabs: click the button to toggle the inline tab bar open/closed.
  tabsBtn.addEventListener('click', toggleTabBar);
  tabAddBtn.addEventListener('click', () => addTab());
  // Overflow chevron: opens on hover or click, lists the clipped tabs, and
  // right-click offers "Reorder tabs". The in-between "⋯" marker shares these
  // exact triggers (see wireOverflowTrigger).
  wireOverflowTrigger(overflowBtn);
  overflowPopup.addEventListener('mouseenter', cancelHideOverflow);
  overflowPopup.addEventListener('mouseleave', scheduleHideOverflow);
  // Re-evaluate which tabs overflow when the window is resized.
  window.addEventListener('resize', () => { layoutTabs(); hideOverflowPopup(); hideTabContextMenu(); });
  // Dismiss the tab right-click menu / dropdown on any outside click.
  document.addEventListener('mousedown', (e) => {
    const t = e.target as Node;
    if (!contextMenu.hidden && !contextMenu.contains(t)) hideTabContextMenu();
    if (!overflowPopup.hidden && !reorderMode &&
        !overflowPopup.contains(t) && !overflowBtn.contains(t) && !contextMenu.contains(t)) {
      hideOverflowPopup();
    }
  });
  // Escape closes (in priority) the right-click menu, reorder mode, then the
  // dropdown — before the editor's Escape hides the whole window. Capture phase
  // so it pre-empts the editor handler; only swallow Escape if it closed one.
  document.addEventListener('keydown', (e) => {
    if (e.key !== 'Escape') return;
    let handled = true;
    if (findActive) closeFind();
    else if (!contextMenu.hidden) hideTabContextMenu();
    else if (reorderMode) exitReorderMode(false);
    else if (!overflowPopup.hidden) hideOverflowPopup();
    else handled = false;
    if (handled) { e.preventDefault(); e.stopPropagation(); }
  }, true);
  window.addEventListener('blur', hideTabContextMenu);

  // Dropdowns. Schedule the show on hover-intent so brief mouse passes
  // (e.g. moving back into the window) don't pop dropdowns open.
  archiveBtn.addEventListener('mouseenter', scheduleShowArchivePopup);
  archiveBtn.addEventListener('mouseleave', () => { cancelShowArchivePopup(); scheduleHideArchivePopup(); });
  archiveBtn.addEventListener('click', showArchivePopup);
  archivePopup.addEventListener('mouseenter', cancelHideArchivePopup);
  archivePopup.addEventListener('mouseleave', scheduleHideArchivePopup);

  // Variables popup
  varsBtn.addEventListener('mouseenter', scheduleShowVarsPopup);
  varsBtn.addEventListener('mouseleave', () => { cancelShowVarsPopup(); scheduleHideVarsPopup(); });
  varsBtn.addEventListener('click', showVarsPopup);
  varsBtn.addEventListener('focus', showVarsPopup);
  varsBtn.addEventListener('blur', hideVarsPopup);
  varsPopup.addEventListener('mouseenter', cancelHideVarsPopup);
  varsPopup.addEventListener('mouseleave', scheduleHideVarsPopup);

  // Update indicator: dot on the settings (⚙) button while there's a pending
  // update. Pull the current phase on load, then react to live changes.
  window.mathPopup.getUpdateState().then(applyUpdateIndicator);
  window.mathPopup.onUpdateState(applyUpdateIndicator);

  // Listen for settings changes pushed via a polling refresh-on-focus.
  window.addEventListener('focus', async () => {
    settings = await window.mathPopup.getSettings();
    pages = settings.pages || [];
    if (!pages.find(p => p.id === activePageId) && pages.length > 0) {
      activePageId = settings.activePageId || pages[0].id;
      const activePage = pages.find(p => p.id === activePageId) || pages[0];
      editor.value = activePage.content;
      previousText = editor.value;
      loadPageModes(activePage);
    }
    applyTheme(settings.theme);
    applyAlwaysOnTop(settings.alwaysOnTop);
    // If settings UI changed the saved zoom, reflect it. We compare to the
    // running value so the user's transient Ctrl+scroll level isn't clobbered
    // every time the window regains focus.
    if (typeof settings.zoom === 'number' && Math.abs(settings.zoom - currentZoom) > 0.001) {
      applyZoom(settings.zoom, { silent: true });
    }
    render();
  });
}

function lineModeAt(i: number): Mode {
  return lineModes[i] ?? 'text';
}
// The 0-based index of the line the caret is on.
function caretLineIndex(): number {
  const upto = editor.value.slice(0, editor.selectionStart);
  let n = 0;
  for (let i = 0; i < upto.length; i++) if (upto[i] === '\n') n++;
  return n;
}
// Inclusive 0-based line range the current selection touches. A collapsed caret
// yields a single line. A selection that ends exactly at a line start (just past a
// newline) doesn't count that trailing line, so selecting "line\n" affects only
// that line.
function selectedLineRange(): { start: number; end: number } {
  const text = editor.value;
  const selStart = editor.selectionStart;
  const selEnd = editor.selectionEnd;
  const lineOf = (off: number) => {
    let n = 0;
    for (let i = 0; i < off && i < text.length; i++) if (text[i] === '\n') n++;
    return n;
  };
  let start = lineOf(selStart);
  let end = lineOf(selEnd);
  if (selEnd > selStart && selEnd > 0 && text[selEnd - 1] === '\n') end = Math.max(start, end - 1);
  return { start: Math.min(start, end), end: Math.max(start, end) };
}
function countLines(t: string): number {
  let n = 1;
  for (let i = 0; i < t.length; i++) if (t[i] === '\n') n++;
  return n;
}
function padLineModes(t: string) {
  const n = countLines(t);
  const fixed: Mode[] = [];
  for (let i = 0; i < n; i++) fixed[i] = lineModes[i] ?? 'text';
  lineModes = fixed;
}
// Load the active page's per-line modes (deriving from the legacy per-page mode
// for notes saved before per-line modes existed) and re-anchor the sync.
function loadPageModes(page: Page) {
  const n = countLines(editor.value);
  if (Array.isArray(page.lineModes) && page.lineModes.length) {
    lineModes = Array.from({ length: n }, (_, i) => page.lineModes![i] ?? 'text');
  } else {
    const seed: Mode = page.mode === 'math' ? 'math' : 'text';
    lineModes = Array.from({ length: n }, () => seed);
  }
  lineModesText = editor.value;
}
// Keep lineModes aligned to the current text. Called at the top of render() —
// the single choke point all edits pass through. Newly inserted lines inherit
// the mode of the line above the insertion (so Enter carries math/text forward).
function syncLineModes() {
  const newText = editor.value;
  if (newText === lineModesText) {
    if (lineModes.length !== countLines(newText)) padLineModes(newText);
    return;
  }
  const oldLines = lineModesText.split('\n');
  const newLines = newText.split('\n');
  const oldLen = oldLines.length;
  const newLen = newLines.length;

  if (newLen === oldLen + 1) {
    // Exactly one line inserted (typically pressing Enter): the caret sits on
    // the new line, so place it there and inherit the line just above it. This
    // is caret-based, so it stays correct even amid blank lines — which a
    // content diff can't tell apart.
    const k = Math.min(caretLineIndex(), newLen - 1);
    const next: Mode[] = new Array(newLen);
    for (let i = 0; i < k; i++) next[i] = lineModes[i] ?? 'text';
    next[k] = lineModes[Math.max(0, k - 1)] ?? 'text';
    for (let i = k + 1; i < newLen; i++) next[i] = lineModes[i - 1] ?? 'text';
    lineModes = next;
    lineModesText = newText;
    return;
  }

  if (oldLen !== newLen) {
    // Line-based diff: the common leading and trailing lines keep their mode;
    // the changed middle is new and inherits the mode of the line just above it
    // (so pressing Enter on a math line keeps the new line math, and merging a
    // line into a math line above doesn't flip that line to text).
    let lead = 0;
    while (lead < oldLen && lead < newLen && oldLines[lead] === newLines[lead]) lead++;
    let trail = 0;
    while (trail < oldLen - lead && trail < newLen - lead &&
           oldLines[oldLen - 1 - trail] === newLines[newLen - 1 - trail]) trail++;
    const next: Mode[] = new Array(newLen);
    for (let i = 0; i < lead; i++) next[i] = lineModes[i] ?? 'text';
    for (let j = 0; j < trail; j++) next[newLen - 1 - j] = lineModes[oldLen - 1 - j] ?? 'text';
    // New middle lines inherit the mode they continue from: a pure insertion
    // takes the line above (lead-1); a merge/replace takes the first changed
    // line (lead), so backspacing a math line into a math line above stays math.
    const oldChanged = (oldLen - trail) > lead;
    const inherit: Mode = oldChanged
      ? (lineModes[lead] ?? 'text')
      : (lineModes[Math.max(0, lead - 1)] ?? 'text');
    for (let i = lead; i < newLen - trail; i++) next[i] = inherit;
    lineModes = next;
  }
  lineModesText = newText;
}
// Flip one line between math and text (the gutter-number click), persist, rerender.
function toggleLineMode(i: number) {
  if (i < 0) return;
  padLineModes(editor.value);
  lineModes[i] = lineModeAt(i) === 'math' ? 'text' : 'math';
  const activePage = pages.find(p => p.id === activePageId);
  if (activePage) activePage.lineModes = [...lineModes];
  window.mathPopup.setSettings({ pages, activePageId, closedPages });
  render();
  editor.focus();
}
// Force an inclusive range of lines to a specific mode (the right-click menu),
// persist, rerender. Used for both a single line and a multi-line selection.
function setLinesMode(startLine: number, endLine: number, mode: Mode) {
  if (startLine < 0) return;
  padLineModes(editor.value);
  for (let i = startLine; i <= endLine && i < lineModes.length; i++) {
    lineModes[i] = mode;
  }
  const activePage = pages.find(p => p.id === activePageId);
  if (activePage) activePage.lineModes = [...lineModes];
  window.mathPopup.setSettings({ pages, activePageId, closedPages });
  render();
  editor.focus();
}

function applyUpdateIndicator(state: { phase: string }) {
  const showDot =
    state.phase === 'available' ||
    state.phase === 'downloading' ||
    state.phase === 'downloaded';
  settingsBtn.classList.toggle('has-update', showDot);
  settingsBtn.title = showDot ? 'Settings — update available' : 'Settings';
}

// "Current mode" = the mode of the line the caret is on. The caret-relative
// behaviours (auto-format, smart-tab, list continuation, menus, status bar) all
// read this, so each just works per line.
function currentMode(): Mode {
  return lineModeAt(caretLineIndex());
}

function applyAlwaysOnTop(on: boolean) {
  pinBtn.classList.toggle('active', on);
  pinBtn.setAttribute('aria-pressed', on ? 'true' : 'false');
  pinBtn.title = on ? 'Stay on top: on' : 'Stay on top: off';
}

function toggleAlwaysOnTop() {
  const next = !settings.alwaysOnTop;
  settings.alwaysOnTop = next;
  applyAlwaysOnTop(next);
  window.mathPopup.setAlwaysOnTop(next);
}

function onInput() {
  justClosedTab = false;
  const previousToken = activeToken;
  activeToken = null;
  // Refs / variable renames can live on any math line regardless of where the
  // caret is, so keep them in sync on every edit.
  maybeSyncRename(previousToken);
  maybeShiftLineRefs();
  noteTypingForUndo();
  previousText = editor.value;
  scheduleSave();
  render();
  ensureCaretLineVisible();
  updateMenuFromCaret();
  updateSignatureTooltip();
}

// When the user inserts or deletes whole lines, rewrite `L<n>` references in
// the surviving (non-newly-typed) lines so they continue to point at the same
// target. Example: deleting line 1 shifts everything up; an `L2 - 50` on the
// (now) line 2 becomes `L1 - 50`. Also handles `L<a>:L<b>` ranges.
function maybeShiftLineRefs() {
  const oldLines = previousText.split('\n');
  const newLines = editor.value.split('\n');
  if (oldLines.length === newLines.length) return;
  const shift = computeLineShift(oldLines, newLines);
  if (!shift) return;
  const caret = editor.selectionStart;
  const caretEnd = editor.selectionEnd;
  const rewritten = rewriteLineRefs(editor.value, newLines, shift, caret, caretEnd);
  if (rewritten.text === editor.value) return;
  captureForUndo();
  editor.value = rewritten.text;
  editor.selectionStart = rewritten.caret;
  editor.selectionEnd = rewritten.caretEnd;
}

// Mirror an edit made to a variable's DEFINITION across every reference to it.
// The rule (and the user-facing contract): renaming the name on its definition
// line (`shares = …`) renames every use; editing a single reference is left local
// (only that one spot changes — that's how you point a spot at something else, or
// introduce a new name). Derived purely from the text diff, so it fires whether
// the caret reached the name by click, keyboard, or selection — no prior "active
// token" needed.
function maybeSyncRename(_previousToken: ActiveToken | null) {
  const oldText = previousText;
  const newText = editor.value;
  if (oldText === newText) return;

  // 1. Locate the single contiguous edit (common prefix/suffix).
  let prefix = 0;
  while (prefix < oldText.length && prefix < newText.length && oldText[prefix] === newText[prefix]) {
    prefix++;
  }
  let suffix = 0;
  while (suffix < oldText.length - prefix && suffix < newText.length - prefix &&
         oldText[oldText.length - 1 - suffix] === newText[newText.length - 1 - suffix]) {
    suffix++;
  }

  const editStart = prefix;
  const oldEditEnd = oldText.length - suffix;
  const newEditEnd = newText.length - suffix;
  const deletedText = oldText.slice(editStart, oldEditEnd);
  const insertedText = newText.slice(editStart, newEditEnd);

  // 2. The edit must stay inside a single identifier (only word chars touched), so
  //    it reads as renaming a name — not restructuring the line.
  if (/[^A-Za-z0-9_]/.test(deletedText) || /[^A-Za-z0-9_]/.test(insertedText)) return;
  const isWord = (c: string) => /[A-Za-z0-9_]/.test(c);
  let wStart = editStart;
  while (wStart > 0 && isWord(oldText[wStart - 1])) wStart--;
  let wEnd = oldEditEnd;
  while (wEnd < oldText.length && isWord(oldText[wEnd])) wEnd++;
  if (wStart === wEnd) return;                       // edit didn't land on an identifier

  const oldName = oldText.slice(wStart, wEnd);
  const newName = oldText.slice(wStart, editStart) + insertedText + oldText.slice(oldEditEnd, wEnd);
  // Both the old and new word must be real identifiers (not a number; not emptied).
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(oldName)) return;
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(newName)) return;
  const oldNameLow = oldName.toLowerCase();

  // 3. Only mirror when the edited name is the variable DEFINED on this line (its
  //    assignment target, i.e. nothing precedes it on the line). A reference — or
  //    a same-line reuse like the 2nd `x` in `x = x + 1` — stays local.
  const lineIndex = (oldText.slice(0, wStart).match(/\n/g) || []).length;
  const lineResult = lastResults[lineIndex];
  if (!lineResult || !lineResult.varName || lineResult.varName.toLowerCase() !== oldNameLow) return;
  const lineStart = oldText.lastIndexOf('\n', wStart - 1) + 1;
  if (oldText.slice(lineStart, wStart).trim() !== '') return;   // not the LHS → it's a reference

  // 4. All whole-word occurrences of the old name.
  const occurrences: { start: number; end: number }[] = [];
  const re = new RegExp(`(^|[^A-Za-z0-9_])${oldName}(?![A-Za-z0-9_])`, 'gi');
  let m: RegExpExecArray | null;
  while ((m = re.exec(oldText)) !== null) {
    const matchStart = m.index + m[1].length;
    occurrences.push({ start: matchStart, end: matchStart + oldName.length });
  }
  if (occurrences.length <= 1) return;               // nothing else references it

  const editedOccIdx = occurrences.findIndex(occ => occ.start === wStart);
  if (editedOccIdx === -1) return;
  const editedOcc = occurrences[editedOccIdx];

  const relStart = editStart - editedOcc.start;
  const relEnd = oldEditEnd - editedOcc.start;
  const delta = insertedText.length - deletedText.length;

  let outText = oldText;
  let newCaretStart = editor.selectionStart;
  let newCaretEnd = editor.selectionEnd;

  // Apply the exact same edit to every occurrence, right-to-left to keep offsets
  // valid. Skip OTHER lines that DEFINE the same name (another LHS) so two
  // separate variables never get merged — the duplicate-name error still flags it.
  for (let i = occurrences.length - 1; i >= 0; i--) {
    const occ = occurrences[i];

    if (i !== editedOccIdx) {
      const occLineIndex = (oldText.slice(0, occ.start).match(/\n/g) || []).length;
      const occLineResult = lastResults[occLineIndex];
      const occIsBase = occLineResult && occLineResult.varName
        && occLineResult.varName.toLowerCase() === oldNameLow;
      if (occIsBase) {
        const occLineStart = oldText.lastIndexOf('\n', occ.start - 1) + 1;
        if (oldText.slice(occLineStart, occ.start).trim() === '') continue;  // another definition — leave it
      }
    }

    const absStart = occ.start + relStart;
    const absEnd = occ.start + relEnd;
    outText = outText.slice(0, absStart) + insertedText + outText.slice(absEnd);

    // Shift caret if an edit happened before it.
    if (i !== editedOccIdx && occ.start < editedOcc.start) {
      newCaretStart += delta;
      newCaretEnd += delta;
    }
  }

  editor.value = outText;
  editor.selectionStart = newCaretStart;
  editor.selectionEnd = newCaretEnd;
  // Keep the renamed token active so its uses stay highlighted as one group.
  activeToken = { type: 'var', name: newName.toLowerCase() };
}

interface LineShift {
  // For each old line index, the new index it ended up at (or undefined if
  // the line was deleted).
  map: Map<number, number>;
  // Range of new indices (inclusive start, exclusive end) that came from the
  // old text. Lines outside this range are newly inserted/typed and should
  // NOT have their L<n> tokens rewritten.
  shiftedPrefixEnd: number;          // [0, shiftedPrefixEnd) is shifted
  shiftedSuffixStart: number;        // [shiftedSuffixStart, newLen) is shifted
}

function computeLineShift(oldLines: string[], newLines: string[]): LineShift | null {
  const oldText = oldLines.join('\\n');
  const newText = newLines.join('\\n');
  const minLen = Math.min(oldText.length, newText.length);
  
  let prefix = 0;
  while (prefix < minLen && oldText[prefix] === newText[prefix]) prefix++;
  
  let suffix = 0;
  while (suffix < minLen - prefix && 
         oldText[oldText.length - 1 - suffix] === newText[newText.length - 1 - suffix]) {
    suffix++;
  }

  const oldPrefixLines = (oldText.slice(0, prefix).match(/\\n/g) || []).length;
  const newPrefixLines = (newText.slice(0, prefix).match(/\\n/g) || []).length;
  const oldSuffixLines = (oldText.slice(oldText.length - suffix).match(/\\n/g) || []).length;
  const newSuffixLines = (newText.slice(newText.length - suffix).match(/\\n/g) || []).length;

  const oldLen = oldLines.length;
  const newLen = newLines.length;
  
  const map = new Map<number, number>();
  let anyShift = false;
  
  for (let i = 0; i < oldPrefixLines; i++) map.set(i, i);
  for (let i = 0; i <= oldSuffixLines && i < oldLen; i++) {
    const oldIdx = oldLen - 1 - i;
    const newIdx = newLen - 1 - i;
    map.set(oldIdx, newIdx);
    if (oldIdx !== newIdx) anyShift = true;
  }
  
  if (!anyShift) return null;
  return { map, shiftedPrefixEnd: oldPrefixLines, shiftedSuffixStart: newLen - oldSuffixLines };
}

function rewriteLineRefs(
  text: string,
  newLines: string[],
  shift: LineShift,
  caret: number,
  caretEnd: number
): { text: string; caret: number; caretEnd: number } {
  // Match an L<a>:L<b> range OR a bare L<n> reference. Order matters here —
  // ranges first so the L\d+ alternative doesn't gobble half of a range.
  const re = /\bL(\d+)\s*:\s*L(\d+)\b|\bL(\d+)\b/gi;
  let pos = 0;            // start offset of current line in `text`
  let newCaret = caret;
  let newCaretEnd = caretEnd;
  const outLines: string[] = [];

  for (let i = 0; i < newLines.length; i++) {
    const line = newLines[i];
    const lineStart = pos;
    const isShifted = i < shift.shiftedPrefixEnd || i >= shift.shiftedSuffixStart;
    if (!isShifted) {
      outLines.push(line);
      pos = lineStart + line.length + 1;
      continue;
    }
    const edits: { start: number; end: number; replacement: string }[] = [];
    re.lastIndex = 0;
    let m: RegExpExecArray | null;
    while ((m = re.exec(line)) !== null) {
      if (m[1] !== undefined && m[2] !== undefined) {
        // L<a>:L<b> range
        const aOld = Number(m[1]) - 1;
        const bOld = Number(m[2]) - 1;
        const aNew = shift.map.get(aOld);
        const bNew = shift.map.get(bOld);
        if (aNew === undefined || bNew === undefined) continue;
        if (aNew === aOld && bNew === bOld) continue;
        const lPrefix = m[0][0]; // 'L' or 'l'
        edits.push({
          start: m.index,
          end: m.index + m[0].length,
          replacement: `${lPrefix}${aNew + 1}:${lPrefix}${bNew + 1}`
        });
      } else if (m[3] !== undefined) {
        const oldIdx = Number(m[3]) - 1;
        const newIdx = shift.map.get(oldIdx);
        if (newIdx === undefined || newIdx === oldIdx) continue;
        const lPrefix = m[0][0];
        edits.push({
          start: m.index,
          end: m.index + m[0].length,
          replacement: lPrefix + String(newIdx + 1)
        });
      }
    }
    let outLine = line;
    // Apply from the end so earlier offsets remain valid.
    for (let k = edits.length - 1; k >= 0; k--) {
      const e = edits[k];
      outLine = outLine.slice(0, e.start) + e.replacement + outLine.slice(e.end);
      const absStart = lineStart + e.start;
      const absEnd = lineStart + e.end;
      const delta = e.replacement.length - (e.end - e.start);
      newCaret = adjustCaret(newCaret, caret, absStart, absEnd, delta, e.replacement.length);
      newCaretEnd = adjustCaret(newCaretEnd, caretEnd, absStart, absEnd, delta, e.replacement.length);
    }
    outLines.push(outLine);
    pos = lineStart + line.length + 1;
  }
  return { text: outLines.join('\n'), caret: newCaret, caretEnd: newCaretEnd };
}

function adjustCaret(
  current: number,
  original: number,
  absStart: number,
  absEnd: number,
  delta: number,
  replacementLen: number
): number {
  if (original >= absEnd) return current + delta;
  if (original > absStart) return absStart + replacementLen;
  return current;
}

// Browsers only auto-scroll a textarea enough to make the caret pixel visible,
// not the whole line. That leaves the bottom of the cursor's line (and its
// gutter row) clipped when typing near the bottom of the editor. Scroll the
// editor so the entire line containing the caret is in view.
function ensureCaretLineVisible() {
  if (editor.selectionStart !== editor.selectionEnd) return;
  const caret = editor.selectionStart;

  const editorStyle = getComputedStyle(editor);
  const lineHeight = parseFloat(editorStyle.lineHeight) || 22;

  const coords = caretCoords(caret);
  const lineTop = coords.top;
  const lineBot = lineTop + lineHeight;

  const viewTop = editor.scrollTop;
  const viewBot = viewTop + editor.clientHeight;

  if (lineBot > viewBot) {
    editor.scrollTop = lineBot - editor.clientHeight;
  } else if (lineTop < viewTop) {
    editor.scrollTop = lineTop;
  }
  // Keep the overlays aligned with whatever scroll position we just landed on.
  syncScroll();
}

function scheduleSave() {
  if (saveTimer) window.clearTimeout(saveTimer);
  saveTimer = window.setTimeout(() => {
    const activePage = pages.find(p => p.id === activePageId);
    if (activePage) { activePage.content = editor.value; activePage.lineModes = [...lineModes]; }
    window.mathPopup.setSettings({ pages, activePageId });
  }, 250);
}

let archiveHoverTimer: number | null = null;
let archiveShowTimer: number | null = null;

// Hover-intent delay: dropdowns only open if the cursor lingers on the
// button. Prevents stray mouse passes (e.g. moving back into the window)
// from popping menus open.
const HOVER_INTENT_MS = 300;

// ---- tab bar (inline, toggled by the tabs button) ----
function toggleTabBar() {
  if (tabBar.classList.contains('open')) hideTabBar();
  else showTabBar();
}
function showTabBar() {
  tabBar.classList.add('open');
  tabsBtn.classList.add('active');
  renderTabBar();
  persistTabBarState(true);
}
function hideTabBar() {
  tabBar.classList.remove('open');
  tabsBtn.classList.remove('active');
  hideOverflowPopup();
  persistTabBarState(false);
}
// Remember whether the tab bar is expanded so it can be restored next launch.
function persistTabBarState(open: boolean) {
  if (!settings || settings.tabBarOpen === open) return;
  settings.tabBarOpen = open;
  window.mathPopup.setSettings({ tabBarOpen: open });
}
// Re-render the chips (e.g. after a switch/add/close) only while the bar is open.
function refreshTabBar() {
  if (tabBar.classList.contains('open')) renderTabBar();
}

const TAB_GAP = 4;          // must match .tab-strip `gap` in popup.css
const OVERFLOW_MARKER = 22; // space reserved for the "from dropdown" double separator

// Decide which tabs fully fit and which go to the overflow dropdown. Tabs are
// never clipped — each either shows in full or is hidden. The active tab is
// always kept visible; if it would have overflowed, it's surfaced at the end
// of the bar with a double-separator marker so it's clearly "from the menu".
// The overflow chevron and the in-between "⋯" marker share triggers: hover opens
// the hidden-tab list, click toggles it, right-click offers Reorder. stopProp
// keeps the marker — which lives inside a tab chip — from also activating that tab.
function wireOverflowTrigger(el: HTMLElement) {
  el.addEventListener('click', (e) => { e.stopPropagation(); toggleOverflowPopup(); });
  el.addEventListener('mouseenter', () => { cancelHideOverflow(); showOverflowPopup(); });
  el.addEventListener('mouseleave', scheduleHideOverflow);
  el.addEventListener('contextmenu', (e) => {
    e.preventDefault();
    e.stopPropagation();
    cancelHideOverflow();
    showTabContextMenu(null, e.clientX, e.clientY);
  });
}

function layoutTabs() {
  if (!tabBar.classList.contains('open')) return;
  hideTooltip();   // a relayout recreates the overflow marker; drop any tooltip anchored to the old one
  const chips = Array.from(tabStrip.children) as HTMLElement[];
  chips.forEach(c => {
    c.classList.remove('tab-hidden', 'from-overflow');
    c.querySelector('.tab-overflow-mark')?.remove();
  });
  if (chips.length === 0) { overflowBtn.hidden = true; return; }

  const widths = chips.map(c => c.offsetWidth);
  const total = widths.reduce((a, b) => a + b, 0) + TAB_GAP * (chips.length - 1);

  // The strip normally hugs its tabs so the "+" trails the last one. Make it fill
  // the bar while we measure, so clientWidth reports the room available for tabs
  // (not just the current content width). We collapse back to hug below.
  tabBar.classList.add('tab-measuring');

  // Everything fits with no chevron?
  overflowBtn.hidden = true;
  if (total <= tabStrip.clientWidth) { tabBar.classList.remove('tab-measuring'); return; }

  // Overflow exists: show the chevron, then recompute against the (now smaller)
  // strip width. `leadingFit` returns how many leading chips fit in `avail`.
  overflowBtn.hidden = false;
  const W = tabStrip.clientWidth;
  tabBar.classList.remove('tab-measuring');   // measured; collapse so "+" hugs the last tab
  const leadingFit = (avail: number) => {
    let used = 0, count = 0;
    for (let i = 0; i < chips.length; i++) {
      const need = (count > 0 ? TAB_GAP : 0) + widths[i];
      if (used + need <= avail) { used += need; count++; } else break;
    }
    return count;
  };

  const activeIdx = Math.max(0, chips.findIndex(c => c.dataset.pageId === activePageId));
  const naturalCount = leadingFit(W);

  if (activeIdx <= naturalCount - 1) {
    // Active fits naturally; hide everything past the prefix.
    for (let i = naturalCount; i < chips.length; i++) chips[i].classList.add('tab-hidden');
    setOverflowCount();
    return;
  }

  // Active would overflow: surface it at the end with a marker.
  const reserve = widths[activeIdx] + TAB_GAP + OVERFLOW_MARKER;
  const leadCount = leadingFit(Math.max(0, W - reserve));
  for (let i = 0; i < chips.length; i++) {
    if (i < leadCount || i === activeIdx) continue;
    chips[i].classList.add('tab-hidden');
  }
  if (leadCount > 0) {
    const active = chips[activeIdx];
    active.classList.add('from-overflow');
    const mark = document.createElement('span');
    mark.className = 'tab-overflow-mark';
    // Empty title suppresses the parent tab's native "name" tooltip from also
    // popping up when the cursor is over the marker (it's a child of that tab).
    mark.title = '';
    // Informational only: hovering it explains there are hidden tabs here, using
    // the shared styled tooltip with a delay so a quick pass doesn't trigger it.
    mark.addEventListener('mouseenter', () => showTooltipHTML(mark,
      `<div class="tip-title">There are extra tabs between here</div>` +
      `<div>Open them from the menu on the right.</div>`, 500));
    mark.addEventListener('mouseleave', hideTooltip);
    active.insertBefore(mark, active.firstChild);
  }
  setOverflowCount();
}

// Show how many tabs are tucked away on the overflow chevron (it renders as a
// "N ⌄" pill), so it's obvious there are more tabs hiding there. Only meaningful
// when the chevron is visible, which is exactly when something overflows.
function setOverflowCount() {
  const n = (Array.from(tabStrip.children) as HTMLElement[])
    .filter(c => c.classList.contains('tab-hidden')).length;
  const countEl = overflowBtn.querySelector('.ov-count');
  if (countEl) countEl.textContent = n > 0 ? String(n) : '';
  const label = n === 1 ? '1 more tab' : `${n} more tabs`;
  overflowBtn.title = label;
  overflowBtn.setAttribute('aria-label', label);
}

function buildOverflowList() {
  overflowPopup.innerHTML = '';
  const hidden = (Array.from(tabStrip.children) as HTMLElement[])
    .filter(c => c.classList.contains('tab-hidden'));
  if (hidden.length === 0) {
    overflowPopup.innerHTML = `<div class="vars-empty">No hidden tabs</div>`;
    return;
  }
  hidden.forEach((chip) => {
    const id = chip.dataset.pageId!;
    const page = pages.find(p => p.id === id);
    if (!page) return;
    const row = document.createElement('div');
    row.className = 'ov-row';

    const name = document.createElement('span');
    name.className = 'ov-name';
    name.textContent = page.title || 'Tab';

    const close = document.createElement('span');
    close.className = 'ov-close';
    close.textContent = '×';
    close.title = 'Close tab';
    close.onclick = (e) => { e.stopPropagation(); closeTab(id); refreshOverflowPopup(); };

    row.appendChild(name);
    row.appendChild(close);
    row.onclick = () => { switchTab(id); hideOverflowPopup(); };
    row.oncontextmenu = (e) => {
      e.preventDefault();
      cancelHideOverflow();   // keep the dropdown open behind the context menu
      showTabContextMenu(id, e.clientX, e.clientY);
    };
    overflowPopup.appendChild(row);
  });
}
// Keep an open dropdown in sync after a close (rebuild, or dismiss if empty).
function refreshOverflowPopup() {
  if (overflowPopup.hidden || reorderMode) return;
  if (overflowBtn.hidden) hideOverflowPopup();
  else buildOverflowList();
}
// Right-align the popup under the chevron (or the + button if the chevron is
// hidden — e.g. Reorder invoked when nothing overflows), clamped to the window.
function positionOverflowPopup() {
  const anchor = overflowBtn.hidden ? tabAddBtn : overflowBtn;
  const rect = anchor.getBoundingClientRect();
  const left = Math.max(8, rect.right - overflowPopup.offsetWidth);
  overflowPopup.style.top = `${rect.bottom + 4}px`;
  overflowPopup.style.left = `${left}px`;
}
function showOverflowPopup() {
  if (overflowBtn.hidden || reorderMode) return;
  buildOverflowList();
  overflowPopup.hidden = false;
  positionOverflowPopup();
}
function hideOverflowPopup() {
  if (reorderMode) return;   // reorder mode is dismissed only via Done / Esc
  overflowPopup.hidden = true;
}
function toggleOverflowPopup() {
  if (overflowPopup.hidden) showOverflowPopup(); else hideOverflowPopup();
}
function scheduleHideOverflow() { overflowHideTimer = window.setTimeout(maybeHideOverflow, 180); }
function cancelHideOverflow() { if (overflowHideTimer) window.clearTimeout(overflowHideTimer); }
// Don't auto-hide the dropdown while its right-click menu or reorder mode is up.
function maybeHideOverflow() {
  if (!contextMenu.hidden || reorderMode) return;
  hideOverflowPopup();
}

// ---- reorder mode: drag ALL tabs into a new order, commit on Done ----
function enterReorderMode() {
  reorderMode = true;
  hideTabContextMenu();
  buildReorderList();
  overflowPopup.hidden = false;
  positionOverflowPopup();
}
function exitReorderMode(commit: boolean) {
  if (commit) commitReorderFromList();
  reorderMode = false;
  overflowPopup.hidden = true;
  renderTabBar();
}
function buildReorderList() {
  overflowPopup.innerHTML = '';
  const head = document.createElement('div');
  head.className = 'ov-reorder-head';
  const label = document.createElement('span');
  label.textContent = 'Drag to reorder';
  const done = document.createElement('button');
  done.className = 'ov-done';
  done.textContent = 'Done';
  done.onclick = () => exitReorderMode(true);
  head.appendChild(label);
  head.appendChild(done);
  overflowPopup.appendChild(head);

  pages.forEach((page) => {
    const row = document.createElement('div');
    row.className = 'reorder-row' + (page.id === activePageId ? ' active' : '');
    row.dataset.pageId = page.id;
    row.draggable = true;

    const handle = document.createElement('span');
    handle.className = 'reorder-handle';
    handle.textContent = '⠿'; // ⠿ braille dots = drag grip
    const name = document.createElement('span');
    name.className = 'reorder-name';
    name.textContent = page.title || 'Tab';
    row.appendChild(handle);
    row.appendChild(name);

    row.addEventListener('dragstart', (e) => {
      row.classList.add('dragging');
      if (e.dataTransfer) {
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', page.id);
      }
    });
    row.addEventListener('dragend', () => row.classList.remove('dragging'));
    row.addEventListener('dragover', (e) => {
      e.preventDefault();
      const dragging = overflowPopup.querySelector('.reorder-row.dragging') as HTMLElement | null;
      if (!dragging || dragging === row) return;
      const r = row.getBoundingClientRect();
      const after = e.clientY > r.top + r.height / 2;
      if (after) row.after(dragging); else row.before(dragging);
    });
    overflowPopup.appendChild(row);
  });
}
function commitReorderFromList() {
  const ids = (Array.from(overflowPopup.querySelectorAll('.reorder-row')) as HTMLElement[])
    .map(r => r.dataset.pageId);
  const reordered: Page[] = [];
  ids.forEach((id) => { const p = pages.find(pp => pp.id === id); if (p) reordered.push(p); });
  if (reordered.length === pages.length) {
    pages.splice(0, pages.length, ...reordered);
    window.mathPopup.setSettings({ pages, activePageId, closedPages });
  }
}

// Rebuild the `pages` order from the chips' DOM order after a drag-reorder.
function commitTabOrderFromDom() {
  const ids = (Array.from(tabStrip.querySelectorAll('.tab-chip')) as HTMLElement[])
    .map(c => c.dataset.pageId);
  const reordered: Page[] = [];
  ids.forEach((id) => { const p = pages.find(pp => pp.id === id); if (p) reordered.push(p); });
  if (reordered.length === pages.length) {
    pages.splice(0, pages.length, ...reordered);
    window.mathPopup.setSettings({ pages, activePageId, closedPages });
  }
  renderTabBar();
}

function showArchivePopup() {
  cancelShowArchivePopup();
  if (archiveHoverTimer) window.clearTimeout(archiveHoverTimer);
  renderArchiveMenu();
  archivePopup.hidden = false;
  const rect = archiveBtn.getBoundingClientRect();
  archivePopup.style.top = `${rect.bottom + 4}px`;
  archivePopup.style.left = `${rect.left}px`;
}
function scheduleShowArchivePopup() {
  cancelShowArchivePopup();
  archiveShowTimer = window.setTimeout(showArchivePopup, HOVER_INTENT_MS);
}
function cancelShowArchivePopup() {
  if (archiveShowTimer) { window.clearTimeout(archiveShowTimer); archiveShowTimer = null; }
}
function scheduleHideArchivePopup() { archiveHoverTimer = window.setTimeout(hideArchivePopup, 150); }
function cancelHideArchivePopup() { if (archiveHoverTimer) window.clearTimeout(archiveHoverTimer); }
function hideArchivePopup() { archivePopup.hidden = true; }

function renderTabBar() {
  tabStrip.innerHTML = '';
  pages.forEach((page, index) => {
    const chip = document.createElement('div');
    chip.className = 'tab-chip' + (page.id === activePageId ? ' active' : '');
    chip.dataset.pageId = page.id;
    chip.draggable = true;
    const label = page.title || `Page ${index + 1}`;
    chip.title = label;
    chip.onclick = (e) => {
      // Don't switch if interacting with the inline rename input
      if ((e.target as HTMLElement).tagName === 'INPUT') return;
      switchTab(page.id);
    };

    // ---- drag to reorder ----
    chip.addEventListener('dragstart', (e) => {
      dragSrcId = page.id;
      chip.classList.add('dragging');
      tabBar.classList.add('reordering');
      hideOverflowPopup();
      if (e.dataTransfer) {
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', page.id);
      }
    });
    chip.addEventListener('dragend', () => {
      chip.classList.remove('dragging');
      tabBar.classList.remove('reordering');
      dragSrcId = null;
      commitTabOrderFromDom();
    });
    chip.addEventListener('dragover', (e) => {
      e.preventDefault();
      if (e.dataTransfer) e.dataTransfer.dropEffect = 'move';
      const dragging = tabStrip.querySelector('.tab-chip.dragging') as HTMLElement | null;
      if (!dragging || dragging === chip) return;
      // Don't reorder against hidden or the surfaced-from-overflow chip.
      if (chip.classList.contains('tab-hidden') || chip.classList.contains('from-overflow')) return;
      const r = chip.getBoundingClientRect();
      const after = e.clientX > r.left + r.width / 2;
      if (after) chip.after(dragging); else chip.before(dragging);
    });

    // Right-click a tab for Rename / Close.
    chip.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      showTabContextMenu(page.id, e.clientX, e.clientY);
    });

    const nameSpan = document.createElement('span');
    nameSpan.className = 'tab-chip-name';
    nameSpan.textContent = label;

    // close (×) — always visible. On a BACKGROUND tab it's "disarmed" (pale red):
    // a click just selects the tab, so you can't close one by mis-clicking while
    // switching. On the ACTIVE tab the × is armed and a click closes it.
    const closeBtn = document.createElement('span');
    closeBtn.className = 'tab-chip-close';
    closeBtn.textContent = '×';
    closeBtn.title = page.id === activePageId ? 'Close tab' : 'Middle-click to close';
    closeBtn.onclick = (e) => {
      if (page.id !== activePageId) return;   // background tab: let it bubble → switchTab
      e.stopPropagation();
      closeTab(page.id);
    };

    // Middle-click closes any tab directly (browser-style), armed or not.
    chip.addEventListener('mousedown', (e) => {
      if (e.button === 1) e.preventDefault();   // suppress the middle-click autoscroll cursor
    });
    chip.addEventListener('auxclick', (e) => {
      if (e.button === 1) { e.preventDefault(); closeTab(page.id); }
    });

    chip.appendChild(nameSpan);
    chip.appendChild(closeBtn);
    tabStrip.appendChild(chip);
  });

  layoutTabs();
}

// Turn a tab's name into an inline text input (Enter/blur saves, Esc cancels).
// Used by both the right-click "Rename" item and the Ctrl+L shortcut.
function startRename(pageId: string) {
  if (!tabBar.classList.contains('open')) showTabBar();
  hideOverflowPopup();
  let chip = tabStrip.querySelector(
    `.tab-chip[data-page-id="${pageId}"]`) as HTMLElement | null;
  // If the tab is currently in the overflow dropdown, switch to it first so it
  // gets surfaced into the bar, then rename it there.
  if (!chip || chip.classList.contains('tab-hidden')) {
    switchTab(pageId);
    chip = tabStrip.querySelector(
      `.tab-chip[data-page-id="${pageId}"]`) as HTMLElement | null;
  }
  const page = pages.find(p => p.id === pageId);
  const nameSpan = chip?.querySelector('.tab-chip-name') as HTMLElement | null;
  if (!chip || !page || !nameSpan) return;

  chip.draggable = false; // let the user select text in the input
  const input = document.createElement('input');
  input.type = 'text';
  input.className = 'tab-rename-input';
  input.maxLength = 15;
  input.value = page.title;

  const saveName = () => {
    const val = input.value.trim();
    if (val) {
      page.title = val;
      window.mathPopup.setSettings({ pages, activePageId, closedPages });
      updatePageIndicator();
    }
    renderTabBar();
  };

  input.onkeydown = (e2) => {
    if (e2.key === 'Enter') { e2.preventDefault(); saveName(); }
    else if (e2.key === 'Escape') { e2.preventDefault(); renderTabBar(); }
  };
  input.onblur = saveName;

  chip.replaceChild(input, nameSpan);
  input.focus();
  input.select();
}

function renameActiveTab() { startRename(activePageId); }

// ---- tab right-click menu ----
// pageId null = menu not tied to a specific tab (e.g. the overflow chevron):
// only the "Reorder tabs" action is shown.
function showTabContextMenu(pageId: string | null, x: number, y: number) {
  contextMenu.innerHTML = '';
  const addItem = (label: string, key: string, danger: boolean, onPick: () => void) => {
    const item = document.createElement('div');
    item.className = 'ctx-item' + (danger ? ' danger' : '');
    const text = document.createElement('span');
    text.textContent = label;
    item.appendChild(text);
    if (key) {
      const k = document.createElement('span');
      k.className = 'ctx-key';
      k.textContent = key;
      item.appendChild(k);
    }
    item.onclick = () => { hideTabContextMenu(); cancelHideOverflow(); onPick(); };
    contextMenu.appendChild(item);
  };
  if (pageId) {
    addItem('Rename', 'Ctrl+L', false, () => startRename(pageId));
    addItem('Close tab', 'Ctrl+W', true, () => { closeTab(pageId); refreshOverflowPopup(); });
  }
  addItem('Reorder tabs', '', false, () => enterReorderMode());

  placeContextMenuAt(x, y);
}
function hideTabContextMenu() { contextMenu.hidden = true; }

// Show the menu (already populated) clamped inside the window.
function placeContextMenuAt(x: number, y: number) {
  contextMenu.hidden = false;
  const mw = contextMenu.offsetWidth;
  const mh = contextMenu.offsetHeight;
  contextMenu.style.left = `${Math.max(8, Math.min(x, window.innerWidth - mw - 8))}px`;
  contextMenu.style.top = `${Math.max(8, Math.min(y, window.innerHeight - mh - 8))}px`;
}

// Right-click menu inside the editor: switch the selected line(s) — or the caret
// line if nothing is selected — to math or text. Reuses the shared context-menu
// element, so the existing outside-click / Escape / blur dismissal covers it.
function showLineModeMenu(x: number, y: number) {
  edRefreshSelectionCache();                 // make sure the range reflects the live selection
  const { start, end } = selectedLineRange();
  const count = end - start + 1;
  // Offer the single opposite action: if every selected line is already math,
  // switch to text; otherwise switch to math (which also unifies a mixed
  // selection toward math). So you always get the "other" mode in one click.
  let allMath = true;
  for (let i = start; i <= end; i++) { if (lineModeAt(i) !== 'math') { allMath = false; break; } }
  const target: Mode = allMath ? 'text' : 'math';

  contextMenu.innerHTML = '';
  const item = document.createElement('div');
  item.className = 'ctx-item';
  const text = document.createElement('span');
  text.textContent = count > 1 ? `Switch ${count} lines to ` : 'Switch to ';
  // Color the mode word the same as everywhere else (green = math, blue = text)
  // so it's obvious which way you're switching.
  const mode = document.createElement('span');
  mode.className = `ctx-mode ctx-mode-${target}`;
  mode.textContent = target === 'math' ? 'Math' : 'Text';
  text.appendChild(mode);
  item.appendChild(text);
  item.onclick = () => { hideTabContextMenu(); setLinesMode(start, end, target); };
  contextMenu.appendChild(item);
  placeContextMenuAt(x, y);
}

function renderArchiveMenu() {
  archivePopup.innerHTML = '';
  if (closedPages.length === 0) {
    archivePopup.innerHTML = `<div class="vars-empty">No closed tabs</div>`;
    return;
  }
  closedPages.forEach((page, index) => {
    const row = document.createElement('div');
    row.className = 'vars-row';
    row.style.cursor = 'pointer';
    row.onclick = () => { restoreTab(index); hideArchivePopup(); };
    
    const titleSpan = document.createElement('span');
    titleSpan.className = 'vars-name';
    titleSpan.textContent = page.title || 'Tab';
    
    const restoreBtn = document.createElement('span');
    restoreBtn.className = 'vars-val';
    restoreBtn.textContent = '↺';
    
    row.appendChild(titleSpan);
    row.appendChild(restoreBtn);
    archivePopup.appendChild(row);
  });
}

function switchTab(id: string) {
  if (id === activePageId) return;
  justClosedTab = false;
  const current = pages.find(p => p.id === activePageId);
  if (current) current.content = editor.value;
  
  activePageId = id;
  const next = pages.find(p => p.id === activePageId)!;
  editor.value = next.content;
  previousText = editor.value;

  loadPageModes(next);
  window.mathPopup.setSettings({ pages, activePageId, closedPages });

  updatePageIndicator();
  render();
  editor.focus();
}

function addTab() {
  if (pages.length >= 99) return;
  justClosedTab = false;
  const current = pages.find(p => p.id === activePageId);
  if (current) { current.content = editor.value; current.lineModes = [...lineModes]; }

  const id = Date.now().toString();
  const title = `Page ${pages.length + 1}`;
  const page: Page = { id, title, content: '', mode: 'text', lineModes: ['text'] };
  pages.push(page);
  activePageId = id;

  editor.value = '';
  previousText = '';
  loadPageModes(page);
  window.mathPopup.setSettings({ pages, activePageId, closedPages });
  
  updatePageIndicator();
  render();
  editor.focus();
}

function closeTab(id: string) {
  const index = pages.findIndex(p => p.id === id);
  if (index === -1) return;
  
  const current = pages[index];
  if (id === activePageId) { current.content = editor.value; current.lineModes = [...lineModes]; }
  closedPages.unshift(current);
  if (closedPages.length > 10) closedPages.pop();

  pages.splice(index, 1);
  if (pages.length === 0) {
    const newId = Date.now().toString();
    pages.push({ id: newId, title: 'Page 1', content: '', mode: 'text', lineModes: ['text'] });
    activePageId = newId;
  } else if (id === activePageId) {
    const nextIndex = Math.min(index, pages.length - 1);
    activePageId = pages[nextIndex].id;
  }

  const next = pages.find(p => p.id === activePageId)!;
  editor.value = next.content;
  previousText = editor.value;

  loadPageModes(next);
  window.mathPopup.setSettings({ pages, activePageId, closedPages });

  updatePageIndicator();
  render();
  editor.focus();
  // Mark for one-shot Ctrl+Z restore (cleared on any edit/navigation).
  justClosedTab = true;
}

function restoreTab(closedIndex: number) {
  justClosedTab = false;
  const page = closedPages.splice(closedIndex, 1)[0];
  const current = pages.find(p => p.id === activePageId);
  if (current) { current.content = editor.value; current.lineModes = [...lineModes]; }

  pages.push(page);
  activePageId = page.id;

  editor.value = page.content;
  previousText = editor.value;

  loadPageModes(page);
  window.mathPopup.setSettings({ pages, activePageId, closedPages });
  
  updatePageIndicator();
  render();
  editor.focus();
}

// Ctrl/Cmd+B / I / U: wrap the selection in markdown markers (bold **, italic *,
// underline __), or strip them if already wrapped (toggle). With no selection,
// drop an empty pair and park the caret between them so you can just start typing.
function toggleInlineFormat(marker: string) {
  const value = editor.value;
  const start = Math.min(editor.selectionStart, editor.selectionEnd);
  const end = Math.max(editor.selectionStart, editor.selectionEnd);
  const sel = value.slice(start, end);
  const m = marker.length;
  captureForUndo();
  if (sel.length >= 2 * m && sel.startsWith(marker) && sel.endsWith(marker)) {
    // Markers sit inside the selection → strip them.
    const inner = sel.slice(m, sel.length - m);
    editor.value = value.slice(0, start) + inner + value.slice(end);
    edSetSelection(start, start + inner.length);
  } else if (value.slice(start - m, start) === marker && value.slice(end, end + m) === marker) {
    // Markers sit just outside the selection → strip them.
    editor.value = value.slice(0, start - m) + sel + value.slice(end + m);
    edSetSelection(start - m, end - m);
  } else {
    // Wrap; keep the original text selected (or park the caret inside if empty).
    editor.value = value.slice(0, start) + marker + sel + marker + value.slice(end);
    edSetSelection(start + m, end + m);
  }
  previousText = editor.value;
  scheduleSave();
  render();
  ensureCaretLineVisible();
}

function onKeyDown(e: KeyboardEvent) {
  // Ctrl+T (new tab), Ctrl+W (close tab), Ctrl+L (rename current tab)
  if (e.ctrlKey && !e.altKey && !e.metaKey && !e.shiftKey) {
    if (e.key === 't' || e.key === 'T') {
      e.preventDefault();
      addTab();
      return;
    }
    if (e.key === 'w' || e.key === 'W') {
      e.preventDefault();
      closeTab(activePageId);
      return;
    }
    if (e.key === 'l' || e.key === 'L') {
      e.preventDefault();
      renameActiveTab();
      return;
    }
  }

  // Ctrl/Cmd+B / I / U → bold / italic / underline the selection (toggle on/off).
  if ((e.ctrlKey || e.metaKey) && !e.altKey && !e.shiftKey) {
    if (e.key === 'b' || e.key === 'B') { e.preventDefault(); toggleInlineFormat('**'); return; }
    if (e.key === 'i' || e.key === 'I') { e.preventDefault(); toggleInlineFormat('*'); return; }
    if (e.key === 'u' || e.key === 'U') { e.preventDefault(); toggleInlineFormat('__'); return; }
  }

  // Ctrl+Shift+T reopens the most recently closed tab (browser-style, and
  // repeatable — each press restores the next one back).
  if (e.ctrlKey && e.shiftKey && !e.altKey && !e.metaKey && (e.key === 't' || e.key === 'T')) {
    e.preventDefault();
    if (closedPages.length > 0) restoreTab(0);
    return;
  }

  // Ctrl+Tab (next tab) and Ctrl+Shift+Tab (prev tab)
  if (e.ctrlKey && !e.altKey && !e.metaKey && (e.key === 'Tab' || e.code === 'Tab')) {
    e.preventDefault();
    if (pages.length <= 1) return;
    const currentIndex = pages.findIndex(p => p.id === activePageId);
    let nextIndex = e.shiftKey ? currentIndex - 1 : currentIndex + 1;
    if (nextIndex < 0) nextIndex = pages.length - 1;
    if (nextIndex >= pages.length) nextIndex = 0;
    switchTab(pages[nextIndex].id);
    return;
  }

  // Ctrl+Z immediately after closing a tab restores it (one-shot), instead of
  // undoing text. The flag is cleared by any edit/navigation, so this only
  // fires when the close was the very last action.
  if ((e.ctrlKey || e.metaKey) && !e.altKey && !e.shiftKey &&
      (e.key === 'z' || e.key === 'Z') && justClosedTab && closedPages.length > 0) {
    e.preventDefault();
    restoreTab(0);
    return;
  }

  // Ctrl+Z / Ctrl+Shift+Z (or Ctrl+Y) undo/redo. We override the textarea's
  // native undo entirely because programmatic edits (smart-tab, auto-format,
  // line-ref shifting, menu inserts) wipe the native history.
  if ((e.ctrlKey || e.metaKey) && !e.altKey && (e.key === 'z' || e.key === 'Z')) {
    e.preventDefault();
    if (e.shiftKey) doRedo(); else doUndo();
    return;
  }
  if ((e.ctrlKey || e.metaKey) && !e.altKey && !e.shiftKey && (e.key === 'y' || e.key === 'Y')) {
    e.preventDefault();
    doRedo();
    return;
  }

  // "/math" or "/text" alone on a line → convert it (the word is stripped) when
  // the user presses Space or Enter.
  if ((e.key === ' ' || e.key === 'Enter') && !e.ctrlKey && !e.altKey && !e.metaKey && !e.shiftKey &&
      handleModeCommand()) {
    e.preventDefault();
    return;
  }

  // The command menu (slash / L popup) eats arrow + Enter + Escape when open.
  if (menuState.open) {
    if (handleMenuKey(e)) return;
  }

  // Word-style list continuation in text mode: Enter after a "- " or "1. " item
  // starts the next item (numbers auto-increment); Enter on an empty item ends
  // the list. Plain Enter only — Shift+Enter stays a normal newline.
  if (currentMode() === 'text' && e.key === 'Enter' &&
      !e.ctrlKey && !e.shiftKey && !e.altKey && !e.metaKey &&
      handleListContinuation()) {
    e.preventDefault();
    return;
  }

  // Tab / Shift+Tab indents or outdents list items in text mode (word-processor
  // style). Falls through to default Tab when the line isn't a list item.
  if (currentMode() === 'text' && e.key === 'Tab' &&
      !e.ctrlKey && !e.altKey && !e.metaKey &&
      handleListIndent(e.shiftKey)) {
    e.preventDefault();
    return;
  }

  // Smart Tab: when the caret sits inside a number, jump past the number to
  // the trailing space (inserting one if missing) and re-run auto-format so
  // any commas are corrected. Math mode only.
  if (currentMode() === 'math' && e.key === 'Tab' &&
      !e.ctrlKey && !e.shiftKey && !e.altKey && !e.metaKey) {
    e.preventDefault();
    handleSmartTab();
    return;
  }

  // Space always inserts a literal space at the caret — even mid-number. Only
  // Tab (handled above) hops past the number. The auto-format pass below still
  // runs so the part left of the caret gets re-commaized as usual.

  // Auto-format / suffix-expansion triggers: space, operators, Enter, comma
  if (currentMode() === 'math' && shouldAutoFormatOnKey(e)) {
    queueMicrotask(() => maybeAutoFormat(e.key));
  }

  // Ctrl+Shift+C: copy current line's result (math mode only)
  if (currentMode() === 'math' && e.ctrlKey && e.shiftKey && (e.key === 'C' || e.key === 'c')) {
    e.preventDefault();
    copyCurrentLineResult();
  }
  // Ctrl+Shift+M: copy whole note as markdown with results inline
  if (e.ctrlKey && e.shiftKey && (e.key === 'M' || e.key === 'm')) {
    e.preventDefault();
    copyAsMarkdown();
  }
  // Esc: hide window. (Tab popups/reorder are handled by a capture-phase
  // Escape listener in bindEvents that runs before this and stops propagation.)
  if (e.key === 'Escape') {
    window.mathPopup.hidePopup();
  }
}

function shouldAutoFormatOnKey(e: KeyboardEvent): boolean {
  if (e.ctrlKey || e.metaKey || e.altKey) return false;
  return e.key === ' ' || e.key === 'Enter' || /[+\-*/^=()]/.test(e.key);
}

// ---- smart Tab ----
// If the caret sits inside (or at either edge of) a number (including a
// trailing custom suffix like `k` or `m`), jump to just past the number,
// ensure there's a trailing space, then run the auto-format pass. Returns
// true if it handled the keystroke.
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

  // Extend `end` past any trailing custom suffix (e.g. "10000k" -> include
  // the k). The user wants the suffix to be treated as part of the number,
  // so smart-tab should not split it off.
  const suffixMatchLen = matchTrailingSuffix(text, end);
  if (suffixMatchLen > 0) {
    end += suffixMatchLen;
  }

  let newText: string;
  let newCaret: number;
  if (text[end] === ' ') {
    newText = text;
    newCaret = end + 1;
  } else {
    newText = text.slice(0, end) + ' ' + text.slice(end);
    newCaret = end + 1;
  }
  if (newText !== text) captureForUndo();
  editor.value = newText;
  editor.selectionStart = editor.selectionEnd = newCaret;
  previousText = editor.value;
  // Re-use the regular auto-format pipeline: it formats the line up to the
  // caret (now sitting after the inserted space) which will recomma the number.
  maybeAutoFormat(' ');
  // maybeAutoFormat early-returns (no render) when no formatting was needed,
  // but we still inserted a space — ensure overlay + save catch up.
  scheduleSave();
  render();
  return true;
}

// Capture state before mutating editor.value programmatically so undo lands
// at a sensible boundary (and not inside a half-applied auto-format).
//
// If a typing burst is in flight, the pre-burst snapshot already captures the
// state we'd want to undo to — adding another snapshot here would split a
// single logical keystroke (e.g. typing space + the autoformat that follows)
// into two undo steps. Skip in that case.
function captureForUndo() {
  if (pendingTypingSnapshot !== null) return;
  pushUndo({ text: editor.value, caretStart: editor.selectionStart, caretEnd: editor.selectionEnd });
}

// If `text` at offset `pos` starts with one of the configured suffix symbols
// AND the suffix isn't followed by another identifier character, return its
// length. Otherwise 0. Used by smart-tab to keep `10000k` intact.
function matchTrailingSuffix(text: string, pos: number): number {
  const suffixes = settings.suffixes ?? [];
  if (!suffixes.length) return 0;
  // Sort longest-first so e.g. "kg" wins over "k".
  const sorted = [...suffixes].sort((a, b) => b.symbol.length - a.symbol.length);
  for (const suf of sorted) {
    const sym = suf.symbol;
    if (!sym) continue;
    const slice = text.slice(pos, pos + sym.length);
    const matches = suf.caseSensitive ? slice === sym : slice.toLowerCase() === sym.toLowerCase();
    if (!matches) continue;
    const after = text[pos + sym.length];
    // Bail when the suffix is followed by another identifier char OR `.` —
    // `.` indicates the user has more number content after (e.g. `5k.5`),
    // and treating `k` as a real suffix there would lead to weird splits.
    if (after !== undefined && /[A-Za-z0-9_.]/.test(after)) continue;
    return sym.length;
  }
  return 0;
}

// ---- Word-style list markers (text mode) ----
// Recognise a leading list marker: a "-", "*" or "+" bullet, or a number with a
// "." or ")" delimiter. A real list item needs a space after the marker ("- x",
// "1. ", "1) x"); the one exception is a bare "1." / "1)" (number + delimiter,
// nothing else), which starts a numbered list. This keeps "-word", "*bold*" and
// decimals like "1.5" from being treated as lists.
interface ListMarker {
  indent: string; bullet?: string; num?: string; delim?: string;
  spaceAfter: string; content: string; markerLen: number;
}
function parseListLine(line: string): ListMarker | null {
  const m = /^(\s*)(?:([-*+])|(\d+)([.)]))(\s*)(.*)$/.exec(line);
  if (!m) return null;
  const [, indent, bullet, num, delim, spaceAfter, content] = m;
  const hasContent = content.trim().length > 0;
  if (spaceAfter.length === 0 && !(num && !hasContent)) return null;
  const markerLen = indent.length + (bullet ? bullet.length : num!.length + delim!.length) + spaceAfter.length;
  return { indent, bullet, num, delim, spaceAfter, content, markerLen };
}

// "/math" or "/text" alone on a line: convert that line's mode and strip the
// command word. Triggered by Space/Enter typed right after the word.
function handleModeCommand(): boolean {
  if (editor.selectionStart !== editor.selectionEnd) return false;
  const text = editor.value;
  const caret = editor.selectionStart;
  const lineStart = text.lastIndexOf('\n', caret - 1) + 1;
  // Match a "/math" or "/text" token ending at the caret — at the line start or
  // after whitespace — so it works mid-line too.
  const m = /(^|\s)\/(math|text)$/i.exec(text.slice(lineStart, caret));
  if (!m) return false;
  const idx = caretLineIndex();
  const tokenStart = lineStart + m.index + m[1].length;   // position of the "/"
  captureForUndo();
  // Strip the command word, keeping everything before it.
  editor.value = text.slice(0, tokenStart) + text.slice(caret);
  editor.selectionStart = editor.selectionEnd = tokenStart;
  previousText = editor.value;
  padLineModes(editor.value);
  lineModes[idx] = m[2].toLowerCase() === 'math' ? 'math' : 'text';
  const activePage = pages.find(p => p.id === activePageId);
  if (activePage) activePage.lineModes = [...lineModes];
  if (menuState.open) hideMenu();
  scheduleSave();
  render();
  return true;
}

// Enter on a list item continues the list (numbers auto-increment, indentation
// preserved); Enter on an empty item ("- " / "1. ") ends it by clearing the
// marker. Returns true if it handled the keystroke (caller suppresses newline).
function handleListContinuation(): boolean {
  if (editor.selectionStart !== editor.selectionEnd) return false;   // selection: default Enter
  const text = editor.value;
  const caret = editor.selectionStart;
  const lineStart = text.lastIndexOf('\n', caret - 1) + 1;
  let lineEnd = text.indexOf('\n', caret);
  if (lineEnd === -1) lineEnd = text.length;

  const mk = parseListLine(text.slice(lineStart, lineEnd));
  if (!mk) return false;
  // Don't hijack Enter typed inside the leading whitespace or the marker itself.
  if (caret < lineStart + mk.markerLen) return false;

  const hasContent = mk.content.trim().length > 0;
  const isEmpty = !hasContent && mk.spaceAfter.length > 0;   // "- " / "1. " with no text

  captureForUndo();
  if (isEmpty) {
    // Clear the empty marker; keep the (now blank) line, caret at its start.
    editor.value = text.slice(0, lineStart) + text.slice(lineEnd);
    editor.selectionStart = editor.selectionEnd = lineStart;
  } else {
    const nextMarker = mk.bullet ? `${mk.bullet} ` : `${parseInt(mk.num!, 10) + 1}${mk.delim} `;
    const insertion = `\n${mk.indent}${nextMarker}`;
    editor.value = text.slice(0, caret) + insertion + text.slice(caret);
    editor.selectionStart = editor.selectionEnd = caret + insertion.length;
  }
  previousText = editor.value;
  scheduleSave();
  render();
  ensureCaretLineVisible();
  return true;
}

// Tab / Shift+Tab on list line(s): indent / outdent by two spaces, like a word
// processor. Operates on every list line the selection spans (or the caret's
// line). Returns true if it handled the keystroke.
const LIST_INDENT = '  ';
function handleListIndent(outdent: boolean): boolean {
  const text = editor.value;
  const selStart = editor.selectionStart;
  const selEnd = editor.selectionEnd;
  const blockStart = text.lastIndexOf('\n', selStart - 1) + 1;
  // A selection ending exactly at a line start shouldn't pull in the next line.
  const endProbe = (selEnd > selStart && selEnd > 0 && text[selEnd - 1] === '\n') ? selEnd - 1 : selEnd;
  let blockEnd = text.indexOf('\n', endProbe);
  if (blockEnd === -1) blockEnd = text.length;

  const original = text.slice(blockStart, blockEnd);
  const lines = original.split('\n');
  if (!lines.some(l => parseListLine(l))) return false;   // nothing list-like to indent

  let firstShift = 0;   // change applied to the first line (for caret math)
  const out = lines.map((l, i) => {
    if (!parseListLine(l)) return l;
    if (outdent) {
      let r = 0;
      if (l[0] === '\t') r = 1; else while (r < LIST_INDENT.length && l[r] === ' ') r++;
      if (i === 0) firstShift = -r;
      return l.slice(r);
    }
    if (i === 0) firstShift = LIST_INDENT.length;
    return LIST_INDENT + l;
  });
  const newBlock = out.join('\n');
  if (newBlock === original) return true;   // outdent past column 0 — swallow Tab, keep focus

  commitTypingBurst();
  captureForUndo();
  editor.value = text.slice(0, blockStart) + newBlock + text.slice(blockEnd);
  if (selStart === selEnd) {
    editor.selectionStart = editor.selectionEnd = Math.max(blockStart, selStart + firstShift);
  } else {
    editor.selectionStart = blockStart;
    editor.selectionEnd = blockStart + newBlock.length;
  }
  previousText = editor.value;
  scheduleSave();
  render();
  ensureCaretLineVisible();
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
  captureForUndo();
  editor.value = head + newLine + tail;
  editor.selectionStart = editor.selectionEnd = newCaretAbs;
  previousText = editor.value;
  scheduleSave();
  render();
}

interface FmtOpts { autoFormatNumbers: boolean; expandSuffixes: boolean; }

export function formatLineForEditor(line: string, suffixes: Suffix[], opts: FmtOpts): { text: string } {
  // Don't format markdown header lines.
  if (/^\s*#{1,6}\s+/.test(line)) return { text: line };

  // Numbers inside quotes are text, not values — never comma-ize or expand
  // them. Split into quoted ("..."/'...') and unquoted spans (capturing group
  // → quotes land on odd indices) and format only the unquoted spans.
  const parts = line.split(/("[^"]*"|'[^']*')/);
  const text = parts
    .map((part, i) => (i % 2 === 1 ? part : formatUnquotedSpan(part, suffixes, opts)))
    .join('');
  return { text };
}

function formatUnquotedSpan(span: string, suffixes: Suffix[], opts: FmtOpts): string {
  let out = span;

  // 0. Bare decimals get a leading zero: ".123" -> "0.123" when preceded by
  //    whitespace, an operator, opening paren, or the start of the line.
  out = out.replace(/(^|[\s+\-*/^=(,])\.(\d)/g, (_m, lead, d) => `${lead}0.${d}`);

  // 1. Expand custom suffixes that come right after a number: "1m" -> "1000000".
  if (opts.expandSuffixes && suffixes.length) {
    const sorted = [...suffixes].sort((a, b) => b.symbol.length - a.symbol.length);
    for (const suf of sorted) {
      const flags = suf.caseSensitive ? 'g' : 'gi';
      const escaped = suf.symbol.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      // Only expand when the suffix is a STANDALONE token (no letter, digit,
      // or `.` immediately after) so we don't munge identifiers like "max",
      // and so a malformed `5k.5` isn't quietly expanded to `5000.5`.
      const re = new RegExp(`(^|[^A-Za-z0-9_])([0-9][0-9,]*(?:\\.[0-9]+)?)${escaped}(?![A-Za-z0-9_.])`, flags);
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
    // Re-comma-ize numbers that already contain commas (incl. ones edited into
    // the wrong places, e.g. "10,000,99"): match the whole digit/comma run and
    // regroup it from scratch. Exclude a leading "(" or "," so we never merge a
    // function's argument list (min(1,2) must stay two args); bare 4+ digit
    // numbers were already handled just above, so a function's first big
    // argument still gets its commas.
    out = out.replace(/(^|[^A-Za-z0-9_.,(])(-?\d[\d,]*(?:\.\d+)?)(?![\d.])/g,
      (_m, lead, num) => `${lead}${commifyNumber(num.replace(/,/g, ''))}`);
  }

  return out;
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
  // Keep per-line modes aligned to the current text (single choke point).
  syncLineModes();
  // Mirror those modes onto the editor's line blocks so math lines reserve the
  // answer column (and text lines use the full width) — the per-line wrap.
  applyEditorLineModes();
  // Evaluate only the math lines; text lines stay inert (no result) but are
  // still classified so the highlighter can style their markdown. Passing the
  // previous results lets the evaluator carry over last-good values mid-edit.
  lastResults = evaluateNote(editor.value, settings.suffixes, lastResults, settings.decimals, lineModes);
  lastCaretLine = document.activeElement === editor ? caretLineIndex() : -1;
  overlay.innerHTML = highlightNote(editor.value, lastResults, lineModes, activeToken, lastCaretLine);
  layoutGutters();
  // Sync scroll AFTER rebuilding the overlay + gutters + result overlay — each
  // innerHTML assignment resets that layer's scrollTop to 0, so syncing earlier
  // would leave them misaligned until the next scroll event (answers/colors
  // appearing to "pop in" only once you scroll).
  syncScroll();
  updateStatus();
  if (findActive) refreshFindAfterEdit();
}

function syncScroll() {
  overlay.scrollTop = editor.scrollTop;
  overlay.scrollLeft = editor.scrollLeft;
  // Keep gutters in vertical sync with the editor's scroll.
  lineGutter.scrollTop = editor.scrollTop;
  resultGutter.scrollTop = editor.scrollTop;
  resultOverlay.scrollTop = editor.scrollTop;
  if (findActive) {
    findLayer.scrollTop = editor.scrollTop;
    findLayer.scrollLeft = editor.scrollLeft;
  }
}

// Build the inline result overlay HTML: a green "= answer" pinned to the right
// edge of each math line (errors render after the "="). Rows are transparent
// and click-through; only the chip itself is interactive (hover to copy).
function resultRowsHTML(heights: number[], caretLine: number): string {
  return heights
    .map((h, i) => {
      const r = lastResults[i];
      const blank = `<div class="row" style="height:${h}px"></div>`;
      if (!r || lineModeAt(i) !== 'math') return blank;
      const activeCls = i === caretLine ? ' active' : '';   // caret on this math line → highlight its answer
      if (r.error) {
        let label = 'error';
        let errCls = 'err-calm';
        let tip = r.errorTooltip ?? r.error ?? 'Error';
        if (r.errorKind === 'incomplete') { label = 'N/A'; errCls = 'err-faint'; tip = r.errorTooltip ?? 'Incomplete expression — keep typing.'; }
        else if (r.errorKind === 'reserved-excel') { label = 'Excel Formula'; tip = r.errorTooltip ?? EXCEL_FORMULA_TOOLTIP; }
        else if (r.errorKind === 'duplicate-var') { label = 'Duplicate'; tip = r.errorTooltip ?? DUPLICATE_VAR_TOOLTIP; }
        else if (r.errorKind === 'reserved-name') { label = 'Reserved'; tip = r.errorTooltip ?? RESERVED_NAME_TOOLTIP; }
        else if (r.errorKind === 'reserved-x') { tip = r.errorTooltip ?? X_RESERVED_TOOLTIP; }
        else if (r.errorKind === 'unquoted-string') { tip = r.errorTooltip ?? UNQUOTED_STRING_TOOLTIP; }
        else if (r.errorKind === 'unknown-var') { label = 'N/A'; errCls = 'err-faint'; tip = r.errorTooltip ?? 'No variable matches that name.'; }
        return `<div class="row" style="height:${h}px"><span class="res ${errCls}${activeCls}" data-tooltip="${escapeAttr(tip)}"><span class="eq">=</span> ${escapeHtml(label)}</span></div>`;
      }
      const txt = r.display ?? '';
      if (txt === '') return blank;
      const copyable = r.numeric !== undefined || r.stringValue !== undefined;
      const iconHtml = copyable ? COPY_ICON_HTML : '';
      const staleCls = r.stale ? ' stale' : '';
      return `<div class="row" style="height:${h}px"><span class="res${staleCls}${activeCls}"><span class="eq">=</span> ${escapeHtml(txt)}${iconHtml}</span></div>`;
    })
    .join('');
}

function layoutGutters() {
  // Rebuilding the gutter/result rows below detaches the elements any hover
  // tooltip is anchored to, so dismiss it now (its mouseleave would never fire).
  hideTooltip();
  const lines = editor.value.split('\n');
  const caretLine = document.activeElement === editor ? caretLineIndex() : -1;
  const editorStyle = getComputedStyle(editor);
  const lineHeight = parseFloat(editorStyle.lineHeight) || 22;

  // ---- Decide the answer-column width BEFORE measuring line heights ----
  // The reserve becomes --editor-reserve, which each MATH line uses as its right
  // padding (.ed-math / .ov-math): math lines wrap before the answer column while
  // text lines keep the full width. The column is exactly as wide as the widest
  // answer — never capped — so the full number shows and the copy icon isn't
  // clipped. First render the chips at placeholder heights to measure the widest.
  resultOverlay.innerHTML = resultRowsHTML(lines.map(() => lineHeight), caretLine);
  let naturalChipW = 0;
  resultOverlay.querySelectorAll<HTMLElement>('.res').forEach(c => {
    if (c.offsetWidth > naturalChipW) naturalChipW = c.offsetWidth;
  });
  const columnW = naturalChipW;
  const GAP = 16;   // breathing room between a wrapped formula and the answer
  const reserve = columnW > 0 ? columnW + GAP : 12;
  document.documentElement.style.setProperty('--editor-reserve', `${reserve}px`);

  // ---- Now read per-line heights, reflecting the reserve just applied. The
  // overlay shares font/padding/width/wrap rules with the editor, so each
  // .ov-line's offsetHeight matches the textarea's visual line height. ----
  const heights: number[] = [];
  const overlayLines = overlay.querySelectorAll<HTMLElement>('.ov-line');
  for (let i = 0; i < lines.length; i++) {
    const el = overlayLines[i];
    const h = el ? el.offsetHeight : lineHeight;
    heights.push(Math.max(lineHeight, h));
  }

  // Sync gutter top/bottom padding with the editor's.
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
    .map((h, i) => {
      const isHl = activeToken?.type === 'lref' && activeToken.line === i + 1;
      const isMath = lineModeAt(i) === 'math';
      const cls = (isHl ? ' hl-lref' : '') + (isMath ? ' math-line' : '') + (i === caretLine ? ' current' : '');
      // Every line shows a faint marker just right of its number (drawn by
      // .line-gutter .row::before): a short elongated dot on a single row, or a
      // bar spanning the rows when it wraps — both the same width.
      return `<div class="row${cls}" style="height:${h}px" data-line="${i + 1}"><span class="lnum">L${i + 1}</span></div>`;
    })
    .join('');
  bindLineGutterTooltips();

  // Final result rows with correct heights, then give every chip the column's
  // min-width so the "=" signs line up vertically. No max-width: a long answer
  // shows in full (and its copy icon stays visible) instead of being clipped.
  resultOverlay.innerHTML = resultRowsHTML(heights, caretLine);
  resultGutter.innerHTML = '';   // results live inline now; the column stays hidden
  if (columnW > 0) {
    resultOverlay.querySelectorAll<HTMLElement>('.res').forEach(c => {
      c.style.minWidth = `${columnW}px`;
    });
  }
  // Re-bind tooltip and click handlers (the chips just got recreated).
  bindResultTooltips();
  bindResultClicks();
}

// Move the caret-line highlight (the result chip `.active` and the gutter
// `.current`) by toggling classes on the EXISTING rows — no rebuild. A caret move
// can't change any answer, so there's no reason to recompute the reserve, re-read
// heights, or recreate the chips; doing that on every click made the whole answer
// column flicker and shift.
function applyCaretHighlight(caretLine: number) {
  resultOverlay.querySelectorAll<HTMLElement>('.res.active').forEach(c => c.classList.remove('active'));
  for (const row of Array.from(lineGutter.children)) (row as HTMLElement).classList.remove('current');
  if (caretLine >= 0) {
    const resRow = resultOverlay.children[caretLine] as HTMLElement | undefined;
    resRow?.querySelector('.res')?.classList.add('active');
    (lineGutter.children[caretLine] as HTMLElement | undefined)?.classList.add('current');
  }
}

// Toggle the gutter's "referenced line" marker for the active L<n> token, again
// without rebuilding the gutter.
function applyActiveLrefHighlight() {
  const children = lineGutter.children;
  for (let i = 0; i < children.length; i++) {
    const isHl = activeToken?.type === 'lref' && activeToken.line === i + 1;
    (children[i] as HTMLElement).classList.toggle('hl-lref', isHl);
  }
}

// Caret moved to a different line. Refresh only what a caret move can affect: the
// syntax overlay (active-token highlight, and a text line's inline ""L#"" ref
// switching between its raw token and the resolved value) and the caret-line
// highlight. Never re-evaluates. Only a text line that carries an inline ref can
// change width when the caret enters/leaves it, so that's the one case that needs
// a real re-layout to keep the result rows aligned.
function refreshCaretLineDisplay() {
  const caretLine = document.activeElement === editor ? caretLineIndex() : -1;
  const prev = lastCaretLine;
  if (caretLine === prev) return;
  lastCaretLine = caretLine;
  overlay.innerHTML = highlightNote(editor.value, lastResults, lineModes, activeToken, caretLine);
  applyEditorConceal(caretLine);
  const lines = editor.value.split('\n');
  const refGeometryChanged = [prev, caretLine].some(
    i => i >= 0 && i < lines.length && lineModeAt(i) === 'text' && /""[lL]\d+""/.test(lines[i])
  );
  // Revealing/concealing markers can change a formatted line's wrap height, so
  // re-layout the gutters when the caret enters or leaves a text line that carries
  // inline markers (cheap test: any * or _).
  const concealShift = (i: number) =>
    i >= 0 && i < lines.length && lineModeAt(i) === 'text' && /[*_]/.test(lines[i]);
  if (refGeometryChanged || concealShift(prev) || concealShift(caretLine)) layoutGutters();
  else applyCaretHighlight(caretLine);
}

function updateStatus() {
  if (currentMode() === 'text') {
    status.textContent = 'Text mode';
    status.className = 'status-msg';
    return;
  }
  // When the user selects across multiple rows, the footer shows sum + avg
  // of the numeric results in the selected range. Falls through to the usual
  // "Ready" / "N errors" message when there's no multi-row selection.
  if (renderSelectionStats()) return;
  const errs = lastResults.filter(r => r.error
    && r.errorKind !== 'reserved-x'
    && r.errorKind !== 'reserved-excel'
    && r.errorKind !== 'reserved-name').length;
  if (errs === 0) {
    status.textContent = 'Ready';
    status.className = 'status-msg ok';
  } else {
    status.textContent = `${errs} error${errs > 1 ? 's' : ''}`;
    status.className = 'status-msg err';
  }
}

// Returns true if the footer was overwritten with selection stats.
// - Single-line selection: evaluates the selected sub-expression and shows "Ans: X".
// - Multi-line selection: shows sum + avg of numeric results in the selected rows.
function renderSelectionStats(): boolean {
  if (currentMode() !== 'math') return false;
  const start = editor.selectionStart;
  const end = editor.selectionEnd;
  if (start === end) return false;
  const text = editor.value;
  const startLine = text.slice(0, start).split('\n').length - 1;
  const endLine = text.slice(0, end).split('\n').length - 1;

  if (startLine === endLine) {
    // Single-line selection: evaluate the highlighted sub-expression.
    const selectedText = text.slice(start, end).trim();
    if (!selectedText) return false;
    const val = evaluateSelectedText(selectedText, lastResults, startLine, settings.suffixes, settings.decimals);
    if (val === undefined || !isFinite(val)) return false;
    status.textContent = `Ans: ${formatResult(val, settings.decimals)}`;
    status.className = 'status-msg ok';
    return true;
  }

  const values: number[] = [];
  for (let i = startLine; i <= endLine; i++) {
    const r = lastResults[i];
    if (r && r.numeric !== undefined && isFinite(r.numeric)) {
      values.push(r.numeric);
    }
  }
  if (values.length === 0) return false;
  const sum = values.reduce((a, b) => a + b, 0);
  const avg = sum / values.length;
  status.textContent = `Sum: ${formatResult(sum, settings.decimals)}  •  Avg: ${formatResult(avg, settings.decimals)}`;
  status.className = 'status-msg ok';
  return true;
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

// ============================================================
// Slash / L command menu
// ============================================================
//
// Triggers:
//   - typing `/` at the start of a line (only whitespace before) opens the
//     slash menu with /no_dec_limit and /clear.
//   - typing `L` not preceded by an identifier char opens the line-ref menu
//     listing every previous line that has a numeric result.
//
// While open, arrows navigate, Enter/Tab confirm, Escape dismisses.
// Typing keeps the menu in sync (filters by the text between trigger char
// and caret). The menu auto-closes if the caret leaves the trigger range,
// the user types whitespace, or selects.

interface SlashCmd {
  /** Text written into the editor (replaces the trigger range). */
  insert: string;
  /** Display label. */
  label: string;
  /** Right-side hint text. */
  hint: string;
  /** Special action handler — when present, replaces the default insert. */
  action?: () => void;
}

interface MenuState {
  open: boolean;
  /** 'slash', 'lineref', or 'varcomp' */
  kind: 'slash' | 'lineref' | 'varcomp' | null;
  /** Index of trigger character in editor.value at trigger time. */
  triggerStart: number;
  items: SlashCmd[];
  filtered: SlashCmd[];
  selectedIdx: number;
}

const menuState: MenuState = {
  open: false,
  kind: null,
  triggerStart: -1,
  items: [],
  filtered: [],
  selectedIdx: 0
};

function buildSlashCommands(): SlashCmd[] {
  return [
    {
      insert: '',
      label: '/math',
      hint: 'Make this line math',
      action: () => setCaretLineMode('math')
    },
    {
      insert: '',
      label: '/text',
      hint: 'Make this line text',
      action: () => setCaretLineMode('text')
    },
    {
      insert: '/no_dec_limit',
      label: '/no_dec_limit',
      hint: 'Up to 6 decimals'
    },
    {
      insert: '',
      label: '/clear',
      hint: 'Clear note',
      action: clearNote
    }
  ];
}

// Set the caret line's mode — backing the /math and /text slash commands.
function setCaretLineMode(mode: Mode) {
  const idx = caretLineIndex();
  padLineModes(editor.value);
  lineModes[idx] = mode;
  const activePage = pages.find(p => p.id === activePageId);
  if (activePage) activePage.lineModes = [...lineModes];
  scheduleSave();
  render();
  editor.focus();
}

function buildVarCompCommands(): SlashCmd[] {
  const out: SlashCmd[] = [];
  const seen = new Set<string>();
  for (const r of lastResults) {
    if (!r.varName || seen.has(r.varName)) continue;
    if (r.errorKind === 'reserved-x' || r.errorKind === 'reserved-excel'
        || r.errorKind === 'reserved-name' || r.errorKind === 'duplicate-var') continue;
    let hint = '';
    if (r.stringValue !== undefined) {
      hint = r.stringValue;
    } else if (r.numeric !== undefined && isFinite(r.numeric)) {
      hint = r.display ?? formatResult(r.numeric, settings.decimals);
    }
    out.push({ insert: r.varName, label: r.varName, hint });
    seen.add(r.varName);
  }
  return out;
}

function buildLineRefCommands(): SlashCmd[] {
  const out: SlashCmd[] = [];
  for (const r of lastResults) {
    if (r.numeric === undefined || !isFinite(r.numeric)) continue;
    if (r.errorKind === 'reserved-x' || r.errorKind === 'reserved-excel'
        || r.errorKind === 'reserved-name') continue;
    const label = `L${r.index + 1}`;
    out.push({
      insert: label,
      label,
      hint: r.display ?? String(r.numeric)
    });
  }
  return out;
}

function clearNote() {
  if (editor.value === '') return;
  captureForUndo();
  editor.value = '';
  editor.selectionStart = editor.selectionEnd = 0;
  previousText = '';
  scheduleSave();
  render();
}

// Decide whether to open / update / close the menu based on the current caret.
// fromClick = true suppresses the varcomp trigger (clicking into the middle of a
// word should highlight it, not open an autocomplete that would duplicate text).
function updateMenuFromCaret(fromClick = false) {
  const caret = editor.selectionStart;
  const text = editor.value;

  // Already open: re-evaluate based on current caret.
  if (menuState.open) {
    // A click while varcomp is open should close it — the user clicked somewhere.
    if (fromClick && menuState.kind === 'varcomp') { hideMenu(); return; }
    const trigger = menuState.triggerStart;
    if (caret < trigger || caret > text.length) { hideMenu(); return; }
    const fragment = text.slice(trigger, caret);
    if (menuState.kind === 'slash') {
      // Must still start with `/` and contain no spaces.
      if (!fragment.startsWith('/') || /\s/.test(fragment)) { hideMenu(); return; }
      filterAndRender(fragment);
      return;
    }
    if (menuState.kind === 'lineref') {
      // Must still start with L/l and only digits after.
      if (!/^L\d*$/i.test(fragment)) { hideMenu(); return; }
      filterAndRender(fragment);
      return;
    }
    if (menuState.kind === 'varcomp') {
      if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(fragment)) { hideMenu(); return; }
      filterAndRender(fragment);
      if (menuState.filtered.length === 0) { hideMenu(); }
      return;
    }
  }

  // Not open: detect a fresh trigger.
  const ch = text[caret - 1];
  if (ch === '/') {
    // Trigger wherever "/" starts a word — line start or right after whitespace
    // — so commands work mid-line too. A "/" following a digit/letter (e.g. the
    // division "5/2") is left alone.
    const before = text[caret - 2];
    if (before !== undefined && !/\s/.test(before)) return;
    openMenu('slash', caret - 1);
    return;
  }
  if ((ch === 'L' || ch === 'l') && currentMode() === 'math') {
    // Only when not part of an existing identifier.
    const prev = text[caret - 2];
    if (prev !== undefined && /[A-Za-z0-9_]/.test(prev)) return;
    // Only when the very next char isn't already a digit (we want to fire
    // on the *first* L typed, not after every L<n> already in place).
    const next = text[caret];
    if (next !== undefined && /\d/.test(next)) return;
    const items = buildLineRefCommands();
    if (items.length === 0) return;
    openMenu('lineref', caret - 1);
    return;
  }

  // Variable completion: only when actively typing in a math line.
  if (!fromClick && currentMode() === 'math') {
    let identStart = caret;
    while (identStart > 0 && /[A-Za-z0-9_]/.test(text[identStart - 1])) identStart--;
    // Must start with a letter/underscore (not a digit mid-number).
    if (identStart < caret && /^[A-Za-z_]/.test(text[identStart])) {
      const fragment = text.slice(identStart, caret);
      const prefix = fragment.toLowerCase();
      const allVars = buildVarCompCommands();
      // Show only when at least one var starts with prefix AND isn't a single exact match.
      const matches = allVars.filter(it => it.insert.startsWith(prefix));
      const hasMore = matches.some(it => it.insert !== prefix);
      if (hasMore) {
        openMenu('varcomp', identStart);
      }
    }
  }
}

function openMenu(kind: 'slash' | 'lineref' | 'varcomp', triggerStart: number) {
  menuState.open = true;
  menuState.kind = kind;
  menuState.triggerStart = triggerStart;
  menuState.items = kind === 'slash' ? buildSlashCommands()
                  : kind === 'lineref' ? buildLineRefCommands()
                  : buildVarCompCommands();
  menuState.selectedIdx = 0;
  const fragment = editor.value.slice(triggerStart, editor.selectionStart);
  // Unhide BEFORE filterAndRender + positionMenu so the items have non-zero
  // offsetHeight when applyMenuMaxHeight measures them. (Hidden elements
  // report offsetHeight 0, which collapsed the menu to a single padding-sized
  // row.) Tuck off-screen first so the user doesn't see a flash.
  cmdMenu.style.left = '-9999px';
  cmdMenu.style.top = '-9999px';
  cmdMenu.hidden = false;
  cmdMenu.setAttribute('aria-hidden', 'false');
  filterAndRender(fragment);
  positionMenu();
}

function hideMenu() {
  menuState.open = false;
  menuState.kind = null;
  menuState.triggerStart = -1;
  menuState.items = [];
  menuState.filtered = [];
  menuState.selectedIdx = 0;
  cmdMenu.hidden = true;
  cmdMenu.setAttribute('aria-hidden', 'true');
}

function filterAndRender(fragment: string) {
  const q = fragment.toLowerCase();
  if (menuState.kind === 'slash') {
    menuState.filtered = menuState.items.filter(it =>
      it.label.toLowerCase().startsWith(q));
  } else if (menuState.kind === 'lineref') {
    // q is like "L" or "L1" or "l12"
    if (q.length <= 1) {
      menuState.filtered = menuState.items.slice();
    } else {
      menuState.filtered = menuState.items.filter(it =>
        it.label.toLowerCase().startsWith(q));
    }
  } else {
    // varcomp: rebuild items so values stay fresh, then prefix-filter.
    menuState.items = buildVarCompCommands();
    menuState.filtered = menuState.items.filter(it => it.insert.startsWith(q));
  }
  if (menuState.filtered.length === 0) {
    // A non-matching slash fragment (e.g. "/2" while typing division) just
    // closes the menu instead of lingering on "No matches".
    if (menuState.kind === 'slash') { hideMenu(); return; }
    cmdMenu.innerHTML = `<div class="cmd-empty">No matches</div>`;
    return;
  }
  if (menuState.selectedIdx >= menuState.filtered.length) {
    menuState.selectedIdx = 0;
  }
  cmdMenu.innerHTML = menuState.filtered.map((it, i) => `
    <div class="cmd-item${i === menuState.selectedIdx ? ' active' : ''}" data-idx="${i}" role="option">
      <span class="cmd-label">${escapeHtml(it.label)}</span>
      <span class="cmd-hint">${escapeHtml(it.hint)}</span>
    </div>
  `).join('');
  // Wire up click handlers (mousedown so blur doesn't kill the click).
  cmdMenu.querySelectorAll<HTMLDivElement>('.cmd-item').forEach(el => {
    el.addEventListener('mousedown', (ev) => {
      ev.preventDefault();
      const idx = Number(el.dataset.idx);
      if (!isNaN(idx)) {
        menuState.selectedIdx = idx;
        confirmMenuSelection();
      }
    });
  });
}

// Cap visible menu height to ~5 items so long lists scroll instead of growing
// past the popup window. Returns the chosen max-height (in pixels).
function applyMenuMaxHeight(): number {
  // Reset any previous cap so we can measure the natural item height.
  cmdMenu.style.maxHeight = '';
  const items = cmdMenu.querySelectorAll<HTMLElement>('.cmd-item');
  const sample = items[0];
  // Default fallback if there are no items (e.g. empty-state row).
  const itemH = sample ? sample.offsetHeight : 28;
  const padding = 8; // approx top+bottom padding of the menu
  const max = itemH * 5 + padding;
  cmdMenu.style.maxHeight = `min(${max}px, calc(100vh - 48px))`;
  return max;
}

function positionMenu() {
  // Compute caret pixel position in viewport coords. cmd-menu is position:
  // fixed and lives at body level, so it can extend outside the editor-stack
  // (no clipping by .editor-stack { overflow: hidden }).
  const editorRect = editor.getBoundingClientRect();
  const coords = caretCoords(menuState.triggerStart);
  const editorStyle = getComputedStyle(editor);
  const lineHeight = parseFloat(editorStyle.lineHeight) || 22;
  const caretViewportTop = editorRect.top + coords.top - editor.scrollTop;
  const caretViewportLeft = editorRect.left + coords.left - editor.scrollLeft;

  applyMenuMaxHeight();
  // Force layout to read offsetHeight after styles applied.
  const menuH = cmdMenu.offsetHeight;
  const menuW = cmdMenu.offsetWidth;

  const margin = 4;
  const spaceBelow = window.innerHeight - (caretViewportTop + lineHeight);
  const spaceAbove = caretViewportTop;

  // Prefer below; flip above when below doesn't fit AND above has more room.
  let top: number;
  if (spaceBelow >= menuH + margin || spaceBelow >= spaceAbove) {
    top = caretViewportTop + lineHeight + 2;
  } else {
    top = caretViewportTop - menuH - 2;
  }
  // Clamp vertical to window
  if (top + menuH > window.innerHeight - margin) top = window.innerHeight - menuH - margin;
  if (top < margin) top = margin;

  let left = caretViewportLeft;
  if (left + menuW > window.innerWidth - margin) {
    left = window.innerWidth - menuW - margin;
  }
  if (left < margin) left = margin;

  cmdMenu.style.top = top + 'px';
  cmdMenu.style.left = left + 'px';
}

// Pixel coordinates of `pos` within the editor, relative to the editor's
// own client box (so top/left are usable directly in styles after offsets).
function caretCoords(pos: number): { top: number; left: number } {
  // The editor is a contenteditable, so we can measure the caret directly: put a
  // collapsed range at `pos` and read its client rect (already accounts for
  // wrapping and scroll). Empty lines have no text rect, so fall back to the
  // line block's own rect.
  const p = edOffsetToPoint(pos);
  const range = document.createRange();
  range.setStart(p.node, p.offset);
  range.collapse(true);
  let rect = range.getBoundingClientRect();
  if (!rect || (rect.top === 0 && rect.left === 0 && rect.height === 0)) {
    const el = p.node.nodeType === Node.TEXT_NODE ? (p.node.parentElement as HTMLElement) : (p.node as HTMLElement);
    if (el && el.getBoundingClientRect) rect = el.getBoundingClientRect();
  }
  const eRect = editor.getBoundingClientRect();
  // Return content-relative coords (add scroll back in) to match what callers
  // like ensureCaretLineVisible expect (they compare against editor.scrollTop).
  return {
    top: rect.top - eRect.top + editor.scrollTop,
    left: rect.left - eRect.left + editor.scrollLeft,
  };
}

function handleMenuKey(e: KeyboardEvent): boolean {
  if (e.key === 'ArrowDown') {
    e.preventDefault();
    moveMenuSelection(1);
    return true;
  }
  if (e.key === 'ArrowUp') {
    e.preventDefault();
    moveMenuSelection(-1);
    return true;
  }
  if (e.key === 'Enter' || e.key === 'Tab') {
    if (menuState.filtered.length > 0) {
      e.preventDefault();
      confirmMenuSelection();
      return true;
    }
    hideMenu();
    return false;
  }
  if (e.key === 'Escape') {
    e.preventDefault();
    hideMenu();
    return true;
  }
  // Any other key: let it through (input handler will re-filter).
  return false;
}

function moveMenuSelection(delta: number) {
  if (menuState.filtered.length === 0) return;
  const n = menuState.filtered.length;
  menuState.selectedIdx = (menuState.selectedIdx + delta + n) % n;
  // Re-render to flip the active class.
  cmdMenu.querySelectorAll<HTMLDivElement>('.cmd-item').forEach((el, i) => {
    el.classList.toggle('active', i === menuState.selectedIdx);
    if (i === menuState.selectedIdx) el.scrollIntoView({ block: 'nearest' });
  });
}

function confirmMenuSelection() {
  const item = menuState.filtered[menuState.selectedIdx];
  if (!item) { hideMenu(); return; }
  if (item.action) {
    // Action commands replace the entire trigger fragment with nothing
    // (the action is responsible for any editor mutation).
    const fragmentEnd = editor.selectionStart;
    const newText = editor.value.slice(0, menuState.triggerStart) + editor.value.slice(fragmentEnd);
    editor.value = newText;
    editor.selectionStart = editor.selectionEnd = menuState.triggerStart;
    previousText = editor.value;
    hideMenu();
    item.action();
    return;
  }
  insertAtTrigger(item.insert);
  hideMenu();
}

function insertAtTrigger(text: string) {
  const start = menuState.triggerStart;
  const end = editor.selectionStart;
  const newValue = editor.value.slice(0, start) + text + editor.value.slice(end);
  if (newValue !== editor.value) captureForUndo();
  editor.value = newValue;
  const caret = start + text.length;
  editor.selectionStart = editor.selectionEnd = caret;
  previousText = editor.value;
  scheduleSave();
  render();
}

// Insert a line reference (L<n>) at the caret — invoked by clicking a line
// number in the gutter. Inside an Excel function call, successive refs are
// auto-separated with a comma so clicking rows builds `SUM(L1, L2, L3)`.
// Anywhere else we drop the bare ref at the caret and let the user supply the
// operators (`5 + ` then click L4 -> `5 + L4`).
function insertLineRefAtCaret(line: number) {
  let start = editor.selectionStart;
  let end = editor.selectionEnd;
  const value = editor.value;

  // If the caret sits right after an empty Excel call — `SUM()|` — hop inside
  // the parentheses so the ref lands where the first argument belongs.
  if (start === end) {
    const m = /([A-Za-z_][A-Za-z0-9_]*)\(\)$/.exec(value.slice(0, start));
    if (m && isExcelFunctionName(m[1])) { start -= 1; end -= 1; }
  }

  const before = value.slice(0, start);
  const sep = needsExcelComma(before) ? ', ' : '';
  const insert = `${sep}L${line}`;
  const newValue = before + insert + value.slice(end);
  if (newValue === value) return;

  captureForUndo();
  editor.value = newValue;
  const caret = start + insert.length;
  editor.selectionStart = editor.selectionEnd = caret;
  previousText = editor.value;
  scheduleSave();
  render();
  editor.focus();
  ensureCaretLineVisible();
}

// True when the caret is inside an Excel function's parentheses AND the
// preceding non-space character ends a value — so the next inserted ref needs a
// leading comma. False right after `(` or `,`, after an operator/range colon,
// or anywhere outside an Excel call (commas are only auto-added "in Excel
// things").
function needsExcelComma(before: string): boolean {
  // Walk backwards to the innermost unclosed '('.
  let depth = 0;
  let openIdx = -1;
  for (let i = before.length - 1; i >= 0; i--) {
    const c = before[i];
    if (c === ')') depth++;
    else if (c === '(') {
      if (depth === 0) { openIdx = i; break; }
      depth--;
    }
  }
  if (openIdx === -1) return false;
  // The identifier immediately before that '(' must be an Excel function name.
  const fn = /([A-Za-z_][A-Za-z0-9_]*)\s*$/.exec(before.slice(0, openIdx));
  if (!fn || !isExcelFunctionName(fn[1])) return false;
  // Only separate when following a value-end token (digit, ref/name char, %,
  // or a closing paren/bracket) — not `(`, `,`, an operator, or a `:` range.
  const trimmed = before.replace(/\s+$/, '');
  return /[A-Za-z0-9_%)\]]/.test(trimmed.charAt(trimmed.length - 1));
}

// ============================================================
// Find (Ctrl+F)
// ============================================================
//
// Chrome-style in-note find. Matches are painted in #find-layer — a layer
// behind the syntax overlay whose own text is transparent, so only the <mark>
// backgrounds show through behind the colored text. The floating bar shows an
// ordinal count (3/12) with prev / next / close, Enter / Shift+Enter and F3 to
// cycle, and Escape to close (handled by the capture-phase Escape listener).

interface FindMatch { start: number; end: number; }
let findActive = false;
let findMatches: FindMatch[] = [];
let findCurrent = -1; // index into findMatches, -1 when none

function openFind() {
  if (findActive) {
    // Already open: just re-focus & select the field, like pressing Ctrl+F twice.
    findInput.focus();
    findInput.select();
    return;
  }
  findActive = true;
  findBar.hidden = false;
  // Prefill from a single-line editor selection, like Chrome.
  const sel = editor.value.slice(editor.selectionStart, editor.selectionEnd);
  if (sel && !sel.includes('\n')) findInput.value = sel;
  autosizeFindInput();
  runFind(findInput.value);
  findInput.focus();
  findInput.select();
}

function closeFind() {
  if (!findActive) return;
  findActive = false;
  findBar.hidden = true;
  findBar.classList.remove('no-results');
  const landing = findCurrent >= 0 ? findMatches[findCurrent] : null;
  findMatches = [];
  findCurrent = -1;
  findLayer.innerHTML = '';
  // Leave the caret on the match the user was viewing, then hand focus back.
  editor.focus();
  if (landing) {
    editor.selectionStart = landing.start;
    editor.selectionEnd = landing.end;
    ensureCaretLineVisible();
  }
}

// Size the find input to its content: a ~6-char resting width that grows as the
// query gets longer, capped so it never runs wide. Measured with a hidden span
// in the input's own font (border-box reset means width includes its padding).
let findMeasureEl: HTMLSpanElement | null = null;
function autosizeFindInput() {
  if (!findMeasureEl) {
    findMeasureEl = document.createElement('span');
    findMeasureEl.setAttribute('aria-hidden', 'true');
    findMeasureEl.style.cssText =
      'position:absolute;visibility:hidden;white-space:pre;top:-9999px;left:-9999px;';
    document.body.appendChild(findMeasureEl);
  }
  const cs = getComputedStyle(findInput);
  findMeasureEl.style.fontFamily = cs.fontFamily;
  findMeasureEl.style.fontSize = cs.fontSize;
  findMeasureEl.style.fontWeight = cs.fontWeight;
  findMeasureEl.style.letterSpacing = cs.letterSpacing;
  findMeasureEl.textContent = '000000'; // ~6-char resting baseline
  const baseline = findMeasureEl.offsetWidth;
  findMeasureEl.textContent = findInput.value;
  const textW = findMeasureEl.offsetWidth;
  // + horizontal padding (4px) and a little caret slack (3px); cap matches CSS.
  const w = Math.max(baseline, textW) + 7;
  findInput.style.width = `${Math.min(240, w)}px`;
}

function computeFindMatches(text: string, query: string): FindMatch[] {
  if (!query) return [];
  const hay = text.toLowerCase();
  const needle = query.toLowerCase();
  const out: FindMatch[] = [];
  for (let i = hay.indexOf(needle); i !== -1; i = hay.indexOf(needle, i + needle.length)) {
    out.push({ start: i, end: i + needle.length });
  }
  return out;
}

function runFind(query: string) {
  findMatches = computeFindMatches(editor.value, query);
  if (findMatches.length === 0) {
    findCurrent = -1;
  } else {
    // Start at the first match at/after the editor caret (where the user was).
    const caret = editor.selectionStart;
    const idx = findMatches.findIndex(m => m.end > caret);
    findCurrent = idx === -1 ? 0 : idx;
  }
  updateFindUi();
  renderFindLayer();
  scrollToCurrentMatch();
}

// Recompute against changed document text while the bar stays open, keeping the
// highlighted ordinal near where it was. Called from render() on any edit.
function refreshFindAfterEdit() {
  const anchor = findCurrent >= 0 ? findMatches[findCurrent].start : editor.selectionStart;
  findMatches = computeFindMatches(editor.value, findInput.value);
  if (findMatches.length === 0) {
    findCurrent = -1;
  } else {
    const idx = findMatches.findIndex(m => m.start >= anchor);
    findCurrent = idx === -1 ? findMatches.length - 1 : idx;
  }
  updateFindUi();
  renderFindLayer();
}

function findNext(dir: 1 | -1) {
  if (findMatches.length === 0) return;
  findCurrent = (findCurrent + dir + findMatches.length) % findMatches.length;
  updateFindUi();
  renderFindLayer();
  scrollToCurrentMatch();
}

function updateFindUi() {
  const total = findMatches.length;
  const hasQuery = findInput.value.length > 0;
  findCount.textContent = total === 0 ? '0/0' : `${findCurrent + 1}/${total}`;
  // A non-empty query with no matches tints the whole bar pale red.
  findBar.classList.toggle('no-results', hasQuery && total === 0);
  findPrevBtn.disabled = total === 0;
  findNextBtn.disabled = total === 0;
}

function renderFindLayer() {
  if (!findActive || findMatches.length === 0) {
    findLayer.innerHTML = '';
    return;
  }
  findLayer.innerHTML = buildFindLayerHtml(editor.value, findMatches, findCurrent);
  findLayer.scrollTop = editor.scrollTop;
  findLayer.scrollLeft = editor.scrollLeft;
}

// Render the whole note as plain (transparent) text with a <mark> around each
// match — no syntax spans, so wrapping a match is a trivial slice, and the
// shared font/wrap rules keep every box aligned with the textarea. The find
// input is single-line, so a match never spans a newline: each belongs to
// exactly one rendered line, and we walk matches with one shared pointer.
function buildFindLayerHtml(text: string, matches: FindMatch[], current: number): string {
  const lines = text.split('\n');
  const out: string[] = [];
  let offset = 0;
  let mi = 0;
  for (let li = 0; li < lines.length; li++) {
    const line = lines[li];
    const lineEnd = offset + line.length;
    let html = '';
    let cursor = offset;
    while (mi < matches.length && matches[mi].start < lineEnd) {
      const m = matches[mi];
      html += escapeHtml(text.slice(cursor, m.start));
      const cls = mi === current ? 'find-hit current' : 'find-hit';
      html += `<mark class="${cls}">${escapeHtml(text.slice(m.start, m.end))}</mark>`;
      cursor = m.end;
      mi++;
    }
    html += escapeHtml(text.slice(cursor, lineEnd));
    // Match the editor's per-line wrap so the highlight boxes stay aligned —
    // including the list hanging indent (see listIndentCols).
    const cls = lineModeAt(li) === 'math' ? 'ov-line ov-math' : 'ov-line';
    const indent = listIndentCols(line);
    const style = indent > 0 ? ` style="padding-left:${indent}ch;text-indent:-${indent}ch"` : '';
    out.push(`<div class="${cls}"${style}>${html || '&#8203;'}</div>`);
    offset = lineEnd + 1;
  }
  return out.join('');
}

function scrollToCurrentMatch() {
  if (!findActive || findCurrent < 0) return;
  const markEl = findLayer.querySelector<HTMLElement>('.find-hit.current');
  if (!markEl) return;
  const top = markEl.offsetTop;
  const bottom = top + markEl.offsetHeight;
  const pad = 28;
  if (top < editor.scrollTop + pad) {
    editor.scrollTop = Math.max(0, top - pad);
  } else if (bottom > editor.scrollTop + editor.clientHeight - pad) {
    editor.scrollTop = bottom - editor.clientHeight + pad;
  }
  syncScroll();
}

// ============================================================
// Undo / Redo
// ============================================================
//
// We override the textarea's native undo because programmatic edits (smart
// tab, auto-format, line-ref shifting, menu inserts) wipe the native history
// and leave it confused. We track snapshots of {text, caret} on:
//   - the START of every typing burst (before the user's first key in a run)
//   - every programmatic mutation, captured BEFORE the change
// Typing bursts group rapid keystrokes into a single undo unit; the burst
// commits when the user pauses (~600ms), presses a "boundary" key (space,
// enter, etc.), or anything programmatic happens.

interface Snapshot { text: string; caretStart: number; caretEnd: number; }

const UNDO_LIMIT = 200;
const TYPING_BURST_MS = 600;
let undoStack: Snapshot[] = [];
let redoStack: Snapshot[] = [];
let pendingTypingSnapshot: Snapshot | null = null;
let typingBurstTimer: number | null = null;

function pushUndo(s: Snapshot) {
  // Drop duplicates (e.g. consecutive captures with no change in between).
  const top = undoStack[undoStack.length - 1];
  if (top && top.text === s.text && top.caretStart === s.caretStart && top.caretEnd === s.caretEnd) return;
  undoStack.push(s);
  if (undoStack.length > UNDO_LIMIT) undoStack.shift();
  redoStack = [];
}

function noteTypingForUndo() {
  // Called from onInput — the value already changed. We want the BEFORE state,
  // which we should have captured via pendingTypingSnapshot when the keystroke
  // was first received. If we missed (e.g. paste, IME), capture the current
  // value as a coarse anchor and move on.
  if (pendingTypingSnapshot === null) {
    pendingTypingSnapshot = {
      text: previousText,
      caretStart: editor.selectionStart,
      caretEnd: editor.selectionEnd
    };
  }
  if (typingBurstTimer) window.clearTimeout(typingBurstTimer);
  typingBurstTimer = window.setTimeout(commitTypingBurst, TYPING_BURST_MS);
}

function commitTypingBurst() {
  if (typingBurstTimer) { window.clearTimeout(typingBurstTimer); typingBurstTimer = null; }
  if (pendingTypingSnapshot && pendingTypingSnapshot.text !== editor.value) {
    pushUndo(pendingTypingSnapshot);
  }
  pendingTypingSnapshot = null;
}

function applySnapshot(s: Snapshot) {
  editor.value = s.text;
  const safeStart = Math.min(s.caretStart, editor.value.length);
  const safeEnd = Math.min(s.caretEnd, editor.value.length);
  editor.selectionStart = safeStart;
  editor.selectionEnd = safeEnd;
  previousText = editor.value;
  scheduleSave();
  render();
  ensureCaretLineVisible();
}

function doUndo() {
  commitTypingBurst();
  if (undoStack.length === 0) return;
  redoStack.push({
    text: editor.value,
    caretStart: editor.selectionStart,
    caretEnd: editor.selectionEnd
  });
  if (redoStack.length > UNDO_LIMIT) redoStack.shift();
  const prev = undoStack.pop()!;
  applySnapshot(prev);
}

function doRedo() {
  commitTypingBurst();
  if (redoStack.length === 0) return;
  undoStack.push({
    text: editor.value,
    caretStart: editor.selectionStart,
    caretEnd: editor.selectionEnd
  });
  if (undoStack.length > UNDO_LIMIT) undoStack.shift();
  const next = redoStack.pop()!;
  applySnapshot(next);
}

// ============================================================
// Hover tooltip for reserved-error rows
// ============================================================

function bindResultTooltips() {
  const chips = resultOverlay.querySelectorAll<HTMLElement>('.res[data-tooltip]');
  chips.forEach(chip => {
    chip.addEventListener('mouseenter', () => showTooltipFor(chip));
    chip.addEventListener('mouseleave', hideTooltip);
  });
}

// Line-number gutter rows use the same custom hover tooltip as result errors
// (instead of the native title attribute).
function bindLineGutterTooltips() {
  lineGutter.querySelectorAll<HTMLElement>('.row[data-line]').forEach(row => {
    row.addEventListener('mouseenter', () => showLineGutterTooltip(row));
    row.addEventListener('mouseleave', hideTooltip);
  });
}

function showLineGutterTooltip(row: HTMLElement) {
  const n = row.dataset.line;
  if (!n) return;
  const isMath = row.classList.contains('math-line');
  const cur = isMath ? 'math' : 'text';
  const other = isMath ? 'text' : 'math';
  const typeSpan = (t: string) => `<span class="tip-type-${t}">${t}</span>`;
  showTooltipHTML(row,
    `<div class="tip-title">L${n} — ${typeSpan(cur)} line</div>` +
    `<div>Single-click: insert this reference</div>` +
    `<div>Double-click: make this ${typeSpan(other)}</div>`,
    650);   // longer hover-intent delay so it doesn't pop up the instant you pass over a number
}

function bindResultClicks() {
  // One row per line (index i aligns with lastResults[i]); only math rows with a
  // value carry a `.res` chip.
  resultOverlay.querySelectorAll<HTMLDivElement>('.row').forEach((row, i) => {
    const chip = row.querySelector<HTMLSpanElement>('.res');
    if (!chip) return;
    const r = lastResults[i];
    if (!r || (r.numeric === undefined && r.stringValue === undefined)) return;
    const val = r.stringValue ?? String(r.numeric);
    const display = r.display ?? val;
    chip.style.cursor = 'pointer';
    chip.addEventListener('click', () => {
      window.mathPopup.copyText(val);
      flashStatus(`Copied ${display}`);
      const iconEl = chip.querySelector<HTMLSpanElement>('.copy-icon');
      if (iconEl) {
        iconEl.innerHTML = COPIED_ICON_HTML;
        iconEl.classList.add('copied');
        setTimeout(() => {
          iconEl.innerHTML = `<svg class="copy-svg" viewBox="0 0 12 12" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="0.75" width="7.25" height="7.25" rx="1.25"/><rect x="0.75" y="4" width="7.25" height="7.25" rx="1.25"/></svg>`;
          iconEl.classList.remove('copied');
        }, 1500);
      }
    });
  });
}

let tooltipShowTimer: number | null = null;

function showTooltipFor(row: HTMLElement) {
  const text = row.dataset.tooltip;
  if (!text) return;
  if (tooltipShowTimer) window.clearTimeout(tooltipShowTimer);
  tooltipShowTimer = window.setTimeout(() => {
    if (!row.isConnected) return;   // row was rebuilt before the delay elapsed
    hoverTooltip.textContent = text;
    hoverTooltip.hidden = false;
    positionTooltipAt(row);
  }, 250);
}

// Like showTooltipFor, but renders rich HTML (used by the line-number tooltip).
function showTooltipHTML(anchor: HTMLElement, html: string, delay = 250) {
  if (tooltipShowTimer) window.clearTimeout(tooltipShowTimer);
  tooltipShowTimer = window.setTimeout(() => {
    if (!anchor.isConnected) return;   // anchor row was rebuilt before the delay elapsed
    hoverTooltip.innerHTML = html;
    hoverTooltip.hidden = false;
    positionTooltipAt(anchor);
  }, delay);
}

// Place the hover tooltip above the anchor (flipping below if there's no room),
// left-aligned with it and clamped to the window.
function positionTooltipAt(anchor: HTMLElement) {
  const rect = anchor.getBoundingClientRect();
  hoverTooltip.style.left = '-9999px';
  hoverTooltip.style.top = '0px';
  const tipRect = hoverTooltip.getBoundingClientRect();
  const padding = 6;
  let top = rect.top - tipRect.height - padding;
  if (top < 4) top = rect.bottom + padding;
  let left = rect.left;
  if (left + tipRect.width + 4 > window.innerWidth) left = window.innerWidth - tipRect.width - 4;
  if (left < 4) left = 4;
  hoverTooltip.style.left = left + 'px';
  hoverTooltip.style.top = top + 'px';
}

function hideTooltip() {
  if (tooltipShowTimer) { window.clearTimeout(tooltipShowTimer); tooltipShowTimer = null; }
  hoverTooltip.hidden = true;
}

// ============================================================
// Signature tooltip (Excel-style intellisense)
// ============================================================
//
// As the user types into an Excel-style function call, a small floating box
// near the caret shows the function signature, the description, and which
// argument they're currently entering (based on commas at depth-0 inside the
// open paren).

const signatureTooltip = document.getElementById('signature-tooltip') as HTMLDivElement;

interface FunctionSig {
  name: string;
  args: string[];
  desc: string;
}

const FUNCTION_SIGNATURES: Record<string, FunctionSig> = {
  sum:     { name: 'SUM',     args: ['value1', '[value2]', '...'], desc: 'Adds all the numbers or line ranges together.' },
  average: { name: 'AVERAGE', args: ['value1', '[value2]', '...'], desc: 'Returns the average (arithmetic mean) of the arguments.' },
  avg:     { name: 'AVG',     args: ['value1', '[value2]', '...'], desc: 'Returns the average (same as AVERAGE).' },
  mean:    { name: 'MEAN',    args: ['value1', '[value2]', '...'], desc: 'Returns the average (same as AVERAGE).' },
  max:     { name: 'MAX',     args: ['value1', '[value2]', '...'], desc: 'Returns the largest value in a set of values.' },
  min:     { name: 'MIN',     args: ['value1', '[value2]', '...'], desc: 'Returns the smallest value in a set of values.' },
  count:   { name: 'COUNT',   args: ['value1', '[value2]', '...'], desc: 'Counts the number of lines/cells that contain numbers.' },
  median:  { name: 'MEDIAN',  args: ['value1', '[value2]', '...'], desc: 'Returns the median (the number in the middle of the set).' },
  round:   { name: 'ROUND',   args: ['number', 'num_digits'],      desc: 'Rounds a number to a specified number of digits.' },
  ceil:    { name: 'CEIL',    args: ['number'],                    desc: 'Rounds a number up to the nearest integer.' },
  floor:   { name: 'FLOOR',   args: ['number'],                    desc: 'Rounds a number down to the nearest integer.' },
  abs:     { name: 'ABS',     args: ['number'],                    desc: 'Returns the absolute value of a number (without its sign).' },
  sqrt:    { name: 'SQRT',    args: ['number'],                    desc: 'Returns the square root of a number.' },
  if:      { name: 'IF',      args: ['logical_test', 'value_if_true', '[value_if_false]'], desc: 'Checks whether a condition is met, and returns one value if TRUE, and another value if FALSE.' },
  today:   { name: 'TODAY',   args: [],                            desc: "Returns today's date as a number." },
  now:     { name: 'NOW',     args: [],                            desc: 'Returns the current date and time as a number.' }
};

function detectActiveSignature(): { sigKey: string; argIndex: number } | null {
  const caret = editor.selectionStart;
  if (caret !== editor.selectionEnd) return null;
  const text = editor.value;
  const lineStart = text.lastIndexOf('\n', caret - 1) + 1;
  const before = text.slice(lineStart, caret);

  // Walk backward from caret, tracking paren depth, until we hit an
  // unmatched `(`. Each top-level `,` we pass while doing so means we've
  // moved past one argument.
  let depth = 0;
  let argCommas = 0;
  let openParenIdx = -1;
  for (let i = before.length - 1; i >= 0; i--) {
    const ch = before[i];
    if (ch === ')') depth++;
    else if (ch === '(') {
      if (depth === 0) { openParenIdx = i; break; }
      depth--;
    } else if (ch === ',' && depth === 0) {
      argCommas++;
    }
  }
  if (openParenIdx === -1) return null;

  // The identifier directly before the `(` (skipping whitespace) is the
  // function name. Match against the Excel signature dictionary.
  let nameEnd = openParenIdx;
  while (nameEnd > 0 && /\s/.test(before[nameEnd - 1])) nameEnd--;
  let nameStart = nameEnd;
  while (nameStart > 0 && /[A-Za-z_]/.test(before[nameStart - 1])) nameStart--;
  if (nameStart === nameEnd) return null;
  const key = before.slice(nameStart, nameEnd).toLowerCase();
  if (!FUNCTION_SIGNATURES[key]) return null;
  return { sigKey: key, argIndex: argCommas };
}

function updateSignatureTooltip() {
  // The slash / L menu always wins for screen real estate near the caret.
  if (menuState.open || currentMode() !== 'math') {
    signatureTooltip.hidden = true;
    return;
  }
  const detected = detectActiveSignature();
  if (!detected) {
    signatureTooltip.hidden = true;
    return;
  }
  const sig = FUNCTION_SIGNATURES[detected.sigKey];
  const argHtml = sig.args.length === 0
    ? '<span class="sig-args">no arguments</span>'
    : sig.args.map((arg, i) => {
        const cls = i === Math.min(detected.argIndex, sig.args.length - 1) ? 'sig-arg-active' : 'sig-args';
        return `<span class="${cls}">${escapeHtml(arg)}</span>`;
      }).join(`<span class="sig-args">, </span>`);
  signatureTooltip.innerHTML = sig.args.length === 0
    ? `<div><span class="sig-name">${escapeHtml(sig.name)}</span><span class="sig-args">()</span></div>` +
      `<div class="sig-desc">${escapeHtml(sig.desc)}</div>`
    : `<div><span class="sig-name">${escapeHtml(sig.name)}</span><span class="sig-args">(</span>${argHtml}<span class="sig-args">)</span></div>` +
      `<div class="sig-desc">${escapeHtml(sig.desc)}</div>`;

  // Position. Prefer above the caret; flip below if no room.
  signatureTooltip.hidden = false;
  signatureTooltip.style.left = '-9999px';
  signatureTooltip.style.top = '0px';
  const editorRect = editor.getBoundingClientRect();
  const editorStyle = getComputedStyle(editor);
  const lineHeight = parseFloat(editorStyle.lineHeight) || 22;
  const coords = caretCoords(editor.selectionStart);
  const caretTop = editorRect.top + coords.top - editor.scrollTop;
  const caretLeft = editorRect.left + coords.left - editor.scrollLeft;
  const tipH = signatureTooltip.offsetHeight;
  const tipW = signatureTooltip.offsetWidth;
  const margin = 4;
  let top = caretTop - tipH - margin;
  if (top < margin) top = caretTop + lineHeight + margin;
  let left = caretLeft;
  if (left + tipW > window.innerWidth - margin) left = window.innerWidth - tipW - margin;
  if (left < margin) left = margin;
  signatureTooltip.style.top = top + 'px';
  signatureTooltip.style.left = left + 'px';
}

function hideSignatureTooltip() { signatureTooltip.hidden = true; }

// ============================================================
// Variables popup (ƒ button)
// ============================================================
//
// Shows every variable assigned in the note (`name = expr`) with its current
// value. Driven entirely off `lastResults` — no separate state to keep in sync.

let varsHideTimer: number | null = null;

function buildVarsList(): { name: string; display: string; line: number }[] {
  const out: { name: string; display: string; line: number }[] = [];
  // Walk in reverse so the LAST assignment to a given name wins (matches the
  // evaluator, which overwrites scope[name] line by line).
  const seen = new Set<string>();
  for (let i = lastResults.length - 1; i >= 0; i--) {
    const r = lastResults[i];
    if (!r.varName || seen.has(r.varName)) continue;
    // Skip rows that errored on a reserved name — those weren't real
    // assignments, just informative pills.
    if (r.errorKind === 'reserved-x' || r.errorKind === 'reserved-excel'
        || r.errorKind === 'reserved-name' || r.errorKind === 'duplicate-var') continue;
    let display: string;
    if (r.stringValue !== undefined) {
      display = r.stringValue;
    } else if (r.numeric !== undefined && isFinite(r.numeric)) {
      display = r.display ?? formatResult(r.numeric, settings.decimals);
    } else {
      continue;
    }
    seen.add(r.varName);
    out.push({ name: r.varName, display, line: r.index + 1 });
  }
  // Restore source order (top-to-bottom of the note).
  out.sort((a, b) => a.line - b.line);
  return out;
}

function renderVarsPopup() {
  const vars = buildVarsList();
  if (vars.length === 0) {
    varsPopup.innerHTML = `<div class="vars-empty">No variables defined yet.</div>`;
    return;
  }
  varsPopup.innerHTML = vars.map(v => `
    <div class="vars-row" data-copy="${escapeAttr(v.display)}">
      <span class="vars-name">${escapeHtml(v.name)}<span class="vars-line">L${v.line}</span></span>
      <span class="vars-value">${escapeHtml(v.display)}</span>
    </div>
  `).join('');
  varsPopup.querySelectorAll<HTMLDivElement>('.vars-row').forEach(row => {
    row.addEventListener('click', () => {
      const val = row.dataset.copy;
      if (val) { window.mathPopup.copyText(val); flashStatus(`Copied ${val}`); }
    });
  });
}

let varsShowTimer: number | null = null;

function scheduleShowVarsPopup() {
  cancelShowVarsPopup();
  varsShowTimer = window.setTimeout(showVarsPopup, HOVER_INTENT_MS);
}

function cancelShowVarsPopup() {
  if (varsShowTimer) { window.clearTimeout(varsShowTimer); varsShowTimer = null; }
}

function showVarsPopup() {
  cancelShowVarsPopup();
  cancelHideVarsPopup();
  renderVarsPopup();
  // Show off-screen first to measure for clamping.
  varsPopup.style.left = '-9999px';
  varsPopup.style.top = '-9999px';
  varsPopup.hidden = false;
  const btnRect = varsBtn.getBoundingClientRect();
  const popRect = varsPopup.getBoundingClientRect();
  const margin = 4;
  let left = btnRect.right - popRect.width;
  if (left < margin) left = margin;
  if (left + popRect.width > window.innerWidth - margin) {
    left = window.innerWidth - popRect.width - margin;
  }
  let top = btnRect.bottom + 4;
  if (top + popRect.height > window.innerHeight - margin) {
    // Flip above when no room below.
    top = btnRect.top - popRect.height - 4;
    if (top < margin) top = margin;
  }
  varsPopup.style.left = left + 'px';
  varsPopup.style.top = top + 'px';
}

function scheduleHideVarsPopup() {
  cancelHideVarsPopup();
  // Small delay so the user can move the cursor from the button onto the
  // popup without it disappearing.
  varsHideTimer = window.setTimeout(hideVarsPopup, 180);
}

function cancelHideVarsPopup() {
  if (varsHideTimer) { window.clearTimeout(varsHideTimer); varsHideTimer = null; }
}

function hideVarsPopup() {
  cancelHideVarsPopup();
  varsPopup.hidden = true;
}

// ============================================================
// Click-to-highlight token references
// ============================================================

function getTokenAtCaret(): ActiveToken | null {
  const caret = editor.selectionStart;
  if (caret !== editor.selectionEnd) return null;
  const text = editor.value;
  let start = caret;
  while (start > 0 && /[A-Za-z0-9_]/.test(text[start - 1])) start--;
  let end = caret;
  while (end < text.length && /[A-Za-z0-9_]/.test(text[end])) end++;
  if (start === end) return null;
  const token = text.slice(start, end);
  if (/^[Ll]\d+$/.test(token)) {
    return { type: 'lref', line: parseInt(token.slice(1)) };
  }
  if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(token)) {
    return { type: 'var', name: token.toLowerCase() };
  }
  return null;
}

function tokenEquals(a: ActiveToken | null, b: ActiveToken | null): boolean {
  if (a === null && b === null) return true;
  if (a === null || b === null) return false;
  if (a.type !== b.type) return false;
  if (a.type === 'var' && b.type === 'var') return a.name === b.name;
  if (a.type === 'lref' && b.type === 'lref') return a.line === b.line;
  return false;
}

function updateActiveToken() {
  if (currentMode() !== 'math') return;
  const newToken = getTokenAtCaret();
  if (tokenEquals(newToken, activeToken)) return;
  activeToken = newToken;
  // Lightweight re-render: refresh the active-token highlight in the overlay and
  // the gutter's referenced-line marker. The answer column is unaffected by which
  // token is active, so it is NOT rebuilt — that avoids the click-time reflow.
  overlay.innerHTML = highlightNote(editor.value, lastResults, lineModes, activeToken, lastCaretLine);
  applyActiveLrefHighlight();
}

document.addEventListener('selectionchange', () => {
  if (document.activeElement === editor) {
    edRefreshSelectionCache();
    updateSignatureTooltip();
    // Refresh footer so sum/avg appears as soon as a multi-row selection is
    // made (and disappears when the selection collapses again).
    updateStatus();
  } else {
    hideSignatureTooltip();
  }
});

// ============================================================
// Wire up keystroke -> undo snapshot capture
// ============================================================

// Capture pre-typing snapshot on the first character of a burst. This fires
// at keydown (before the value changes) so we record the BEFORE state.
editor.addEventListener('keydown', (e) => {
  // Ignore keys that don't produce input on their own (modifiers, navigation,
  // shortcuts already handled in onKeyDown).
  if (e.ctrlKey || e.metaKey || e.altKey) return;
  const isTextInput =
    e.key.length === 1 ||
    e.key === 'Enter' ||
    e.key === 'Backspace' ||
    e.key === 'Delete' ||
    e.key === 'Tab';
  if (!isTextInput) return;
  // Only capture once per burst.
  if (pendingTypingSnapshot === null) {
    pendingTypingSnapshot = {
      text: editor.value,
      caretStart: editor.selectionStart,
      caretEnd: editor.selectionEnd
    };
  }
  // Word boundaries flush the burst so each word/line is its own undo step.
  if (e.key === ' ' || e.key === 'Enter') {
    // Defer commit to next tick so the input event has applied first.
    queueMicrotask(commitTypingBurst);
  }
});

// Copy/cut/paste remember each line's math/text mode, so lifting lines out of one
// page and dropping them on another (anywhere in the app) keeps each line's type.
// We stash {exact text, modes} on copy/cut and re-apply on paste when the pasted
// text matches; pasting unrelated/external text just falls back to normal sync.
let copiedLineModes: { text: string; modes: Mode[] } | null = null;
function rememberCopiedModes(text: string, startOffset: number) {
  const startLine = edValue.slice(0, startOffset).split('\n').length - 1;
  const count = text.split('\n').length;
  const modes: Mode[] = [];
  for (let k = 0; k < count; k++) modes.push(lineModeAt(startLine + k));
  copiedLineModes = { text, modes };
}

// Copy: take over so the clipboard text is exactly edValue's slice (so it matches
// what paste sees), and record the selected lines' modes.
editor.addEventListener('copy', (e) => {
  const s = Math.min(edCaretStart, edCaretEnd);
  const en = Math.max(edCaretStart, edCaretEnd);
  if (s === en) return;                       // nothing selected — let the browser no-op
  const text = edValue.slice(s, en);
  e.preventDefault();
  e.clipboardData?.setData('text/plain', text);
  rememberCopiedModes(text, s);
});

// Paste: insert plain text ourselves (contenteditable would otherwise inject
// HTML), at the current selection, as a single undo step.
editor.addEventListener('paste', (e) => {
  e.preventDefault();
  if (edComposing) return;
  const t = (e.clipboardData?.getData('text/plain') ?? '').replace(/\r\n?/g, '\n');
  const s = Math.min(edCaretStart, edCaretEnd);
  const en = Math.max(edCaretStart, edCaretEnd);
  captureForUndo();
  buildEditorDOM(edValue.slice(0, s) + t + edValue.slice(en));
  edSetSelection(s + t.length, s + t.length);
  previousText = edValue;
  // Restore remembered per-line modes when re-pasting in-app content. Anchor the
  // mode-sync to the new text first (so render()'s syncLineModes won't recompute
  // the pasted lines), then stamp the saved modes onto the pasted range.
  if (copiedLineModes && copiedLineModes.text === t && t.length > 0) {
    syncLineModes();
    const startLine = edValue.slice(0, s).split('\n').length - 1;
    for (let k = 0; k < copiedLineModes.modes.length && startLine + k < lineModes.length; k++) {
      lineModes[startLine + k] = copiedLineModes.modes[k];
    }
    const active = pages.find((p) => p.id === activePageId);
    if (active) active.lineModes = [...lineModes];
  }
  scheduleSave();
  render();
  ensureCaretLineVisible();
});

// Cut: take over (like paste) so the clipboard text exactly matches edValue and we
// can record the lifted lines' modes before removing them.
editor.addEventListener('cut', (e) => {
  if (edComposing) return;
  const s = Math.min(edCaretStart, edCaretEnd);
  const en = Math.max(edCaretStart, edCaretEnd);
  if (s === en) return;                       // nothing selected — let the browser no-op
  e.preventDefault();
  const text = edValue.slice(s, en);
  e.clipboardData?.setData('text/plain', text);
  rememberCopiedModes(text, s);
  commitTypingBurst();
  captureForUndo();
  buildEditorDOM(edValue.slice(0, s) + edValue.slice(en));
  edSetSelection(s, s);
  previousText = edValue;
  scheduleSave();
  render();
  ensureCaretLineVisible();
});

init();
