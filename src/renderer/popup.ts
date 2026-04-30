import { evaluateNote, evaluateSelectedText, LineResult, EXCEL_FORMULA_TOOLTIP, X_RESERVED_TOOLTIP, UNQUOTED_STRING_TOOLTIP, RESERVED_NAME_TOOLTIP, DUPLICATE_VAR_TOOLTIP } from './evaluator';
import { highlightNote, ActiveToken } from './highlighter';
import { formatWithCommas, formatResult } from './formatter';
import type { Mode, Page, Settings, Suffix, ThemePref } from '../shared/types';

const editor = document.getElementById('editor') as HTMLTextAreaElement;
const overlay = document.getElementById('syntax-overlay') as HTMLPreElement;
const measure = document.getElementById('measure') as HTMLDivElement;
const lineGutter = document.getElementById('line-gutter') as HTMLDivElement;
const resultGutter = document.getElementById('result-gutter') as HTMLDivElement;
const status = document.getElementById('status-msg') as HTMLSpanElement;
const closeBtn = document.getElementById('close-window') as HTMLButtonElement;
const settingsBtn = document.getElementById('open-settings') as HTMLButtonElement;
const pinBtn = document.getElementById('toggle-pin') as HTMLButtonElement;
const helpBtn = document.getElementById('open-help') as HTMLButtonElement;
const varsBtn = document.getElementById('show-vars') as HTMLButtonElement;
const varsPopup = document.getElementById('vars-popup') as HTMLDivElement;
const pageIndicator = document.getElementById('page-indicator') as HTMLSpanElement;
const cmdMenu = document.getElementById('cmd-menu') as HTMLDivElement;
const hoverTooltip = document.getElementById('hover-tooltip') as HTMLDivElement;

let settings: Settings;
let closedPages: Page[] = [];
let pages: Page[] = [];
let activePageId: string = '';
let lastResults: LineResult[] = [];
let saveTimer: number | null = null;
let activeToken: ActiveToken | null = null;
const tabsBtn = document.getElementById('tabs-btn') as HTMLButtonElement;
const archiveBtn = document.getElementById('archive-btn') as HTMLButtonElement;
const modeToggleBtn = document.getElementById('mode-toggle-btn') as HTMLButtonElement;
const tabsPopup = document.getElementById('tabs-popup') as HTMLDivElement;
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
    pages.push({ id, title: 'Page 1', content: settings.noteContent ?? '', mode: settings.mode ?? 'math' });
    activePageId = id;
  }
  const activePage = pages.find(p => p.id === activePageId) || pages[0];
  activePageId = activePage.id;
  
  editor.value = activePage.content;
  previousText = editor.value;
  applyTheme(settings.theme);
  applyMode(activePage.mode);
  applyAlwaysOnTop(settings.alwaysOnTop);
  bindEvents();
  // Re-render syntax overlay if the system theme flips while the app is open.
  window.mathPopup.onThemeChanged(() => render());
  updatePageIndicator();
  render();
  editor.focus();
}

function updatePageIndicator() {
  const activePage = pages.find(p => p.id === activePageId);
  if (activePage) pageIndicator.textContent = activePage.title || 'Page';
}

function applyTheme(theme: ThemePref) {
  document.documentElement.setAttribute('data-theme', theme);
}

function bindEvents() {
  editor.addEventListener('input', onInput);
  editor.addEventListener('scroll', syncScroll);
  editor.addEventListener('keydown', onKeyDown);
  editor.addEventListener('blur', () => {
    // Defer so click on the menu can take effect.
    setTimeout(() => {
      if (!cmdMenu.contains(document.activeElement)) hideMenu();
    }, 100);
    hideSignatureTooltip();
  });
  editor.addEventListener('click', () => { updateMenuFromCaret(true); updateActiveToken(); });
  window.addEventListener('resize', () => render());

  closeBtn.addEventListener('click', () => window.mathPopup.hidePopup());
  settingsBtn.addEventListener('click', () => window.mathPopup.openSettings());
  pinBtn.addEventListener('click', toggleAlwaysOnTop);
  helpBtn.addEventListener('click', () => window.mathPopup.openHelp());

  // Dropdowns
  tabsBtn.addEventListener('mouseenter', showTabsPopup);
  tabsBtn.addEventListener('mouseleave', scheduleHideTabsPopup);
  tabsPopup.addEventListener('mouseenter', cancelHideTabsPopup);
  tabsPopup.addEventListener('mouseleave', scheduleHideTabsPopup);

  archiveBtn.addEventListener('mouseenter', showArchivePopup);
  archiveBtn.addEventListener('mouseleave', scheduleHideArchivePopup);
  archivePopup.addEventListener('mouseenter', cancelHideArchivePopup);
  archivePopup.addEventListener('mouseleave', scheduleHideArchivePopup);

  // Variables popup
  varsBtn.addEventListener('mouseenter', showVarsPopup);
  varsBtn.addEventListener('mouseleave', scheduleHideVarsPopup);
  varsBtn.addEventListener('focus', showVarsPopup);
  varsBtn.addEventListener('blur', hideVarsPopup);
  varsPopup.addEventListener('mouseenter', cancelHideVarsPopup);
  varsPopup.addEventListener('mouseleave', scheduleHideVarsPopup);

  modeToggleBtn.addEventListener('click', () => {
    const activePage = pages.find(p => p.id === activePageId);
    if (activePage) setMode(activePage.mode === 'math' ? 'text' : 'math');
  });

  // Listen for settings changes pushed via a polling refresh-on-focus.
  window.addEventListener('focus', async () => {
    settings = await window.mathPopup.getSettings();
    pages = settings.pages || [];
    if (!pages.find(p => p.id === activePageId) && pages.length > 0) {
      activePageId = settings.activePageId || pages[0].id;
      const activePage = pages.find(p => p.id === activePageId) || pages[0];
      editor.value = activePage.content;
      previousText = editor.value;
      applyMode(activePage.mode);
    }
    applyTheme(settings.theme);
    applyAlwaysOnTop(settings.alwaysOnTop);
    render();
  });
}

function setMode(mode: Mode) {
  const activePage = pages.find(p => p.id === activePageId);
  if (!activePage || activePage.mode === mode) return;
  activePage.mode = mode;
  applyMode(mode);
  window.mathPopup.setSettings({ pages, activePageId, closedPages });
  render();
  editor.focus();
}

function applyMode(mode: Mode) {
  document.body.classList.toggle('text-mode', mode === 'text');
  modeToggleBtn.title = mode === 'math' ? 'Mode: Math' : 'Mode: Text';
  modeToggleBtn.innerHTML = mode === 'math' ? '∑' : 'Aa';
}

// Per-tab mode lives on activePage.mode. settings.mode is a legacy/global
// fallback that's no longer reliable (the focus handler reloads it from disk
// and we never persist it on toggle), so always read the active page.
function currentMode(): Mode {
  return pages.find(p => p.id === activePageId)?.mode ?? settings?.mode ?? 'math';
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
  const previousToken = activeToken;
  activeToken = null;
  if (currentMode() === 'math') {
    maybeSyncRename(previousToken);
    maybeShiftLineRefs();
  }
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

// Synchronously apply edits made to the active token to all of its other occurrences.
function maybeSyncRename(previousToken: ActiveToken | null) {
  if (!previousToken || previousToken.type !== 'var') return;

  const oldText = previousText;
  const newText = editor.value;
  if (oldText === newText) return;

  // 1. Find the single contiguous edit.
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

  // 2. Find all occurrences of the variable in the old text.
  const occurrences: {start: number, end: number}[] = [];
  const escapedName = previousToken.name.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\$&');
  const re = new RegExp(`(^|[^A-Za-z0-9_])${escapedName}(?![A-Za-z0-9_])`, 'gi');
  let m;
  while ((m = re.exec(oldText)) !== null) {
    const matchStart = m.index + m[1].length;
    occurrences.push({ start: matchStart, end: matchStart + previousToken.name.length });
  }

  if (occurrences.length <= 1) return;

  // 3. Ensure the edit is fully contained within exactly one occurrence.
  const editedOccIdx = occurrences.findIndex(occ => occ.start <= editStart && oldEditEnd <= occ.end);
  if (editedOccIdx === -1) return;

  const editedOcc = occurrences[editedOccIdx];

  // 4. Ensure we are editing the "base" definition, not a reference.
  const lineIndex = (oldText.slice(0, editedOcc.start).match(/\n/g) || []).length;
  const lineResult = lastResults[lineIndex];
  if (!lineResult || !lineResult.varName || lineResult.varName.toLowerCase() !== previousToken.name.toLowerCase()) {
    return; // The line doesn't define this variable.
  }

  const lineStart = oldText.lastIndexOf('\n', editedOcc.start - 1) + 1;
  const textBeforeOccOnLine = oldText.slice(lineStart, editedOcc.start);
  const isFirstOnLine = !new RegExp(`(^|[^A-Za-z0-9_])${previousToken.name.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\\\$&')}(?![A-Za-z0-9_])`, 'i').test(textBeforeOccOnLine);
  if (!isFirstOnLine) {
    return; // We are editing a reference that happens to be on the same line as the definition.
  }

  const relStart = editStart - editedOcc.start;
  const relEnd = oldEditEnd - editedOcc.start;
  const delta = insertedText.length - deletedText.length;

  let outText = oldText;
  let newCaretStart = editor.selectionStart;
  let newCaretEnd = editor.selectionEnd;

  // Apply the exact same edit to all occurrences, from right to left to preserve offsets.
  for (let i = occurrences.length - 1; i >= 0; i--) {
    const occ = occurrences[i];

    // Do not mirror edits to OTHER base variable definitions.
    if (i !== editedOccIdx) {
      const occLineIndex = (oldText.slice(0, occ.start).match(/\n/g) || []).length;
      const occLineResult = lastResults[occLineIndex];
      const isBaseLine = occLineResult && occLineResult.varName && occLineResult.varName.toLowerCase() === previousToken.name.toLowerCase();
      
      if (isBaseLine) {
        const occLineStart = oldText.lastIndexOf('\n', occ.start - 1) + 1;
        const occTextBefore = oldText.slice(occLineStart, occ.start);
        const occIsFirst = !new RegExp(`(^|[^A-Za-z0-9_])${escapedName}(?![A-Za-z0-9_])`, 'i').test(occTextBefore);
        if (occIsFirst) {
          continue; // Skip applying edit to this other base definition
        }
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

  // Keep the token active if it still looks like a valid variable,
  // so the user can continue typing continuously.
  const newName = oldText.slice(editedOcc.start, editedOcc.start + relStart) + insertedText + oldText.slice(editedOcc.start + relEnd, editedOcc.end);
  if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(newName)) {
    activeToken = { type: 'var', name: newName.toLowerCase() };
  } else {
    activeToken = null;
  }
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
}

function scheduleSave() {
  if (saveTimer) window.clearTimeout(saveTimer);
  saveTimer = window.setTimeout(() => {
    const activePage = pages.find(p => p.id === activePageId);
    if (activePage) activePage.content = editor.value;
    window.mathPopup.setSettings({ pages, activePageId });
  }, 250);
}

let tabsHoverTimer: number | null = null;
let archiveHoverTimer: number | null = null;

function showTabsPopup() {
  if (tabsHoverTimer) window.clearTimeout(tabsHoverTimer);
  renderTabsMenu();
  tabsPopup.hidden = false;
  const rect = tabsBtn.getBoundingClientRect();
  tabsPopup.style.top = `${rect.bottom + 4}px`;
  tabsPopup.style.left = `${rect.left}px`;
}
function scheduleHideTabsPopup() { tabsHoverTimer = window.setTimeout(hideTabsPopup, 150); }
function cancelHideTabsPopup() { if (tabsHoverTimer) window.clearTimeout(tabsHoverTimer); }
function hideTabsPopup() { tabsPopup.hidden = true; }

function showArchivePopup() {
  if (archiveHoverTimer) window.clearTimeout(archiveHoverTimer);
  renderArchiveMenu();
  archivePopup.hidden = false;
  const rect = archiveBtn.getBoundingClientRect();
  archivePopup.style.top = `${rect.bottom + 4}px`;
  archivePopup.style.left = `${rect.left}px`;
}
function scheduleHideArchivePopup() { archiveHoverTimer = window.setTimeout(hideArchivePopup, 150); }
function cancelHideArchivePopup() { if (archiveHoverTimer) window.clearTimeout(archiveHoverTimer); }
function hideArchivePopup() { archivePopup.hidden = true; }

function renderTabsMenu() {
  tabsPopup.innerHTML = '';
  pages.forEach((page, index) => {
    const row = document.createElement('div');
    row.className = 'vars-row' + (page.id === activePageId ? ' active' : '');
    row.style.cursor = 'pointer';
    row.onclick = (e) => {
      // Don't trigger if clicking inside the input
      if ((e.target as HTMLElement).tagName === 'INPUT') return;
      switchTab(page.id); 
      hideTabsPopup(); 
    };
    
    const titleSpan = document.createElement('span');
    titleSpan.className = 'vars-name';
    titleSpan.textContent = page.title || `Page ${index + 1}`;
    
    const controls = document.createElement('div');
    controls.style.display = 'flex';
    controls.style.gap = '4px';
    
    const editBtn = document.createElement('span');
    editBtn.className = 'vars-val';
    editBtn.textContent = '✏️';
    editBtn.style.cursor = 'pointer';
    editBtn.onclick = (e) => {
      e.stopPropagation();
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
        renderTabsMenu();
      };
      
      input.onkeydown = (e2) => {
        if (e2.key === 'Enter') {
          e2.preventDefault();
          saveName();
        }
      };
      input.onblur = saveName;
      
      row.replaceChild(input, titleSpan);
      input.focus();
    };

    const closeBtn = document.createElement('span');
    closeBtn.className = 'vars-val';
    closeBtn.textContent = '×';
    closeBtn.style.cursor = 'pointer';
    closeBtn.onclick = (e) => {
      e.stopPropagation();
      closeTab(page.id);
      renderTabsMenu();
    };
    
    controls.appendChild(editBtn);
    controls.appendChild(closeBtn);
    row.appendChild(titleSpan);
    row.appendChild(controls);
    tabsPopup.appendChild(row);
  });
  
  if (pages.length < 99) {
    const footer = document.createElement('div');
    footer.className = 'popup-footer';
    const addBtn = document.createElement('button');
    addBtn.className = 'popup-btn';
    addBtn.textContent = '+ New Tab';
    addBtn.onclick = () => { addTab(); hideTabsPopup(); };
    footer.appendChild(addBtn);
    tabsPopup.appendChild(footer);
  }
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
  const current = pages.find(p => p.id === activePageId);
  if (current) current.content = editor.value;
  
  activePageId = id;
  const next = pages.find(p => p.id === activePageId)!;
  editor.value = next.content;
  previousText = editor.value;
  
  applyMode(next.mode);
  window.mathPopup.setSettings({ pages, activePageId, closedPages });
  
  updatePageIndicator();
  render();
  editor.focus();
}

function addTab() {
  if (pages.length >= 99) return;
  const current = pages.find(p => p.id === activePageId);
  if (current) current.content = editor.value;
  
  const id = Date.now().toString();
  const title = `Page ${pages.length + 1}`;
  pages.push({ id, title, content: '', mode: 'math' });
  activePageId = id;
  
  editor.value = '';
  previousText = '';
  applyMode('math');
  window.mathPopup.setSettings({ pages, activePageId, closedPages });
  
  updatePageIndicator();
  render();
  editor.focus();
}

function closeTab(id: string) {
  const index = pages.findIndex(p => p.id === id);
  if (index === -1) return;
  
  const current = pages[index];
  if (id === activePageId) current.content = editor.value;
  closedPages.unshift(current);
  if (closedPages.length > 10) closedPages.pop();
  
  pages.splice(index, 1);
  if (pages.length === 0) {
    const newId = Date.now().toString();
    pages.push({ id: newId, title: 'Page 1', content: '', mode: 'math' });
    activePageId = newId;
  } else if (id === activePageId) {
    const nextIndex = Math.min(index, pages.length - 1);
    activePageId = pages[nextIndex].id;
  }
  
  const next = pages.find(p => p.id === activePageId)!;
  editor.value = next.content;
  previousText = editor.value;
  
  applyMode(next.mode);
  window.mathPopup.setSettings({ pages, activePageId, closedPages });
  
  updatePageIndicator();
  render();
  editor.focus();
}

function restoreTab(closedIndex: number) {
  const page = closedPages.splice(closedIndex, 1)[0];
  const current = pages.find(p => p.id === activePageId);
  if (current) current.content = editor.value;
  
  pages.push(page);
  activePageId = page.id;
  
  editor.value = page.content;
  previousText = editor.value;
  
  applyMode(page.mode);
  window.mathPopup.setSettings({ pages, activePageId, closedPages });
  
  updatePageIndicator();
  render();
  editor.focus();
}

function onKeyDown(e: KeyboardEvent) {
  // Ctrl+T (new tab) and Ctrl+W (close tab)
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

  // The command menu (slash / L popup) eats arrow + Enter + Escape when open.
  if (menuState.open) {
    if (handleMenuKey(e)) return;
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
  // Esc: hide window (only if no menu is open — the menu intercepts above)
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
  const mode = currentMode();
  if (mode === 'math') {
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
  overlay.innerHTML = highlightNote(editor.value, lastResults, mode, activeToken);
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
    const text = lines[i].length === 0 ? ' ' : lines[i];
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
    .map((h, i) => {
      const isHl = activeToken?.type === 'lref' && activeToken.line === i + 1;
      const cls = isHl ? ' hl-lref' : '';
      return `<div class="row${cls}" style="height:${h}px">L${i + 1}</div>`;
    })
    .join('');

  // Build result column
  resultGutter.innerHTML = heights
    .map((h, i) => {
      const r = lastResults[i];
      if (!r) return `<div class="row empty" style="height:${h}px"></div>`;
      if (r.error) {
        if (r.errorKind === 'reserved-x') {
          // Custom hover tooltip only — no `title` attribute, otherwise the
          // native OS tooltip stacks on top of our styled one.
          const tip = escapeAttr(r.errorTooltip ?? X_RESERVED_TOOLTIP);
          return `<div class="row link-error" style="height:${h}px" data-tooltip="${tip}">N/A</div>`;
        }
        if (r.errorKind === 'reserved-excel') {
          const tip = escapeAttr(r.errorTooltip ?? EXCEL_FORMULA_TOOLTIP);
          return `<div class="row link-error excel" style="height:${h}px" data-tooltip="${tip}">Excel Formula</div>`;
        }
        if (r.errorKind === 'reserved-name') {
          const tip = escapeAttr(r.errorTooltip ?? RESERVED_NAME_TOOLTIP);
          return `<div class="row link-error" style="height:${h}px" data-tooltip="${tip}">Reserved</div>`;
        }
        if (r.errorKind === 'unquoted-string') {
          const tip = escapeAttr(r.errorTooltip ?? UNQUOTED_STRING_TOOLTIP);
          return `<div class="row error" style="height:${h}px" data-tooltip="${tip}">err</div>`;
        }
        if (r.errorKind === 'duplicate-var') {
          const tip = escapeAttr(r.errorTooltip ?? DUPLICATE_VAR_TOOLTIP);
          return `<div class="row link-error" style="height:${h}px" data-tooltip="${tip}">Duplicate</div>`;
        }
        // Other general errors (parse errors, etc.): render as blank in the
        // gutter. The line highlighter already shows a red underline/bg on
        // the offending line, so an additional "err" pill is just noise
        // while the user is mid-typing.
        return `<div class="row empty" style="height:${h}px"></div>`;
      }
      const txt = r.display ?? '';
      let cls = 'row';
      if (txt === '') cls = 'row empty';
      else if (r.stale) cls = 'row stale';
      const copyable = txt !== '' && (r.numeric !== undefined || r.stringValue !== undefined);
      const iconHtml = copyable ? COPY_ICON_HTML : '';
      return `<div class="${cls}" style="height:${h}px">${escapeHtml(txt)}${iconHtml}</div>`;
    })
    .join('');
  // Re-bind tooltip and click handlers (the rows just got recreated).
  bindResultTooltips();
  bindResultClicks();
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
  if (currentMode() !== 'math') {
    if (menuState.open) hideMenu();
    return;
  }
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
    // Only when the slash is the first non-whitespace on its line.
    const lineStart = text.lastIndexOf('\n', caret - 2) + 1;
    const before = text.slice(lineStart, caret - 1);
    if (!/^\s*$/.test(before)) return;
    openMenu('slash', caret - 1);
    return;
  }
  if (ch === 'L' || ch === 'l') {
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

  // Variable completion: only when actively typing (not on click).
  if (!fromClick) {
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
  const text = editor.value;
  const before = text.slice(0, pos);
  // Build measure content: text + a marker span at the caret.
  measure.style.display = 'block';
  measure.style.visibility = 'hidden';
  measure.textContent = '';
  const pre = document.createTextNode(before);
  const marker = document.createElement('span');
  marker.textContent = '​';
  measure.appendChild(pre);
  measure.appendChild(marker);
  const mRect = marker.getBoundingClientRect();
  const eRect = editor.getBoundingClientRect();
  const top = mRect.top - eRect.top;
  const left = mRect.left - eRect.left;
  measure.textContent = '';
  measure.style.display = '';
  return { top, left };
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
  const rows = resultGutter.querySelectorAll<HTMLDivElement>('.row[data-tooltip]');
  rows.forEach(row => {
    row.addEventListener('mouseenter', () => showTooltipFor(row));
    row.addEventListener('mouseleave', hideTooltip);
  });
}

function bindResultClicks() {
  // Select ALL rows (including empty) so index i aligns with lastResults[i].
  resultGutter.querySelectorAll<HTMLDivElement>('.row').forEach((row, i) => {
    if (row.classList.contains('empty')) return;
    const r = lastResults[i];
    if (!r || (r.numeric === undefined && r.stringValue === undefined)) return;
    const val = r.stringValue ?? String(r.numeric);
    const display = r.display ?? val;
    row.style.cursor = 'pointer';
    row.addEventListener('click', () => {
      window.mathPopup.copyText(val);
      flashStatus(`Copied ${display}`);
      const iconEl = row.querySelector<HTMLSpanElement>('.copy-icon');
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

function showTooltipFor(row: HTMLDivElement) {
  const text = row.dataset.tooltip;
  if (!text) return;
  if (tooltipShowTimer) window.clearTimeout(tooltipShowTimer);
  tooltipShowTimer = window.setTimeout(() => {
    hoverTooltip.textContent = text;
    hoverTooltip.hidden = false;
    // Position above the row, horizontally aligned with its left edge but
    // clamped to the window.
    const rect = row.getBoundingClientRect();
    // Show first to measure
    hoverTooltip.style.left = '-9999px';
    hoverTooltip.style.top = '0px';
    const tipRect = hoverTooltip.getBoundingClientRect();
    const padding = 6;
    let top = rect.top - tipRect.height - padding;
    if (top < 4) top = rect.bottom + padding; // flip below if no room above
    let left = rect.left;
    if (left + tipRect.width + 4 > window.innerWidth) {
      left = window.innerWidth - tipRect.width - 4;
    }
    if (left < 4) left = 4;
    hoverTooltip.style.left = left + 'px';
    hoverTooltip.style.top = top + 'px';
  }, 250);
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

function showVarsPopup() {
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
  // Lightweight re-render: just overlay + gutters, no re-evaluation needed.
  overlay.innerHTML = highlightNote(editor.value, lastResults, currentMode(), activeToken);
  layoutGutters();
}

document.addEventListener('selectionchange', () => {
  if (document.activeElement === editor) {
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

// Paste / cut should also be captured.
editor.addEventListener('paste', () => commitTypingBurst());
editor.addEventListener('cut', () => commitTypingBurst());

init();
