import type { Settings, Suffix } from '../shared/types';
import { DEFAULT_SETTINGS } from '../shared/types';

const launchAtStartupEl = document.getElementById('launch-at-startup') as HTMLInputElement;
const autoFormatEl = document.getElementById('auto-format') as HTMLInputElement;
const expandSuffixesEl = document.getElementById('expand-suffixes') as HTMLInputElement;
const decimalsEl = document.getElementById('decimals') as HTMLInputElement;
const tableBody = document.querySelector('#suffix-table tbody') as HTMLTableSectionElement;
const addBtn = document.getElementById('add-suffix') as HTMLButtonElement;
const resetBtn = document.getElementById('reset-defaults') as HTMLButtonElement;
const saveStatus = document.getElementById('save-status') as HTMLSpanElement;

let settings: Settings;
let dirtyTimer: number | null = null;

async function init() {
  settings = await window.mathPopup.getSettings();
  hydrate();
  bind();
}

function hydrate() {
  launchAtStartupEl.checked = settings.launchAtStartup;
  autoFormatEl.checked = settings.autoFormatNumbers;
  expandSuffixesEl.checked = settings.expandSuffixesInEditor;
  decimalsEl.value = String(settings.decimals);
  renderSuffixRows();
}

function bind() {
  launchAtStartupEl.addEventListener('change', () => save({ launchAtStartup: launchAtStartupEl.checked }));
  autoFormatEl.addEventListener('change', () => save({ autoFormatNumbers: autoFormatEl.checked }));
  expandSuffixesEl.addEventListener('change', () => save({ expandSuffixesInEditor: expandSuffixesEl.checked }));
  decimalsEl.addEventListener('change', () => {
    const v = Math.max(0, Math.min(10, Number(decimalsEl.value) || 2));
    decimalsEl.value = String(v);
    save({ decimals: v });
  });
  addBtn.addEventListener('click', () => {
    settings.suffixes = [...settings.suffixes, { symbol: '', multiplier: 1, caseSensitive: false }];
    renderSuffixRows();
    save({ suffixes: settings.suffixes });
  });
  resetBtn.addEventListener('click', () => {
    settings.suffixes = JSON.parse(JSON.stringify(DEFAULT_SETTINGS.suffixes));
    renderSuffixRows();
    save({ suffixes: settings.suffixes });
  });
}

function renderSuffixRows() {
  tableBody.innerHTML = '';
  settings.suffixes.forEach((suf, i) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><input type="text" data-field="symbol" value="${escapeAttr(suf.symbol)}" /></td>
      <td><input type="number" data-field="multiplier" value="${suf.multiplier}" step="any" /></td>
      <td><input type="checkbox" data-field="caseSensitive" ${suf.caseSensitive ? 'checked' : ''} /></td>
      <td class="actions"><button class="btn subtle danger" data-action="remove">×</button></td>
    `;
    tr.querySelectorAll('input').forEach(input => {
      input.addEventListener('input', () => updateSuffixFromRow(i, tr));
      input.addEventListener('change', () => updateSuffixFromRow(i, tr));
    });
    tr.querySelector('[data-action="remove"]')!.addEventListener('click', () => {
      settings.suffixes.splice(i, 1);
      renderSuffixRows();
      save({ suffixes: settings.suffixes });
    });
    tableBody.appendChild(tr);
  });
}

function updateSuffixFromRow(i: number, tr: HTMLTableRowElement) {
  const symbol = (tr.querySelector('[data-field="symbol"]') as HTMLInputElement).value.trim();
  const multiplier = Number((tr.querySelector('[data-field="multiplier"]') as HTMLInputElement).value);
  const caseSensitive = (tr.querySelector('[data-field="caseSensitive"]') as HTMLInputElement).checked;
  settings.suffixes[i] = { symbol, multiplier: isFinite(multiplier) ? multiplier : 1, caseSensitive };
  save({ suffixes: settings.suffixes });
}

function save(partial: Partial<Settings>) {
  setDirty();
  window.mathPopup.setSettings(partial).then(updated => {
    settings = updated;
    setSaved();
  });
}

function setDirty() {
  saveStatus.textContent = 'Saving…';
  saveStatus.classList.add('dirty');
  if (dirtyTimer) window.clearTimeout(dirtyTimer);
}

function setSaved() {
  if (dirtyTimer) window.clearTimeout(dirtyTimer);
  dirtyTimer = window.setTimeout(() => {
    saveStatus.textContent = 'Saved';
    saveStatus.classList.remove('dirty');
  }, 120);
}

function escapeAttr(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
}

init();
