import type { Settings, Suffix, ThemePref } from '../shared/types';
import { DEFAULT_SETTINGS } from '../shared/types';
import type { UpdateState } from '../main/preload';

const launchAtStartupEl = document.getElementById('launch-at-startup') as HTMLInputElement;
const taskbarIconEl = document.getElementById('taskbar-icon') as HTMLInputElement;
const autoFormatEl = document.getElementById('auto-format') as HTMLInputElement;
const expandSuffixesEl = document.getElementById('expand-suffixes') as HTMLInputElement;
const decimalsEl = document.getElementById('decimals') as HTMLInputElement;
const zoomDefaultEl = document.getElementById('zoom-default') as HTMLInputElement;
const tableBody = document.querySelector('#suffix-table tbody') as HTMLTableSectionElement;
const addBtn = document.getElementById('add-suffix') as HTMLButtonElement;
const resetBtn = document.getElementById('reset-defaults') as HTMLButtonElement;
const saveStatus = document.getElementById('save-status') as HTMLSpanElement;
const themePicker = document.getElementById('theme-picker') as HTMLDivElement;
const appVersionEl = document.getElementById('app-version') as HTMLSpanElement;
const checkUpdatesBtn = document.getElementById('check-updates') as HTMLButtonElement;
const installUpdateBtn = document.getElementById('install-update') as HTMLButtonElement;
const updateStatusEl = document.getElementById('update-status') as HTMLSpanElement;
const updateBanner = document.getElementById('update-banner') as HTMLDivElement;
const updateBannerTitle = document.getElementById('update-banner-title') as HTMLDivElement;
const updateBannerProgress = document.getElementById('update-banner-progress') as HTMLDivElement;
const updateBannerProgressFill = document.getElementById('update-banner-progress-fill') as HTMLDivElement;
const updateBannerAction = document.getElementById('update-banner-action') as HTMLButtonElement;

let settings: Settings;
let dirtyTimer: number | null = null;

async function init() {
  settings = await window.mathPopup.getSettings();
  applyTheme(settings.theme);
  hydrate();
  bind();

  window.mathPopup.getAppVersion().then(version => {
    appVersionEl.textContent = version;
  });

  // Reflect whatever update phase the app is already in (e.g. an update was
  // downloaded earlier in this session before the user opened settings).
  window.mathPopup.getUpdateState().then(applyUpdateState);
}

function applyTheme(theme: ThemePref) {
  document.documentElement.setAttribute('data-theme', theme);
  themePicker.querySelectorAll<HTMLButtonElement>('.seg-btn').forEach(btn => {
    const active = btn.dataset.theme === theme;
    btn.classList.toggle('active', active);
    btn.setAttribute('aria-checked', active ? 'true' : 'false');
  });
}

function hydrate() {
  launchAtStartupEl.checked = settings.launchAtStartup;
  taskbarIconEl.checked = settings.showTaskbarIcon;
  autoFormatEl.checked = settings.autoFormatNumbers;
  expandSuffixesEl.checked = settings.expandSuffixesInEditor;
  decimalsEl.value = String(settings.decimals);
  zoomDefaultEl.value = String(Math.round((settings.zoomDefault ?? 1) * 100));
  renderSuffixRows();
}

function bind() {
  launchAtStartupEl.addEventListener('change', () => save({ launchAtStartup: launchAtStartupEl.checked }));
  taskbarIconEl.addEventListener('change', () => save({ showTaskbarIcon: taskbarIconEl.checked }));
  autoFormatEl.addEventListener('change', () => save({ autoFormatNumbers: autoFormatEl.checked }));
  expandSuffixesEl.addEventListener('change', () => save({ expandSuffixesInEditor: expandSuffixesEl.checked }));
  decimalsEl.addEventListener('change', () => {
    const v = Math.max(0, Math.min(10, Number(decimalsEl.value) || 2));
    decimalsEl.value = String(v);
    save({ decimals: v });
  });
  zoomDefaultEl.addEventListener('change', () => {
    const pct = Math.max(50, Math.min(200, Number(zoomDefaultEl.value) || 100));
    // Snap to the nearest 10% for cleaner ratios; the popup also clamps.
    const snapped = Math.round(pct / 10) * 10;
    zoomDefaultEl.value = String(snapped);
    save({ zoomDefault: snapped / 100 });
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
  themePicker.querySelectorAll<HTMLButtonElement>('.seg-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const theme = btn.dataset.theme as ThemePref;
      applyTheme(theme);
      save({ theme });
    });
  });
  // Main pushes a theme:changed event whenever the OS scheme flips. The CSS
  // media query handles the visual swap on its own when theme === 'system';
  // we don't need to do anything here, but the listener keeps the channel open.
  window.mathPopup.onThemeChanged(() => { /* CSS reacts via media query */ });

  checkUpdatesBtn.addEventListener('click', () => {
    checkUpdatesBtn.disabled = true;
    updateStatusEl.textContent = 'Checking...';
    window.mathPopup.checkForUpdates();
  });

  installUpdateBtn.addEventListener('click', () => {
    installUpdateBtn.disabled = true;
    updateStatusEl.textContent = 'Relaunching...';
    window.mathPopup.installUpdate();
  });

  updateBannerAction.addEventListener('click', () => {
    // The banner only ever shows an actionable button in the 'downloaded'
    // phase, so this always means "install & restart now".
    updateBannerAction.disabled = true;
    window.mathPopup.installUpdate();
  });

  window.mathPopup.onUpdateState(applyUpdateState);
}

function applyUpdateState(state: UpdateState) {
  // Banner: only shows for the three "something is happening" phases.
  if (state.phase === 'available' || state.phase === 'downloading' || state.phase === 'downloaded') {
    updateBanner.hidden = false;
    if (state.phase === 'available') {
      updateBannerTitle.textContent = state.version
        ? `Update available (v${state.version}) — preparing download…`
        : 'Update available — preparing download…';
      updateBannerProgress.hidden = true;
      updateBannerProgressFill.style.width = '0%';
      updateBannerAction.hidden = true;
    } else if (state.phase === 'downloading') {
      const pct = Math.max(0, Math.min(100, state.percent ?? 0));
      updateBannerTitle.textContent = `Downloading update — ${pct}%`;
      updateBannerProgress.hidden = false;
      updateBannerProgressFill.style.width = `${pct}%`;
      updateBannerAction.hidden = true;
    } else {
      updateBannerTitle.textContent = state.version
        ? `Update ready (v${state.version})`
        : 'Update ready';
      updateBannerProgress.hidden = true;
      updateBannerProgressFill.style.width = '100%';
      updateBannerAction.hidden = false;
      updateBannerAction.disabled = false;
      updateBannerAction.textContent = 'Restart & Install';
    }
  } else {
    updateBanner.hidden = true;
  }

  // About row: keep the inline status/buttons in sync with the same state.
  if (state.phase === 'checking') {
    updateStatusEl.textContent = 'Checking for updates…';
    checkUpdatesBtn.disabled = true;
    installUpdateBtn.style.display = 'none';
  } else if (state.phase === 'available') {
    updateStatusEl.textContent = 'Update available — downloading…';
    checkUpdatesBtn.style.display = 'none';
    installUpdateBtn.style.display = 'none';
  } else if (state.phase === 'downloading') {
    updateStatusEl.textContent = `Downloading (${state.percent ?? 0}%)…`;
    checkUpdatesBtn.style.display = 'none';
    installUpdateBtn.style.display = 'none';
  } else if (state.phase === 'downloaded') {
    updateStatusEl.textContent = 'Ready to install.';
    checkUpdatesBtn.style.display = 'none';
    installUpdateBtn.style.display = '';
    installUpdateBtn.disabled = false;
  } else if (state.phase === 'not-available') {
    updateStatusEl.textContent = 'App is up to date.';
    checkUpdatesBtn.style.display = '';
    checkUpdatesBtn.disabled = false;
    installUpdateBtn.style.display = 'none';
  } else if (state.phase === 'error') {
    updateStatusEl.textContent = state.error ? `Failed: ${state.error}` : 'Update check failed.';
    checkUpdatesBtn.style.display = '';
    checkUpdatesBtn.disabled = false;
    installUpdateBtn.style.display = 'none';
  } else {
    updateStatusEl.textContent = '';
    checkUpdatesBtn.style.display = '';
    checkUpdatesBtn.disabled = false;
    installUpdateBtn.style.display = 'none';
  }
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
