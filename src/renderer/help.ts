import type { ThemePref } from '../shared/types';

async function init() {
  const settings = await window.mathPopup.getSettings();
  applyTheme(settings.theme);
  // CSS handles the dark-mode swap when theme === 'system'; we only need to
  // listen so the channel stays open and any future hooks can react.
  window.mathPopup.onThemeChanged(() => { /* CSS reacts via media query */ });
  setupTocHighlight();
  setupHelpSearch();
  setupNativeToc();
}

// Native iOS app only (html gets .cap-native from the web shim): the fixed
// 220px sidebar left no room to read on phones and skinny iPad windows, so the
// TOC becomes a slide-in drawer behind a ☰ button next to the Back control.
// CSS for all of it lives in web.css; desktop and the plain web build never
// have the class and keep the sidebar.
function setupNativeToc() {
  if (!document.documentElement.classList.contains('cap-native')) return;
  const toc = document.querySelector<HTMLElement>('nav.toc');
  if (!toc) return;
  const root = document.documentElement;
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'toc-toggle';
  btn.setAttribute('aria-label', 'Contents');
  btn.textContent = '☰';
  const scrim = document.createElement('div');
  scrim.className = 'toc-scrim';
  const close = () => root.classList.remove('toc-open');
  btn.addEventListener('click', () => root.classList.toggle('toc-open'));
  scrim.addEventListener('click', close);
  // Picking a section closes the drawer so the jump target is visible.
  toc.addEventListener('click', (e) => {
    if ((e.target as HTMLElement).closest('a')) close();
  });
  document.body.append(btn, scrim);
}

function setupHelpSearch() {
  const input = document.getElementById('help-search') as HTMLInputElement;
  const status = document.getElementById('help-search-status') as HTMLSpanElement;
  const sections = Array.from(document.querySelectorAll<HTMLElement>('main section'));
  const tocLinks = Array.from(document.querySelectorAll<HTMLAnchorElement>('.toc a'));

  const search = () => {
    const query = input.value.trim().toLocaleLowerCase();
    let matches = 0;
    for (const section of sections) {
      const match = !query || (section.textContent ?? '').toLocaleLowerCase().includes(query);
      section.hidden = !match;
      if (match) matches++;
      const link = tocLinks.find(a => a.getAttribute('href') === `#${section.id}`);
      if (link) link.hidden = !match;
    }
    status.textContent = query ? `${matches} ${matches === 1 ? 'section' : 'sections'}` : '';
  };

  input.addEventListener('input', search);
  window.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && !e.shiftKey && !e.altKey && (e.key === 'f' || e.key === 'F')) {
      e.preventDefault();
      input.focus();
      input.select();
    }
    if (e.key === 'Escape' && document.activeElement === input && input.value) {
      input.value = '';
      search();
    }
  });
}

function applyTheme(theme: ThemePref) {
  document.documentElement.setAttribute('data-theme', theme);
}

function setupTocHighlight() {
  const links = Array.from(document.querySelectorAll<HTMLAnchorElement>('.toc a'));
  const sections = links
    .map(a => document.querySelector<HTMLElement>(a.getAttribute('href') ?? ''))
    .filter((s): s is HTMLElement => s !== null);

  const observer = new IntersectionObserver(entries => {
    for (const e of entries) {
      if (!e.isIntersecting) continue;
      const id = e.target.id;
      links.forEach(a => a.classList.toggle('active', a.getAttribute('href') === `#${id}`));
    }
  }, { rootMargin: '-20% 0px -70% 0px', threshold: 0 });

  sections.forEach(s => observer.observe(s));
}

init();
