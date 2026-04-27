import type { ThemePref } from '../shared/types';

async function init() {
  const settings = await window.mathPopup.getSettings();
  applyTheme(settings.theme);
  // CSS handles the dark-mode swap when theme === 'system'; we only need to
  // listen so the channel stays open and any future hooks can react.
  window.mathPopup.onThemeChanged(() => { /* CSS reacts via media query */ });
  setupTocHighlight();
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
