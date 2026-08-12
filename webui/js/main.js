/* Bootstrap: pull state from Python, wire the shell, pick the opening view. */

import { api, apiReady } from './api.js';
import { hydrate, onChange, state, isFullyTrained } from './state.js';
import { initEditor, focusEditor } from './views/editor.js';
import { initLearn, selectChar } from './views/learn.js';
import { initAI } from './views/ai.js';
import { applyTheme, openSettings, openTelemetryConsent } from './views/settings.js';
import { toastError } from './ui/toast.js';

const views = {
  editor: document.getElementById('view-editor'),
  learn: document.getElementById('view-learn')
};

const navButtons = [...document.querySelectorAll('#nav .segmented__btn')];

function showView(name) {
  Object.entries(views).forEach(([key, el]) => el.classList.toggle('is-active', key === name));
  navButtons.forEach(btn => {
    const active = btn.dataset.view === name;
    btn.classList.toggle('is-active', active);
    btn.setAttribute('aria-selected', String(active));
  });
  if (name === 'editor') focusEditor();
}

navButtons.forEach(btn => btn.addEventListener('click', () => showView(btn.dataset.view)));

/* ── shell buttons ─────────────────────────────────────────────────── */
document.getElementById('btn-settings').addEventListener('click', openSettings);
document.getElementById('btn-website')
  .addEventListener('click', () => api.open_external(state.urls.website));
document.getElementById('btn-status')
  .addEventListener('click', () => api.open_external(state.urls.status));

/* ── banners ───────────────────────────────────────────────────────── */
function renderBanners() {
  const rootBanner = document.getElementById('banner-root');
  if (state.platform.is_root) {
    document.getElementById('banner-root-text').textContent =
      `Files are read from ${state.dataDir}. Running as root is only needed `
      + 'while /dev/uinput is root-owned — a udev rule for the "input" group avoids it.';
    rootBanner.classList.remove('is-hidden');
  } else {
    rootBanner.classList.add('is-hidden');
  }

  const inputBanner = document.getElementById('banner-input');
  if (state.inputError) {
    document.getElementById('banner-input-text').textContent =
      `${state.inputError} Training still works; writing does not.`;
    inputBanner.classList.remove('is-hidden');
  } else {
    inputBanner.classList.add('is-hidden');
  }
}

/* ── nav badge ─────────────────────────────────────────────────────── */
function renderBadge() {
  document.getElementById('nav-learn-badge').textContent =
    `${state.learned.size}/${state.alphabet.length}`;
}

/* ── errors ────────────────────────────────────────────────────────── */
window.addEventListener('error', e => {
  api.report_error(`UI: ${e.message} @ ${e.filename}:${e.lineno}`);
});
window.addEventListener('unhandledrejection', e => {
  const reason = e.reason && e.reason.message ? e.reason.message : String(e.reason);
  api.report_error(`UI promise: ${reason}`);
});

/* ── boot ──────────────────────────────────────────────────────────── */
async function boot() {
  await apiReady;
  let data;
  try {
    data = await api.bootstrap();
  } catch (e) {
    toastError('Could not talk to the backend.');
    throw e;
  }

  hydrate(data);
  applyTheme(state.settings.theme || 'system');

  initEditor({ text: data.text, onGotoLearn: () => showView('learn') });
  initLearn();
  initAI({
    onInsert: text => {
      const field = document.getElementById('input-text');
      field.value = text;
      field.dispatchEvent(new Event('input'));
      showView('editor');
    }
  });

  onChange(() => { renderBadge(); renderBanners(); });
  renderBadge();
  renderBanners();

  // A fresh install lands in training; a trained alphabet lands in the editor.
  showView(isFullyTrained() ? 'editor' : 'learn');
  if (!isFullyTrained()) selectChar(state.currentChar);

  if (data.first_run) openTelemetryConsent();
}

boot();
