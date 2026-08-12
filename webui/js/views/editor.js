/* Write view: the text field, the run controls and the live drawing status. */

import { api } from '../api.js';
import { state, setDrawing, setLearned } from '../state.js';
import { toast, toastOk, toastError, toastWarn } from '../ui/toast.js';
import { confirm } from '../ui/modal.js';
import { pickOriginFromFrame } from '../ui/originPicker.js';

const els = {
  text: document.getElementById('input-text'),
  count: document.getElementById('meta-count'),
  unknown: document.getElementById('meta-unknown'),
  pill: document.getElementById('statuspill'),
  status: document.getElementById('status-text'),
  progress: document.getElementById('progress'),
  fill: document.getElementById('progress-fill'),
  pct: document.getElementById('progress-pct'),
  char: document.getElementById('progress-char'),
  draw: document.getElementById('btn-draw'),
  stop: document.getElementById('btn-stop'),
  reset: document.getElementById('btn-reset'),
  gotoLearn: document.getElementById('btn-goto-learn'),
  quickBtn: document.getElementById('quick-ai-btn'),
  quickMenu: document.getElementById('quick-ai-menu')
};

let saveTimer = null;
let checkTimer = null;
let onNavigateToLearn = () => {};

function setStatus(text, kind = '') {
  els.status.textContent = text;
  els.pill.className = 'statuspill' + (kind ? ` is-${kind}` : '');
}

function countdownMessage() {
  // On Linux the start point was already picked from a screenshot; on
  // Windows, OneNote is found and focused automatically.
  const isLinux = state.platform && state.platform.linux;
  return isLinux ? 'Get ready' : 'Focus OneNote';
}

/* On Wayland the real cursor position can't be read reliably, so instead of
   "hover the mouse there" the user clicks the spot on a live screenshot. */
function waitForOriginFrame() {
  return new Promise(resolve => window.T2P.on('origin:frame', resolve));
}

async function pickLinuxOrigin() {
  setStatus('Taking a screenshot…', 'busy');
  const started = await api.pick_origin();
  if (!started.ok) {
    toastError(started.message || 'Could not start the screen picker.');
    setStatus('Ready', 'ready');
    return null;
  }
  const frame = await waitForOriginFrame();
  if (!frame.ok) {
    toastError(frame.message || 'Could not capture the screen.');
    setStatus('Ready', 'ready');
    return null;
  }
  const point = await pickOriginFromFrame(frame);
  if (!point) {
    setStatus('Ready', 'ready');
    return null;
  }
  return point;
}

function setProgress(value, char = '') {
  const pct = Math.round(Math.max(0, Math.min(1, value)) * 100);
  els.fill.style.width = `${pct}%`;
  els.pct.textContent = `${pct}%`;
  els.char.textContent = char ? `drawing “${char}”` : '';
}

/* ── text field ────────────────────────────────────────────────────── */
async function refreshMeta() {
  const text = els.text.value;
  const res = await api.check_text(text);
  els.count.textContent = `${res.total_chars} character${res.total_chars === 1 ? '' : 's'}`;
  if (res.unknown && res.unknown.length) {
    els.unknown.textContent = `Not trained: ${res.unknown.join(' ')}`;
    els.unknown.className = 'meta-warn';
  } else {
    els.unknown.textContent = 'All characters trained';
    els.unknown.className = 'meta-ok';
  }
}

els.text.addEventListener('input', () => {
  clearTimeout(saveTimer);
  clearTimeout(checkTimer);
  saveTimer = setTimeout(() => api.save_text(els.text.value), 600);
  checkTimer = setTimeout(refreshMeta, 250);
});

/* ── drawing ───────────────────────────────────────────────────────── */
async function startDrawing() {
  if (state.drawing) return;

  let origin = null;
  if (state.platform && state.platform.linux) {
    origin = await pickLinuxOrigin();
    if (!origin) return; // cancelled or failed; status already reset
  }

  const res = await api.start_drawing(els.text.value, origin ? [origin.x, origin.y] : null);
  if (!res.ok) {
    if (res.error === 'unknown') {
      setStatus('Untrained characters', 'warn');
      toastWarn(res.message);
    } else {
      setStatus(res.message || 'Could not start', 'error');
      toastError(res.message || res.error || 'Could not start drawing.');
    }
    return;
  }
  setDrawing(true);
  els.draw.disabled = true;
  els.stop.disabled = false;
  els.progress.classList.remove('is-hidden');
  setProgress(0);
}

els.draw.addEventListener('click', startDrawing);

els.stop.addEventListener('click', () => {
  api.stop_drawing();
  setStatus('Stopping…', 'warn');
});

els.gotoLearn.addEventListener('click', () => onNavigateToLearn());

els.reset.addEventListener('click', async () => {
  const ok = await confirm({
    title: 'Reset all characters?',
    message: 'This deletes every trained character. You will have to draw the '
           + 'whole alphabet again. This cannot be undone.',
    confirmLabel: 'Delete everything',
    danger: true
  });
  if (!ok) return;
  const res = await api.reset_letters();
  setLearned(res.learned);
  toast('All characters deleted.', 'warn');
  onNavigateToLearn();
});

/* ── quick AI menu ─────────────────────────────────────────────────── */
function closeQuickMenu() { els.quickMenu.hidden = true; }

els.quickBtn.addEventListener('click', e => {
  e.stopPropagation();
  els.quickMenu.hidden = !els.quickMenu.hidden;
});

document.addEventListener('click', e => {
  if (!els.quickMenu.hidden && !els.quickMenu.parentElement.contains(e.target)) closeQuickMenu();
});

els.quickMenu.addEventListener('click', e => {
  const btn = e.target.closest('button[data-action]');
  if (!btn) return;
  closeQuickMenu();
  runQuickAction(btn.dataset.action);
});

function runQuickAction(action) {
  const text = els.text.value.trim();
  if (!text) { toastWarn('Type some text first.'); return; }
  const extra = action === 'replace_unknown'
    ? { known_chars: [...state.learned].join('') }
    : null;
  setStatus('AI is working…', 'busy');
  api.ai_action_run(action, els.text.value, extra);
}

window.T2P.on('ai:action', res => {
  if (res.ok) {
    els.text.value = res.result;
    api.save_text(res.result);
    refreshMeta();
    setStatus('Ready', 'ready');
    toastOk('AI result applied.');
  } else {
    setStatus('AI failed', 'error');
    toastError(res.message || res.error || 'The AI request failed.');
  }
});

/* ── engine events ─────────────────────────────────────────────────── */
window.T2P.on('countdown', ({ seconds }) => {
  setStatus(`${countdownMessage()} — starting in ${seconds}s`, 'busy');
});

window.T2P.on('started', () => setStatus('Drawing…', 'busy'));

window.T2P.on('progress', ({ value, char }) => {
  setProgress(value, char);
  if (!state.drawing) return;
  setStatus('Drawing…', 'busy');
});

window.T2P.on('scroll', () => setStatus('Scrolling…', 'busy'));

window.T2P.on('error', ({ message }) => {
  setStatus(message || 'Error', 'error');
  toastError(message || 'The drawing engine reported an error.');
});

window.T2P.on('done', ({ stopped }) => {
  setDrawing(false);
  els.draw.disabled = false;
  els.stop.disabled = true;
  if (stopped) {
    setStatus('Stopped', 'warn');
  } else {
    setProgress(1);
    setStatus('Finished', 'ready');
    toastOk('Text written.');
  }
  setTimeout(() => {
    if (!state.drawing) els.progress.classList.add('is-hidden');
  }, 2500);
});

/* ── init ──────────────────────────────────────────────────────────── */
export function initEditor({ text, onGotoLearn }) {
  els.text.value = text || '';
  onNavigateToLearn = onGotoLearn || (() => {});

  // On Linux, the start point is picked from a screenshot, not the live cursor
  const hint = document.getElementById('editor-hint');
  if (hint && state.platform && state.platform.linux) {
    hint.textContent = "You'll pick the start point on a screenshot before writing begins.";
  }

  setStatus('Ready', 'ready');
  refreshMeta();
}

export function focusEditor() {
  els.text.focus();
}
