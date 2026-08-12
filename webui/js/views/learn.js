/* Training view: trace a character, store its strokes.

   The logical canvas stays 700 × 450 with the guide glyph centred at
   (300, 200) — exactly the geometry the old Tk canvas used — so strokes
   recorded here stay compatible with an existing letter_db.json. */

import { api } from '../api.js';
import { state, setCurrentChar, setLearned, onChange } from '../state.js';
import { toastOk, toastError } from '../ui/toast.js';

const LOGICAL_W = 700;
const LOGICAL_H = 450;
const GUIDE_X = 300;
const GUIDE_Y = 200;
// Tk's ("Arial", 260, "bold") is 260 *points*; at 96 dpi that renders as
// roughly 347 CSS pixels. Matching it keeps new glyphs the same size as
// characters trained with the old UI.
const GUIDE_FONT_PX = 347;
const MIN_POINT_DISTANCE = 2;

const canvas = document.getElementById('canvas');
const frame = canvas.parentElement;
const ctx = canvas.getContext('2d');

const els = {
  char: document.getElementById('learn-char'),
  charmap: document.getElementById('charmap'),
  undo: document.getElementById('btn-undo'),
  clear: document.getElementById('btn-clear'),
  save: document.getElementById('btn-save-letter'),
  skip: document.getElementById('btn-skip'),
  prev: document.getElementById('btn-prev'),
  next: document.getElementById('btn-next'),
  fill: document.getElementById('learn-progress-fill'),
  text: document.getElementById('learn-progress-text'),
  pct: document.getElementById('learn-progress-pct')
};

let strokes = [];
let active = null;
// Cached so a redraw per pointer move does not hit getComputedStyle each time.
let colors = { bg: '#fff', guide: '#e0e0ea', ink: '#16161f' };

function refreshColors() {
  const s = getComputedStyle(document.documentElement);
  colors = {
    bg: s.getPropertyValue('--canvas-bg').trim() || '#fff',
    guide: s.getPropertyValue('--canvas-guide').trim() || '#e0e0ea',
    ink: s.getPropertyValue('--canvas-ink').trim() || '#16161f'
  };
}

/* ── canvas sizing ─────────────────────────────────────────────────── */
function resize() {
  const dpr = window.devicePixelRatio || 1;
  canvas.width = LOGICAL_W * dpr;
  canvas.height = LOGICAL_H * dpr;
  canvas.style.aspectRatio = `${LOGICAL_W} / ${LOGICAL_H}`;
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  redraw();
}

function toLogical(event) {
  const rect = canvas.getBoundingClientRect();
  return [
    Math.round((event.clientX - rect.left) / rect.width * LOGICAL_W),
    Math.round((event.clientY - rect.top) / rect.height * LOGICAL_H)
  ];
}

/* ── rendering ─────────────────────────────────────────────────────── */
function redraw() {
  ctx.clearRect(0, 0, LOGICAL_W, LOGICAL_H);

  ctx.fillStyle = colors.bg;
  ctx.fillRect(0, 0, LOGICAL_W, LOGICAL_H);

  // guide glyph
  ctx.save();
  ctx.fillStyle = colors.guide;
  ctx.font = `bold ${GUIDE_FONT_PX}px Arial, sans-serif`;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText(state.currentChar, GUIDE_X, GUIDE_Y);
  ctx.restore();

  // ink
  ctx.strokeStyle = colors.ink;
  ctx.lineWidth = 3;
  ctx.lineCap = 'round';
  ctx.lineJoin = 'round';
  for (const stroke of strokes) {
    if (stroke.length < 2) continue;
    ctx.beginPath();
    ctx.moveTo(stroke[0][0], stroke[0][1]);
    for (let i = 1; i < stroke.length; i++) ctx.lineTo(stroke[i][0], stroke[i][1]);
    ctx.stroke();
  }

  const hasInk = strokes.some(s => s.length >= 2);
  frame.classList.toggle('has-ink', hasInk || !!active);
  els.undo.disabled = strokes.length === 0;
  els.save.disabled = !hasInk;
}

/* ── pointer capture ───────────────────────────────────────────────── */
canvas.addEventListener('pointerdown', e => {
  e.preventDefault();
  // Capture keeps the stroke alive when the pointer leaves the canvas; it
  // throws for pointer ids the engine does not know, which must not abort.
  try { canvas.setPointerCapture(e.pointerId); } catch (_) { /* non-fatal */ }
  active = [toLogical(e)];
  strokes.push(active);
  redraw();
});

canvas.addEventListener('pointermove', e => {
  if (!active) return;
  const [x, y] = toLogical(e);
  const [lx, ly] = active[active.length - 1];
  // Thin out the stream: the engine samples every third point, so very dense
  // input would only slow drawing down without adding fidelity.
  if (Math.hypot(x - lx, y - ly) < MIN_POINT_DISTANCE) return;
  active.push([x, y]);
  redraw();
});

function endStroke() {
  if (!active) return;
  if (active.length < 2) strokes.pop();  // a tap is not a stroke
  active = null;
  redraw();
}

canvas.addEventListener('pointerup', endStroke);
canvas.addEventListener('pointercancel', endStroke);

/* ── actions ───────────────────────────────────────────────────────── */
function reset() {
  strokes = [];
  active = null;
  redraw();
}

els.undo.addEventListener('click', () => { strokes.pop(); redraw(); });
els.clear.addEventListener('click', reset);

els.save.addEventListener('click', async () => {
  const payload = strokes.filter(s => s.length >= 2);
  if (!payload.length) return;
  const res = await api.save_letter(state.currentChar, payload);
  if (!res.ok) { toastError(res.error || 'Could not save the character.'); return; }
  setLearned(res.learned);
  toastOk(`“${state.currentChar}” saved`);
  reset();
  step(1);
});

els.skip.addEventListener('click', () => { reset(); step(1); });
els.prev.addEventListener('click', () => { reset(); step(-1); });
els.next.addEventListener('click', () => { reset(); step(1); });

function step(delta) {
  const idx = state.alphabet.indexOf(state.currentChar);
  const next = idx + delta;
  if (next < 0 || next >= state.alphabet.length) return;
  setCurrentChar(state.alphabet[next]);
}

export function selectChar(ch) {
  reset();
  setCurrentChar(ch);
}

/* ── charmap + progress ────────────────────────────────────────────── */
function renderCharmap() {
  els.charmap.innerHTML = '';
  const frag = document.createDocumentFragment();
  for (const ch of state.alphabet) {
    const cell = document.createElement('button');
    cell.className = 'charmap__cell';
    cell.textContent = ch === ' ' ? '␣' : ch;
    cell.title = state.learned.has(ch) ? `${ch} — trained` : `${ch} — not trained`;
    if (state.learned.has(ch)) cell.classList.add('is-learned');
    if (ch === state.currentChar) cell.classList.add('is-current');
    cell.addEventListener('click', () => selectChar(ch));
    frag.appendChild(cell);
  }
  els.charmap.appendChild(frag);
}

function renderProgress() {
  const total = state.alphabet.length || 1;
  const done = state.learned.size;
  const pct = Math.round(done / total * 100);
  els.fill.style.width = `${pct}%`;
  els.text.textContent = `${done} / ${state.alphabet.length}`;
  els.pct.textContent = `${pct}%`;
}

function render() {
  els.char.textContent = state.currentChar;
  const idx = state.alphabet.indexOf(state.currentChar);
  els.prev.disabled = idx <= 0;
  els.next.disabled = idx < 0 || idx >= state.alphabet.length - 1;
  renderCharmap();
  renderProgress();
  redraw();
}

export function initLearn() {
  refreshColors();
  resize();
  render();
  onChange(render);
  window.addEventListener('resize', resize);
  matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => {
    refreshColors();
    redraw();
  });
}

/** Called after the theme setting changes, so the canvas repaints in-place. */
export function refreshCanvasTheme() {
  refreshColors();
  redraw();
}
