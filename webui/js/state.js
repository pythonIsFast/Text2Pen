/* Shared app state. Views mutate it through the setters and re-render on
   'state:changed', which keeps the nav badge, charmap and editor in sync. */

export const state = {
  alphabet: [],
  learned: new Set(),
  settings: {},
  platform: {},
  dataDir: '',
  urls: {},
  inputError: null,
  currentChar: 'a',
  drawing: false
};

const listeners = new Set();

export function onChange(fn) {
  listeners.add(fn);
  return () => listeners.delete(fn);
}

function notify() {
  listeners.forEach(fn => {
    try { fn(state); } catch (e) { console.error(e); }
  });
}

export function hydrate(payload) {
  state.alphabet = payload.alphabet || [];
  state.learned = new Set(payload.learned || []);
  state.settings = payload.settings || {};
  state.platform = payload.platform || {};
  state.dataDir = payload.data_dir || '';
  state.urls = payload.urls || {};
  state.inputError = payload.input_error || null;
  state.currentChar = firstUntrained();
  notify();
}

export function setLearned(list) {
  state.learned = new Set(list || []);
  notify();
}

export function setCurrentChar(ch) {
  state.currentChar = ch;
  notify();
}

export function setDrawing(flag) {
  state.drawing = flag;
  notify();
}

export function patchSettings(patch) {
  Object.assign(state.settings, patch);
  notify();
}

export function firstUntrained() {
  const next = state.alphabet.find(ch => !state.learned.has(ch));
  return next || state.alphabet[0] || 'a';
}

export function isFullyTrained() {
  return state.alphabet.length > 0 && state.learned.size >= state.alphabet.length;
}
