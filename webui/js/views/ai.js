/* AI drawer: chat with optional image attachment. */

import { api } from '../api.js';
import { toastError } from '../ui/toast.js';

const els = {
  drawer: document.getElementById('drawer'),
  toggle: document.getElementById('btn-ai'),
  close: document.getElementById('drawer-close'),
  limits: document.getElementById('ai-limits'),
  chat: document.getElementById('chat'),
  form: document.getElementById('composer'),
  input: document.getElementById('chat-input'),
  file: document.getElementById('chat-file'),
  attachment: document.getElementById('chat-attachment'),
  thumb: document.getElementById('chat-attachment-thumb'),
  name: document.getElementById('chat-attachment-name'),
  clearAttachment: document.getElementById('chat-attachment-clear')
};

const WELCOME = `Hi! I'm your Text2Pen assistant.

I can help you with:
• correcting or rewriting text
• analysing an image you attach
• answering questions

When I suggest text, you can push it straight into the text field.`;

let history = [];
let attached = null;      // { b64, name, dataUrl }
let typingBubble = null;
let onInsertText = () => {};

/* ── bubbles ───────────────────────────────────────────────────────── */
function addBubble(role, text, { imageDataUrl = null, actions = false, error = false } = {}) {
  const bubble = document.createElement('div');
  bubble.className = `bubble bubble--${role === 'user' ? 'user' : 'ai'}`;
  if (error) bubble.classList.add('bubble--error');

  if (imageDataUrl) {
    const img = document.createElement('img');
    img.src = imageDataUrl;
    img.alt = '';
    bubble.appendChild(img);
  }

  const body = document.createElement('div');
  body.textContent = text;
  bubble.appendChild(body);

  if (actions) {
    const row = document.createElement('div');
    row.className = 'bubble__actions';

    const insert = document.createElement('button');
    insert.className = 'btn btn--subtle btn--sm';
    insert.textContent = 'Insert into text field';
    insert.addEventListener('click', () => {
      onInsertText(stripCodeFence(text));
      row.remove();
    });

    const dismiss = document.createElement('button');
    dismiss.className = 'btn btn--ghost btn--sm';
    dismiss.textContent = 'Dismiss';
    dismiss.addEventListener('click', () => row.remove());

    row.append(insert, dismiss);
    bubble.appendChild(row);
  }

  els.chat.appendChild(bubble);
  els.chat.scrollTop = els.chat.scrollHeight;
  return bubble;
}

function addTyping(hasImage) {
  const bubble = document.createElement('div');
  bubble.className = 'bubble bubble--ai bubble--typing';
  const label = hasImage ? 'Analysing image' : 'Thinking';
  bubble.textContent = `${label}.`;

  let dots = 1;
  const timer = setInterval(() => {
    dots = dots % 3 + 1;
    bubble.textContent = label + '.'.repeat(dots);
  }, 450);
  bubble._stop = () => clearInterval(timer);

  if (hasImage) {
    const note = document.createElement('div');
    note.style.cssText = 'font-size:11px;opacity:.7;margin-top:4px';
    note.textContent = 'The first request can take ~30 s while the server wakes up.';
    bubble.appendChild(note);
  }

  els.chat.appendChild(bubble);
  els.chat.scrollTop = els.chat.scrollHeight;
  return bubble;
}

function removeTyping() {
  if (!typingBubble) return;
  if (typingBubble._stop) typingBubble._stop();
  typingBubble.remove();
  typingBubble = null;
}

/** A reply long enough — or fenced — probably contains usable prose. */
function looksLikeSuggestion(reply) {
  const keywords = ['here is', "here's", 'suggestion:', 'text:', 'corrected:',
                    'revised:', 'version:', 'rewritten:', '```'];
  const lower = reply.toLowerCase();
  return keywords.some(k => lower.includes(k)) || reply.length > 80;
}

function stripCodeFence(text) {
  if (!text.includes('```')) return text.trim();
  const parts = text.split('```');
  if (parts.length >= 3) {
    let body = parts[1];
    const newline = body.indexOf('\n');
    if (newline !== -1) body = body.slice(newline + 1);   // drop the language tag
    return body.trim();
  }
  return text.replaceAll('```', '').trim();
}

/* ── attachment ────────────────────────────────────────────────────── */
els.file.addEventListener('change', () => {
  const file = els.file.files && els.file.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    const dataUrl = String(reader.result);
    attached = { b64: dataUrl.split(',')[1] || '', name: file.name, dataUrl };
    els.thumb.src = dataUrl;
    els.name.textContent = file.name;
    els.attachment.classList.remove('is-hidden');
  };
  reader.onerror = () => toastError('Could not read that image.');
  reader.readAsDataURL(file);
});

function clearAttachment() {
  attached = null;
  els.file.value = '';
  els.attachment.classList.add('is-hidden');
}

els.clearAttachment.addEventListener('click', clearAttachment);

/* ── sending ───────────────────────────────────────────────────────── */
function send() {
  const text = els.input.value.trim();
  if (!text && !attached) return;

  addBubble('user', text || '(image)', { imageDataUrl: attached ? attached.dataUrl : null });

  const outgoing = { role: 'user', content: text };
  if (attached) outgoing.image_b64 = attached.b64;

  const priorHistory = history.slice();
  history.push(outgoing);

  const imageB64 = attached ? attached.b64 : null;
  els.input.value = '';
  els.input.style.height = 'auto';
  clearAttachment();

  typingBubble = addTyping(!!imageB64);
  api.ai_chat_send(priorHistory, text, imageB64);
}

els.form.addEventListener('submit', e => { e.preventDefault(); send(); });

els.input.addEventListener('keydown', e => {
  if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); send(); }
});

els.input.addEventListener('input', () => {
  els.input.style.height = 'auto';
  els.input.style.height = `${Math.min(els.input.scrollHeight, 140)}px`;
});

window.T2P.on('ai:reply', res => {
  removeTyping();
  if (!res.ok) {
    addBubble('assistant', res.message || res.error || 'The request failed.', { error: true });
    return;
  }
  history.push({ role: 'assistant', content: res.result });
  addBubble('assistant', res.result, { actions: looksLikeSuggestion(res.result) });
  api.ai_limits_refresh();
});

window.T2P.on('ai:limits', res => {
  if (!res.ok) { els.limits.textContent = 'Limits unavailable'; return; }
  const l = res.limits || {};
  els.limits.textContent =
    `Chat ${l.chat_used ?? 0}/${l.chat_limit ?? 10} · `
    + `Text actions ${l.text_used ?? 0}/${l.text_limit ?? 1000} today`;
});

/* ── open / close ──────────────────────────────────────────────────── */
function setOpen(open) {
  els.drawer.classList.toggle('is-open', open);
  els.drawer.setAttribute('aria-hidden', String(!open));
  els.toggle.classList.toggle('btn--accent', !open);
  els.toggle.classList.toggle('btn--subtle', open);
  if (open) {
    if (!els.chat.childElementCount) addBubble('assistant', WELCOME);
    api.ai_limits_refresh();
    els.input.focus();
  }
}

els.toggle.addEventListener('click', () => setOpen(!els.drawer.classList.contains('is-open')));
els.close.addEventListener('click', () => setOpen(false));

export function initAI({ onInsert }) {
  onInsertText = onInsert || (() => {});
}
