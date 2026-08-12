/* Transient notifications, bottom right. */

const root = document.getElementById('toasts');

export function toast(message, kind = 'info', duration = 3600) {
  const el = document.createElement('div');
  el.className = `toast toast--${kind}`;
  el.textContent = message;
  root.appendChild(el);

  const remove = () => {
    el.classList.add('is-leaving');
    el.addEventListener('animationend', () => el.remove(), { once: true });
  };
  setTimeout(remove, duration);
  return remove;
}

export const toastOk = m => toast(m, 'ok');
export const toastWarn = m => toast(m, 'warn', 5000);
export const toastError = m => toast(m, 'error', 6000);
