/* Modal dialogs. `open()` returns the element plus a close handle so callers
   can read their own form fields before closing. */

const root = document.getElementById('modal-root');

export function open({ title, icon = '', bodyHTML = '', actions = [], dismissible = true, wide = false }) {
  const backdrop = document.createElement('div');
  backdrop.className = 'modal-backdrop';
  backdrop.innerHTML = `
    <div class="modal${wide ? ' modal--wide' : ''}" role="dialog" aria-modal="true">
      <div class="modal__head">
        <h2>${icon} ${title}</h2>
      </div>
      <div class="modal__body">${bodyHTML}</div>
      <div class="modal__foot"></div>
    </div>`;

  const foot = backdrop.querySelector('.modal__foot');
  const close = () => {
    document.removeEventListener('keydown', onKey);
    backdrop.remove();
  };

  actions.forEach(action => {
    const btn = document.createElement('button');
    btn.className = `btn ${action.className || 'btn--ghost'}`;
    btn.textContent = action.label;
    btn.addEventListener('click', () => {
      // An action may veto closing by returning false.
      if (action.onClick && action.onClick(backdrop) === false) return;
      close();
    });
    foot.appendChild(btn);
  });

  function onKey(e) {
    if (e.key === 'Escape' && dismissible) close();
  }

  if (dismissible) {
    backdrop.addEventListener('mousedown', e => {
      if (e.target === backdrop) close();
    });
  }
  document.addEventListener('keydown', onKey);

  root.appendChild(backdrop);
  const firstInput = backdrop.querySelector('input, select, textarea, button');
  if (firstInput) firstInput.focus();

  return { el: backdrop, close };
}

export function confirm({ title, message, confirmLabel = 'Confirm', danger = false }) {
  return new Promise(resolve => {
    let settled = false;
    const finish = value => { settled = true; resolve(value); };

    const handle = open({
      title,
      bodyHTML: `<p class="notice">${message}</p>`,
      actions: [
        { label: 'Cancel', className: 'btn--ghost', onClick: () => finish(false) },
        {
          label: confirmLabel,
          className: danger ? 'btn--danger' : 'btn--accent',
          onClick: () => finish(true)
        }
      ]
    });

    // Dismissing via Escape or backdrop counts as cancel.
    const observer = new MutationObserver(() => {
      if (!handle.el.isConnected) {
        observer.disconnect();
        if (!settled) resolve(false);
      }
    });
    observer.observe(document.getElementById('modal-root'), { childList: true });
  });
}
