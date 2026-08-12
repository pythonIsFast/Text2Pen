/* Linux/Wayland only: shows a portal screenshot and lets the user click where
   writing should begin. Replaces "hover the real cursor there" because the
   real cursor position can't be read reliably on Wayland (see engine.py). */

import { open } from './modal.js';

export function pickOriginFromFrame(frame) {
  return new Promise(resolve => {
    const bodyHTML = `
      <p class="notice">Click the spot where writing should begin.</p>
      <div class="origin-picker__frame" id="origin-picker-frame">
        <img id="origin-picker-img" src="data:image/png;base64,${frame.image_b64}" alt="Your screen">
        <div class="origin-picker__crosshair" id="origin-picker-crosshair" hidden></div>
      </div>`;

    let chosen = null;

    const handle = open({
      title: 'Pick the start point',
      bodyHTML,
      dismissible: true,
      wide: true,
      actions: [
        { label: 'Cancel', className: 'btn--ghost', onClick: () => { chosen = null; } },
        {
          label: 'Start here', className: 'btn--accent',
          onClick: () => { if (!chosen) return false; }
        }
      ]
    });

    const img = handle.el.querySelector('#origin-picker-img');
    const cross = handle.el.querySelector('#origin-picker-crosshair');

    img.addEventListener('click', e => {
      const rect = img.getBoundingClientRect();
      const fracX = (e.clientX - rect.left) / rect.width;
      const fracY = (e.clientY - rect.top) / rect.height;
      chosen = {
        x: Math.round(frame.offset_x + fracX * frame.width),
        y: Math.round(frame.offset_y + fracY * frame.height)
      };
      cross.hidden = false;
      cross.style.left = `${e.clientX - rect.left}px`;
      cross.style.top = `${e.clientY - rect.top}px`;
    });

    const observer = new MutationObserver(() => {
      if (!handle.el.isConnected) {
        observer.disconnect();
        resolve(chosen);
      }
    });
    observer.observe(document.getElementById('modal-root'), { childList: true });
  });
}
