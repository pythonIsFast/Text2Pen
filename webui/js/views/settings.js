/* Settings dialog and the first-run telemetry consent. */

import { api } from '../api.js';
import { state, patchSettings } from '../state.js';
import { open } from '../ui/modal.js';
import { toastOk } from '../ui/toast.js';
import { refreshCanvasTheme } from './learn.js';

export function applyTheme(theme) {
  document.documentElement.dataset.theme = theme || 'system';
  // The canvas paints its own background, so it needs an explicit repaint.
  refreshCanvasTheme();
}

export function openSettings() {
  const s = state.settings;

  const handle = open({
    title: 'Settings',
    bodyHTML: `
      <div class="field">
        <div class="field__row">
          <div>
            <div class="field__label">Character size</div>
            <div class="field__hint">Scale factor applied to every stroke.</div>
          </div>
          <input class="input" id="set-size" type="number"
                 min="0.05" max="1" step="0.05" value="${s.character_size ?? 0.1}">
        </div>
      </div>

      <div class="field">
        <div class="field__row">
          <div>
            <div class="field__label">Line spacing</div>
            <div class="field__hint">Pixels between two baselines.</div>
          </div>
          <input class="input" id="set-spacing" type="number"
                 min="10" max="200" step="5" value="${s.line_spacing ?? 60}">
        </div>
      </div>

      <div class="field">
        <div class="field__row">
          <div>
            <div class="field__label">Start delay</div>
            <div class="field__hint">Seconds to switch windows before drawing starts.</div>
          </div>
          <input class="input" id="set-delay" type="number"
                 min="1" max="30" step="1" value="${s.start_delay ?? 4}">
        </div>
      </div>

      <div class="field">
        <div class="field__row">
          <div>
            <div class="field__label">Appearance</div>
            <div class="field__hint">Follows the system theme by default.</div>
          </div>
          <select class="select" id="set-theme">
            <option value="system">System</option>
            <option value="dark">Dark</option>
            <option value="light">Light</option>
          </select>
        </div>
      </div>

      <div class="field">
        <div class="field__row">
          <div>
            <div class="field__label">Progress overlay</div>
            <div class="field__hint">Small always-on-top bar while drawing.</div>
          </div>
          <label class="switch">
            <input type="checkbox" id="set-overlay" ${s.show_overlay ? 'checked' : ''}>
            <span class="switch__track"></span>
          </label>
        </div>
      </div>

      <div class="divider"></div>

      <div class="field">
        <div class="field__row">
          <div>
            <div class="field__label">Anonymous telemetry</div>
            <div class="field__hint">Send crash and error reports only.</div>
          </div>
          <label class="switch">
            <input type="checkbox" id="set-telemetry" ${s.telemetry_opted_in ? 'checked' : ''}>
            <span class="switch__track"></span>
          </label>
        </div>
      </div>

      <div class="notice">
        <strong>AI limits today</strong>
        <div id="set-limits" style="margin-top:6px">Loading…</div>
      </div>

      <div class="notice">
        <strong>Data location</strong>
        <div style="margin-top:6px"><code>${state.dataDir}</code></div>
      </div>
    `,
    actions: [
      {
        label: 'Save',
        className: 'btn--accent',
        onClick: el => {
          const size = parseFloat(el.querySelector('#set-size').value);
          const spacing = parseInt(el.querySelector('#set-spacing').value, 10);
          const delay = parseInt(el.querySelector('#set-delay').value, 10);
          const theme = el.querySelector('#set-theme').value;
          const patch = {
            character_size: Number.isFinite(size) ? Math.min(Math.max(size, 0.05), 1) : 0.1,
            line_spacing: Number.isFinite(spacing) ? Math.min(Math.max(spacing, 10), 200) : 60,
            start_delay: Number.isFinite(delay) ? Math.min(Math.max(delay, 1), 30) : 4,
            theme,
            show_overlay: el.querySelector('#set-overlay').checked,
            telemetry_opted_in: el.querySelector('#set-telemetry').checked
          };
          patchSettings(patch);
          applyTheme(theme);
          api.save_settings(patch).then(() => toastOk('Settings saved.'));
        }
      },
      { label: 'Cancel', className: 'btn--ghost' }
    ]
  });

  handle.el.querySelector('#set-theme').value = s.theme || 'system';

  // Limits arrive asynchronously; the dialog may already be gone by then.
  const off = window.T2P.on('ai:limits', res => {
    const box = handle.el.querySelector('#set-limits');
    if (!box) return;
    box.textContent = res.ok
      ? `Chat ${res.limits.chat_used ?? 0}/${res.limits.chat_limit ?? 10} · `
        + `Text actions ${res.limits.text_used ?? 0}/${res.limits.text_limit ?? 1000}`
      : 'Not available';
  });
  api.ai_limits_refresh();
  return off;
}

export function openTelemetryConsent() {
  open({
    title: 'Welcome to Text2Pen',
    dismissible: false,
    bodyHTML: `
      <p>Help improve Text2Pen by sharing anonymous diagnostics.</p>
      <div class="notice">
        <strong>What would be sent</strong>
        <ul>
          <li>Crash reports and error messages</li>
          <li>Basic usage statistics</li>
          <li>Which features are used</li>
        </ul>
        Reports contain no personal data and your username is stripped out.
        Data may be processed outside the EU.
      </div>
      <div class="field">
        <div class="field__row">
          <div class="field__label">Send anonymous diagnostics</div>
          <label class="switch">
            <input type="checkbox" id="consent-toggle">
            <span class="switch__track"></span>
          </label>
        </div>
      </div>
    `,
    actions: [{
      label: 'Continue',
      className: 'btn--accent',
      onClick: el => {
        const opted = el.querySelector('#consent-toggle').checked;
        patchSettings({ telemetry_opted_in: opted });
        api.save_settings({ telemetry_opted_in: opted });
      }
    }]
  });
}
