/* Bridge to the Python backend.
   Every call is awaited against `pywebviewready`, so views can call the API
   during module init without racing the bridge injection. */

const ready = new Promise(resolve => {
  if (window.pywebview && window.pywebview.api) resolve();
  else window.addEventListener('pywebviewready', () => resolve(), { once: true });
});

async function call(name, args) {
  await ready;
  const fn = window.pywebview.api[name];
  if (typeof fn !== 'function') throw new Error(`Unknown backend method: ${name}`);
  return fn(...args);
}

/** api.some_method(a, b) → Python Api.some_method(a, b) */
export const api = new Proxy({}, {
  get: (_target, name) => (...args) => call(name, args)
});

export const apiReady = ready;
