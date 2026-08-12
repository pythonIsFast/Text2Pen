"""Text2Pen — entry point.

Thin shell around the backend: opens a pywebview window that renders the
frontend in webui/ and exposes the backend to it through `Api`.
"""

import json
import os
import sys
import time
from threading import Thread

import webview

from backend.engine import DrawingEngine
from backend.input_controller import create_input_controller
from backend.paths import (APP_DATA_DIR, IS_LINUX, IS_WAYLAND, IS_WINDOWS,
                           UPDATE_EXE, ensure_app_data_dir, is_root, webui_index)
from backend.services import STATUS_URL, WEBSITE_URL, AIClient, Telemetry
from backend.storage import ALPHABET, Storage
from backend.text_parser import count_drawable_chars, find_unknown_chars, parse_text_with_tables

OVERLAY_HTML = """
<!doctype html><meta charset="utf-8">
<style>
  html,body{margin:0;height:100%;background:#12121a;overflow:hidden;
            font:600 11px/1 system-ui,sans-serif;color:#e8e6f0}
  .wrap{display:flex;align-items:center;gap:8px;height:100%;padding:0 10px;box-sizing:border-box}
  .track{flex:1;height:8px;border-radius:99px;background:#2a2a38;overflow:hidden}
  .fill{height:100%;width:0%;border-radius:99px;
        background:linear-gradient(90deg,#7719AA,#b45cf0);transition:width .12s linear}
  .pct{min-width:34px;text-align:right;font-variant-numeric:tabular-nums}
</style>
<div class="wrap"><div class="track"><div class="fill" id="f"></div></div>
<div class="pct" id="p">0%</div></div>
<script>
  window.setProgress = v => {
    const pct = Math.round(Math.max(0, Math.min(1, v)) * 100);
    document.getElementById('f').style.width = pct + '%';
    document.getElementById('p').textContent = pct + '%';
  };
</script>
"""


# --------------------------------------------------------------------- updater
def _is_update_running():
    if not IS_WINDOWS:
        return False
    try:
        import psutil
    except ImportError:
        return False
    for proc in psutil.process_iter(["exe"]):
        try:
            exe = proc.info.get("exe")
            if exe and os.path.normcase(exe) == os.path.normcase(UPDATE_EXE):
                return True
        except Exception:
            continue
    return False


def _handle_update_handoff():
    """Cooperate with Update.exe: step aside while it runs, swap in new updater."""
    if _is_update_running():
        print("Update is running, exiting Text2Pen to allow the update…")
        time.sleep(2)
        sys.exit(0)

    if IS_WINDOWS:
        staged = os.path.join(APP_DATA_DIR, "Update.exe-newest")
        if os.path.exists(staged):
            try:
                os.replace(staged, UPDATE_EXE)
            except OSError as e:
                print(f"Could not swap in the new updater: {e}")

    if len(sys.argv) > 1 and sys.argv[1] == "updaterStart":
        sys.exit(0)


# -------------------------------------------------------------------- JS bridge
class Api:
    """Every public method here is callable from JS as window.pywebview.api.<name>."""

    def __init__(self):
        ensure_app_data_dir()
        self.storage = Storage()
        self.letter_db = self.storage.load_letters()
        self.settings, self.first_run = self.storage.load_settings()
        self.telemetry = Telemetry(self.settings.get("telemetry_opted_in"))
        self.ai = AIClient()

        self.window = None
        self._overlay = None
        self._input_error = None
        self._last_origin = None

        try:
            controller = create_input_controller()
        except Exception as e:
            # A missing /dev/uinput permission must not stop the UI from opening —
            # the frontend shows the reason instead of the app dying on launch.
            controller = None
            self._input_error = str(e)
            print(f"Input controller unavailable: {e}")

        self.engine = DrawingEngine(controller, self._emit) if controller else None

    # ------------------------------------------------------------ event bridge
    def _emit(self, event, payload=None):
        if self.window is None:
            return
        js = (f"window.T2P && window.T2P.emit({json.dumps(event)}, "
              f"{json.dumps(payload or {})})")
        try:
            self.window.evaluate_js(js)
        except Exception:
            pass
        if event == "progress":
            self._update_overlay((payload or {}).get("value", 0))
        elif event == "started":
            self._open_overlay()
        elif event == "done":
            self._close_overlay()

    # ------------------------------------------------------------------ overlay
    def _open_overlay(self):
        if not self.settings.get("show_overlay", True) or self._overlay is not None:
            return
        try:
            width, height = 320, 48
            margin = 20
            x, y = margin, margin
            try:
                screen = webview.screens[0]
                # Default corner: top-right. If the user picked a start point
                # (Linux origin picker) near that corner, use the corner
                # farthest from it instead, so the overlay never sits on top
                # of the handwriting.
                corner_x = max(screen.width - width - margin, 0)
                corner_y = margin
                origin = self._last_origin
                if origin is not None:
                    ox, oy = origin
                    if ox < screen.width / 2:
                        corner_x = max(screen.width - width - margin, 0)
                    else:
                        corner_x = margin
                    if oy < screen.height / 2:
                        corner_y = max(screen.height - height - margin, 0)
                    else:
                        corner_y = margin
                x, y = corner_x, corner_y
            except Exception:
                pass
            self._overlay = webview.create_window(
                "Text2Pen Progress", html=OVERLAY_HTML,
                width=width, height=height, x=x, y=y,
                frameless=True, on_top=True, resizable=False,
                focus=False, easy_drag=True,
            )
        except Exception as e:
            self._overlay = None
            print(f"Progress overlay unavailable: {e}")

    def _update_overlay(self, value):
        if self._overlay is None:
            return
        try:
            self._overlay.evaluate_js(f"window.setProgress({float(value)})")
        except Exception:
            pass

    def _close_overlay(self):
        if self._overlay is None:
            return
        try:
            self._overlay.destroy()
        except Exception:
            pass
        self._overlay = None

    # ---------------------------------------------------------------- bootstrap
    def bootstrap(self):
        """Everything the frontend needs on load."""
        return {
            "alphabet": list(ALPHABET),
            "learned": sorted(self.letter_db.keys()),
            "learned_count": len(self.letter_db),
            "settings": self.settings,
            "first_run": self.first_run or self.settings.get("telemetry_opted_in") is None,
            "text": self.storage.load_text(),
            "platform": {
                "windows": IS_WINDOWS,
                "linux": IS_LINUX,
                "wayland": IS_WAYLAND,
                "is_root": is_root(),
            },
            "data_dir": APP_DATA_DIR,
            "input_error": self._input_error,
            "urls": {"website": WEBSITE_URL, "status": STATUS_URL},
        }

    # ------------------------------------------------------------------ letters
    def save_letter(self, char, strokes):
        if not char or not isinstance(strokes, list) or not strokes:
            return {"ok": False, "error": "No strokes to save."}
        # Normalise to the [[x, y], …] integer form the engine expects.
        clean = []
        for stroke in strokes:
            points = [[int(p[0]), int(p[1])] for p in stroke if len(p) >= 2]
            if len(points) >= 2:
                clean.append(points)
        if not clean:
            return {"ok": False, "error": "Strokes were too short to save."}
        self.letter_db[char] = clean
        self.storage.save_letters(self.letter_db)
        return {"ok": True, "learned": sorted(self.letter_db.keys()),
                "learned_count": len(self.letter_db)}

    def delete_letter(self, char):
        self.letter_db.pop(char, None)
        self.storage.save_letters(self.letter_db)
        return {"ok": True, "learned": sorted(self.letter_db.keys()),
                "learned_count": len(self.letter_db)}

    def reset_letters(self):
        self.letter_db = {}
        self.storage.save_letters(self.letter_db)
        return {"ok": True, "learned": [], "learned_count": 0}

    def get_letter(self, char):
        return {"ok": True, "strokes": self.letter_db.get(char, [])}

    # ----------------------------------------------------------------- settings
    def save_settings(self, patch):
        if isinstance(patch, dict):
            self.settings.update(patch)
        self.storage.save_settings(self.settings)
        self.telemetry.enabled = bool(self.settings.get("telemetry_opted_in"))
        return {"ok": True, "settings": self.settings}

    # --------------------------------------------------------------------- text
    def save_text(self, text):
        return {"ok": self.storage.save_text(text or "")}

    def check_text(self, text):
        blocks = parse_text_with_tables(text or "")
        return {
            "ok": True,
            "unknown": find_unknown_chars(blocks, self.letter_db),
            "total_chars": count_drawable_chars(blocks),
        }

    # ------------------------------------------------------------------ drawing
    def pick_origin(self):
        """Linux/Wayland only: screenshot a monitor via the ScreenCast portal
        so the user can click where writing should begin. The real cursor
        position can't be queried reliably on Wayland (see input_controller.py),
        so this replaces "hover the mouse" with "click the spot in a picture"."""
        if not IS_LINUX:
            return {"ok": False, "error": "unsupported",
                    "message": "Origin picking is only needed on Linux."}

        def run():
            from backend.portal_capture import capture_screen
            restore_token = self.settings.get("screencast_restore_token")
            res = capture_screen(restore_token)
            if res.get("ok") and res.get("restore_token"):
                self.settings["screencast_restore_token"] = res["restore_token"]
                self.storage.save_settings(self.settings)
            self._emit("origin:frame", res)

        Thread(target=run, daemon=True).start()
        return {"ok": True}

    def start_drawing(self, text, origin=None):
        if self.engine is None:
            return {"ok": False, "error": "input",
                    "message": self._input_error or "Input device unavailable."}
        self.storage.save_text(text or "")
        blocks = parse_text_with_tables(text or "")
        unknown = find_unknown_chars(blocks, self.letter_db)
        if unknown:
            return {"ok": False, "error": "unknown", "unknown": unknown,
                    "message": f"Not trained yet: {' '.join(unknown)}"}
        origin_xy = None
        if isinstance(origin, (list, tuple)) and len(origin) == 2:
            origin_xy = (int(origin[0]), int(origin[1]))
        self._last_origin = origin_xy
        return self.engine.start(text or "", self.letter_db, self.settings, origin_xy)

    def stop_drawing(self):
        if self.engine:
            self.engine.stop()
        return {"ok": True}

    def is_drawing(self):
        return {"running": bool(self.engine and self.engine.is_running())}

    # ----------------------------------------------------------------------- AI
    # Network calls run detached and answer through events, so a slow backend
    # can never block the UI thread.
    def ai_limits_refresh(self):
        def run():
            res = self.ai.limits()
            self._emit("ai:limits", res)
        Thread(target=run, daemon=True).start()
        return {"ok": True}

    def ai_chat_send(self, history, message, image_b64=None):
        def run():
            res = self.ai.chat(history or [], message or "", image_b64)
            self._emit("ai:reply", res)
        Thread(target=run, daemon=True).start()
        return {"ok": True}

    def ai_action_run(self, action, text, extra=None):
        def run():
            res = self.ai.action(action, text or "", extra)
            res["action"] = action
            self._emit("ai:action", res)
        Thread(target=run, daemon=True).start()
        return {"ok": True}

    def known_chars(self):
        return {"ok": True, "chars": "".join(sorted(self.letter_db.keys()))}

    # -------------------------------------------------------------------- misc
    def report_error(self, message):
        """Frontend-side crash reporting, honouring the telemetry opt-in."""
        Thread(target=self.telemetry.send_error, args=(message,), daemon=True).start()
        return {"ok": True}

    def open_external(self, url):
        if not isinstance(url, str) or not url.startswith(("http://", "https://")):
            return {"ok": False, "error": "Refused to open a non-http(s) URL."}
        import webbrowser
        webbrowser.open(url)
        return {"ok": True}

    def reveal_data_dir(self):
        return {"ok": True, "path": APP_DATA_DIR}


def main():
    _handle_update_handoff()

    api = Api()
    window = webview.create_window(
        "Text2Pen",
        webui_index(),
        js_api=api,
        width=1280,
        height=860,
        min_size=(940, 620),
        background_color="#12121a",
    )
    api.window = window

    def on_closed():
        if api.engine:
            api.engine.stop()

    window.events.closed += on_closed

    # http_server serves webui/ over localhost so ES modules load with a real
    # origin; from file:// the browser engine would block them via CORS.
    webview.start(http_server=True, debug=bool(os.environ.get("TEXT2PEN_DEBUG")))


if __name__ == "__main__":
    main()
