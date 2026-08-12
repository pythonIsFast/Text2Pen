"""Low-level pointer control, one implementation per platform."""

import re
import subprocess
import sys
import time

from .paths import IS_LINUX, IS_WINDOWS

LINUX_DEBUG = IS_LINUX

if IS_WINDOWS:
    import win32api
    import win32con
    import win32gui

if IS_LINUX:
    try:
        import uinput
    except ImportError:
        uinput = None
    try:
        from pynput.mouse import Controller as PynputMouseController
    except ImportError:
        PynputMouseController = None


def debug_linux(msg):
    if LINUX_DEBUG:
        try:
            print(f"[linux-debug] {msg}")
            sys.stdout.flush()
        except Exception:
            pass


class WindowsInputController:
    def get_cursor_pos(self):
        return win32api.GetCursorPos()

    def set_cursor_pos(self, x, y):
        win32api.SetCursorPos((int(x), int(y)))

    def mouse_down(self):
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)

    def mouse_up(self):
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)

    def mouse_move_rel(self, dx, dy):
        win32api.mouse_event(win32con.MOUSEEVENTF_MOVE, int(dx), int(dy))

    def mouse_wheel(self, delta):
        win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, int(delta))

    def find_onenote_window(self):
        def cb(hwnd, out):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                cls = win32gui.GetClassName(hwnd)
                if "OneNote" in title or "OneNote" in cls:
                    out.append(hwnd)
            return True
        res = []
        win32gui.EnumWindows(cb, res)
        return res[0] if res else None

    def focus_window(self, hwnd):
        win32gui.SetForegroundWindow(hwnd)

    def client_to_screen(self, hwnd, x, y):
        return win32gui.ClientToScreen(hwnd, (x, y))


class UInputController:
    def __init__(self):
        if uinput is None:
            raise RuntimeError(
                "uinput is required on Linux. Install it with 'pip install python-uinput'."
            )
        if PynputMouseController is None:
            raise RuntimeError(
                "pynput is required on Linux. Install it with 'pip install pynput'."
            )

        screen_w, screen_h = self._detect_resolution()
        debug_linux(f"UInputController initialized with resolution {screen_w}x{screen_h}")

        events = (
            uinput.ABS_X + (0, screen_w, 0, 0),
            uinput.ABS_Y + (0, screen_h, 0, 0),
            uinput.BTN_LEFT,
            uinput.REL_WHEEL,
        )

        self.device = uinput.Device(events)
        self.mouse = PynputMouseController()
        self.screen_w = screen_w
        self.screen_h = screen_h
        time.sleep(0.5)  # let the virtual device settle

    @staticmethod
    def _detect_resolution():
        # Virtual screen across all monitors
        try:
            output = subprocess.check_output(["xrandr", "--current"],
                                            stderr=subprocess.DEVNULL).decode()
            match = re.search(r"current\s+(\d+)\s+x\s+(\d+)", output)
            if match:
                return int(match.group(1)), int(match.group(2))
        except Exception:
            pass
        # Primary screen only
        try:
            output = subprocess.check_output(["xdotool", "getdisplaygeometry"],
                                            stderr=subprocess.DEVNULL).decode().strip()
            parts = output.split()
            return int(parts[0]), int(parts[1])
        except Exception:
            return 1920, 1080

    def get_cursor_pos(self):
        try:
            x, y = self.mouse.position
            debug_linux(f"get_cursor_pos: ({x}, {y}) [screen: {self.screen_w}x{self.screen_h}]")
            return x, y
        except Exception as e:
            debug_linux(f"get_cursor_pos failed: {e}")
            return None

    def set_cursor_pos(self, x, y):
        try:
            x = max(0, min(int(x), self.screen_w - 1))
            y = max(0, min(int(y), self.screen_h - 1))
            self.device.emit(uinput.ABS_X, x, syn=False)
            self.device.emit(uinput.ABS_Y, y, syn=True)
            time.sleep(0.005)
        except Exception as e:
            debug_linux(f"set_cursor_pos failed: {e}")

    def mouse_down(self):
        try:
            self.device.emit(uinput.BTN_LEFT, 1, syn=True)
            time.sleep(0.012)
        except Exception as e:
            debug_linux(f"mouse_down failed: {e}")

    def mouse_up(self):
        try:
            self.device.emit(uinput.BTN_LEFT, 0, syn=True)
            # A clear pause here matters more than most other delays: without
            # it, the next stroke's press can reach the compositor before this
            # release is processed, and apps read that as one held-down drag
            # instead of two separate strokes.
            time.sleep(0.02)
        except Exception as e:
            debug_linux(f"mouse_up failed: {e}")

    def move_smooth(self, x1, y1, x2, y2, step_px=5):
        """Interpolated absolute move, without touching the button — the
        caller holds it down (or not) for the whole stroke around this."""
        try:
            x1, y1, x2, y2 = int(x1), int(y1), int(x2), int(y2)
            dx, dy = x2 - x1, y2 - y1
            distance = (dx * dx + dy * dy) ** 0.5
            steps = max(int(distance / step_px), 1)
            for i in range(1, steps + 1):
                nx = max(0, min(x1 + int(dx * i / steps), self.screen_w - 1))
                ny = max(0, min(y1 + int(dy * i / steps), self.screen_h - 1))
                self.device.emit(uinput.ABS_X, nx, syn=False)
                self.device.emit(uinput.ABS_Y, ny, syn=True)
                time.sleep(0.002)
        except Exception as e:
            debug_linux(f"move_smooth failed: {e}")

    def mouse_move_rel(self, dx, dy):
        pos = self.get_cursor_pos()
        if pos is None:
            return
        x, y = pos
        self.set_cursor_pos(x + dx, y + dy)

    def mouse_wheel(self, delta):
        try:
            self.device.emit(uinput.REL_WHEEL, int(delta / 100), syn=True)
            time.sleep(0.01)
        except Exception as e:
            debug_linux(f"mouse_wheel failed: {e}")

    def find_onenote_window(self):
        return None

    def focus_window(self, hwnd):
        return None

    def client_to_screen(self, hwnd, x, y):
        return x, y


def create_input_controller():
    if IS_WINDOWS:
        return WindowsInputController()
    return UInputController()
