"""The drawing engine: turns trained strokes into real pointer movement.

Deliberately free of any GUI framework. Progress is reported through an
`emit(event, payload)` callback so any frontend can render it.
"""

import random
import time
from threading import Lock, Thread

from .input_controller import debug_linux
from .paths import IS_WINDOWS
from .text_parser import count_drawable_chars, parse_text_with_tables

# Emit at most one progress event per this many seconds, so a long text does
# not flood the frontend bridge with thousands of calls.
PROGRESS_INTERVAL = 0.12


class DrawingEngine:
    def __init__(self, controller, emit):
        self.input = controller
        self.emit = emit
        self._thread = None
        self._lock = Lock()
        self.stop_requested = False
        self.total_chars = 0
        self.drawn_chars = 0
        self._last_progress = 0.0
        self._ctx = {}
        self.letter_db = {}
        self._failsafe_check_count = 0  # Check every N strokes instead of every point

    # ------------------------------------------------------------------ state
    def is_running(self):
        return self._thread is not None and self._thread.is_alive()

    def stop(self):
        self.stop_requested = True

    # ------------------------------------------------------------------ start
    def start(self, text, letter_db, settings, origin_xy=None):
        """Validate and kick off drawing. Returns a dict describing the outcome."""
        with self._lock:
            if self.is_running():
                return {"ok": False, "error": "busy", "message": "Drawing is already running."}

            blocks = parse_text_with_tables(text)
            self.total_chars = count_drawable_chars(blocks)
            if self.total_chars == 0:
                return {"ok": False, "error": "empty",
                        "message": "There is nothing drawable in this text."}

            self.letter_db = letter_db
            self.drawn_chars = 0
            self._last_progress = 0.0
            self.stop_requested = False

            self._thread = Thread(target=self._run, args=(blocks, settings, origin_xy), daemon=True)
            self._thread.start()
            return {"ok": True, "total_chars": self.total_chars}

    # -------------------------------------------------------------- main loop
    def _run(self, blocks, settings, origin_xy=None):
        try:
            delay = int(settings.get("start_delay", 4))
            for remaining in range(delay, 0, -1):
                if self.stop_requested:
                    self.emit("done", {"stopped": True})
                    return
                self.emit("countdown", {"seconds": remaining})
                time.sleep(1)

            self.emit("started", {"total_chars": self.total_chars})

            origin = self._resolve_origin(origin_xy)
            if origin is None:
                return
            client_left, client_top = origin

            self._ctx = {
                "canvas_x": client_left,
                "canvas_y": client_top + (90 if IS_WINDOWS else 0),
                "line_spacing_px": settings.get("line_spacing", 60),
                "scale": settings.get("character_size", 0.1),
            }

            # On Linux, start exactly at the picked origin; on Windows, indent
            # into the OneNote client area.
            offset_x = 50 if IS_WINDOWS else 0
            offset_y = 50 if IS_WINDOWS else 0
            current_lines = 0
            for block in blocks:
                if self.stop_requested or self.failsafe():
                    break
                if block["type"] == "text":
                    offset_x, offset_y, current_lines = self._draw_text_block(
                        block["content"], offset_x, offset_y, current_lines)
                else:
                    offset_x, offset_y, current_lines = self._draw_table(
                        block["rows"], offset_x, offset_y, current_lines)
                    offset_y += self._ctx["line_spacing_px"]
                    current_lines += 1

            self.emit("progress", {"value": 1.0 if not self.stop_requested else self._fraction(),
                                   "char": ""})
            self.emit("done", {"stopped": self.stop_requested})
        except Exception as e:  # never let a background thread die silently
            debug_linux(f"engine crashed: {e}")
            self.emit("error", {"message": str(e)})
            self.emit("done", {"stopped": True})

    def _resolve_origin(self, origin_xy=None):
        """Where on screen the first character goes."""
        if IS_WINDOWS:
            hwnd = self.input.find_onenote_window()
            if not hwnd:
                self.emit("error", {"message": "OneNote window not found."})
                self.emit("done", {"stopped": True})
                return None
            self.input.focus_window(hwnd)
            time.sleep(0.4)
            return self.input.client_to_screen(hwnd, 0, 0)

        # On Wayland, querying the real cursor position is unreliable (XWayland
        # only sees it while the pointer is over an XWayland-mapped surface,
        # so pynput/xdotool can report a stale position). The frontend has the
        # user click the start point on a portal screenshot instead, and passes
        # it in as origin_xy. get_cursor_pos() is kept only as a last-resort
        # fallback for setups without the portal (e.g. plain X11 sessions).
        if origin_xy is not None:
            debug_linux(f"draw origin from picked point: {origin_xy[0]},{origin_xy[1]}")
            return origin_xy

        pos = self.input.get_cursor_pos()
        if pos is None:
            self.emit("error", {"message": "Cursor position unavailable. "
                                           "Install xdotool or check uinput permissions."})
            self.emit("done", {"stopped": True})
            return None
        debug_linux(f"draw origin from cursor (no picked point given): {pos[0]},{pos[1]}")
        return pos

    # -------------------------------------------------------------- failsafe
    def failsafe(self):
        """Check every 50 calls instead of every point (expensive on Linux)."""
        self._failsafe_check_count += 1
        if self._failsafe_check_count < 50:
            return False
        self._failsafe_check_count = 0

        pos = self.input.get_cursor_pos()
        if pos is None:
            return False
        x, y = pos
        if x < 10 and y < 10:
            self.stop_requested = True
            return True
        return False

    # ------------------------------------------------------------- geometry
    def _bounds(self, ch):
        strokes = self.letter_db[ch]
        min_x = min_y = float("inf")
        max_x = max_y = float("-inf")
        for stroke in strokes:
            for x, y in stroke:
                min_x = min(min_x, x)
                max_x = max(max_x, x)
                min_y = min(min_y, y)
                max_y = max(max_y, y)
        return min_x, min_y, max_x, max_y

    def letter_width(self, ch):
        min_x, _, max_x, _ = self._bounds(ch)
        return max_x - min_x

    def letter_height(self, ch):
        _, min_y, _, max_y = self._bounds(ch)
        return max_y - min_y

    # -------------------------------------------------------------- progress
    def _fraction(self):
        if not self.total_chars:
            return 0.0
        return min(self.drawn_chars / self.total_chars, 1.0)

    def _report(self, ch):
        now = time.monotonic()
        if now - self._last_progress < PROGRESS_INTERVAL:
            return
        self._last_progress = now
        self.emit("progress", {"value": self._fraction(), "char": ch})

    # --------------------------------------------------------- draw primitives
    def _draw_character(self, ch, offset_x, offset_y, scale, jitter=True):
        if ch not in self.letter_db:
            return 0

        self.drawn_chars += 1
        self._report(ch)

        raw_width = self.letter_width(ch)
        letter_scale = scale * random.uniform(0.85, 1.0) if jitter else scale
        effective_width = max(raw_width, 60)
        letter_spacing = max(
            int(effective_width * letter_scale * 1.23) + 8,
            int(50 * letter_scale) + 8,
        )

        canvas_x = self._ctx["canvas_x"]
        canvas_y = self._ctx["canvas_y"]
        offset_letter_y = random.randint(-8, 8) if jitter else 0

        for stroke in self.letter_db[ch]:
            if self.stop_requested or self.failsafe() or len(stroke) < 2:
                continue

            start_x, start_y = stroke[0]
            start_y += offset_letter_y
            sx = canvas_x + int(start_x * letter_scale) + int(offset_x)
            sy = canvas_y + int(start_y * letter_scale) + int(offset_y)

            self.input.set_cursor_pos(int(sx), int(sy))
            # No sleep before starting to draw

            last_x, last_y = sx, sy
            self.input.mouse_down()

            for x, y in stroke[1::3]:
                if self.stop_requested or self.failsafe():
                    break
                y_off = random.randint(-2, 2) if jitter else 0
                nx = canvas_x + int(x * letter_scale) + int(offset_x)
                ny = canvas_y + int((y + y_off) * letter_scale) + int(offset_y)
                if jitter:
                    nx += random.randint(-1, 1)
                    ny += random.randint(-1, 1)
                self._move_to(last_x, last_y, nx, ny)
                last_x, last_y = nx, ny

            self.input.mouse_up()

        return letter_spacing

    def _move_to(self, last_x, last_y, nx, ny):
        """One pen segment, assuming the button is already held down for
        this whole stroke (pressed once at its start, released once at its
        end — see _draw_character/_draw_line_abs)."""
        if IS_WINDOWS:
            self.input.mouse_move_rel(int(nx - last_x), int(ny - last_y))
        elif hasattr(self.input, "move_smooth"):
            self.input.move_smooth(last_x, last_y, nx, ny)
        else:
            self.input.set_cursor_pos(nx, ny)

    def _scroll_if_needed(self, current_lines):
        if current_lines >= 5:
            debug_linux(f"scrolling at line {current_lines}")
            self.emit("scroll", {})
            self.input.mouse_wheel(-1000)
            time.sleep(1)
            return 0, 50
        return current_lines, None

    def _draw_text_block(self, text, offset_x, offset_y, current_lines):
        scale = self._ctx["scale"]
        line_spacing_px = self._ctx["line_spacing_px"]
        chars_per_line = int(6 / scale)
        base_char_spacing_px = self._space_width_px(scale)

        for line in text.split("\n"):
            if self.stop_requested or self.failsafe():
                break
            offset_x = 50
            chars_in_line = 0

            for ch in line:
                if self.stop_requested or self.failsafe():
                    break
                if chars_in_line >= chars_per_line and ch == " ":
                    current_lines += 1
                    current_lines, new_y = self._scroll_if_needed(current_lines)
                    offset_y = new_y if new_y is not None else offset_y + line_spacing_px
                    offset_x = 50
                    chars_in_line = 0
                    time.sleep(0.7)
                    continue
                if ch == " ":
                    offset_x += base_char_spacing_px
                    chars_in_line += 1
                    continue
                offset_x += self._draw_character(ch, offset_x, offset_y, scale, jitter=True)
                chars_in_line += 1

            offset_y += line_spacing_px
            current_lines += 1
            current_lines, new_y = self._scroll_if_needed(current_lines)
            if new_y is not None:
                offset_y = new_y

        return offset_x, offset_y, current_lines

    def _draw_line_abs(self, x1, y1, x2, y2):
        steps = max(abs(x2 - x1), abs(y2 - y1), 1)
        coarse_steps = max(steps // 3, 1)
        self.input.set_cursor_pos(int(x1), int(y1))
        # No pre-sleep
        last_x, last_y = x1, y1
        self.input.mouse_down()

        for i in range(1, coarse_steps + 1):
            nx = x1 + int((x2 - x1) * i / coarse_steps)
            ny = y1 + int((y2 - y1) * i / coarse_steps)
            if i < coarse_steps:  # wobble slightly so ruled lines look drawn
                if abs(x2 - x1) > abs(y2 - y1):
                    ny += random.randint(-2, 2)
                else:
                    nx += random.randint(-2, 2)
            self._move_to(last_x, last_y, nx, ny)
            last_x, last_y = nx, ny

        self.input.mouse_up()

    @staticmethod
    def _space_width_px(scale):
        # Wider than a typical inter-letter gap, so a space reads as an
        # actual word break rather than a tight letter gap.
        return int(320 * scale) + 20

    def _measure_col_widths(self, padded_rows, scale):
        col_count = len(padded_rows[0])
        col_widths = []
        for col_idx in range(col_count):
            max_cell_width = 0
            for row in padded_rows:
                cell_px = 0
                for ch in str(row[col_idx]):
                    if ch == " ":
                        cell_px += self._space_width_px(scale)
                    elif ch in self.letter_db:
                        effective_w = max(self.letter_width(ch), 60)
                        cell_px += max(int(effective_w * scale * 1.23) + 8,
                                       int(50 * scale) + 8)
                max_cell_width = max(max_cell_width, cell_px)
            col_widths.append(max(max_cell_width + 40, 80))
        return col_widths

    def _measure_row_height(self, scale):
        heights = [self.letter_height(ch) for ch in self.letter_db
                   if self.letter_height(ch) > 0]
        avg_letter_h = (sum(heights) / len(heights)) if heights else 200
        return max(int(avg_letter_h * scale) + 40, 60)

    def _draw_table(self, table, offset_x, offset_y, current_lines):
        if not table:
            return offset_x, offset_y, current_lines

        scale = self._ctx["scale"]
        canvas_x = self._ctx["canvas_x"]
        canvas_y = self._ctx["canvas_y"]
        col_count = max(len(row) for row in table)
        padded_rows = [row + [""] * (col_count - len(row)) for row in table]
        col_widths = self._measure_col_widths(padded_rows, scale)
        row_height = self._measure_row_height(scale)
        table_width = sum(col_widths)
        origin_sx = canvas_x + int(offset_x)
        origin_sy = canvas_y + int(offset_y)

        for row_idx, row in enumerate(padded_rows):
            if self.stop_requested or self.failsafe():
                break
            row_top_sy = origin_sy + row_idx * row_height
            row_bottom_sy = row_top_sy + row_height
            if row_idx == 0:
                self._draw_line_abs(origin_sx, row_top_sy, origin_sx + table_width, row_top_sy)
            self._draw_line_abs(origin_sx, row_bottom_sy, origin_sx + table_width, row_bottom_sy)

            col_sx = origin_sx
            for col_idx, cell in enumerate(row):
                if self.stop_requested or self.failsafe():
                    break
                self._draw_line_abs(col_sx, row_top_sy, col_sx, row_bottom_sy)
                char_offset_x = col_sx + 15 - canvas_x
                char_offset_y = row_top_sy + 12 - canvas_y
                for ch in str(cell):
                    if self.stop_requested or self.failsafe():
                        break
                    if ch == " ":
                        char_offset_x += self._space_width_px(scale)
                    else:
                        char_offset_x += self._draw_character(
                            ch, char_offset_x, char_offset_y, scale, jitter=False)
                col_sx += col_widths[col_idx]

            self._draw_line_abs(origin_sx + table_width, row_top_sy,
                                origin_sx + table_width, row_bottom_sy)
            current_lines += 1

        return 50, offset_y + len(padded_rows) * row_height, current_lines
