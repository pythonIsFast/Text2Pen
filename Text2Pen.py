from tkinter import Canvas, Frame, Text, Scrollbar, Toplevel, StringVar, ROUND, Spinbox, filedialog
from tkinter import ttk
from threading import Thread
import customtkinter as ctk
import time
import win32gui
import win32con
import win32api
import json
import os
import random
import psutil
import sys
import requests
import platform
import webbrowser
import ctypes
import base64
from PIL import Image
import io

TELEMETRY_URL = "https://text2pen-backend.onrender.com/telemetry"
AI_URL = "https://text2pen-backend.onrender.com/ai"
AI_CHAT_URL = "https://text2pen-backend.onrender.com/ai/chat"
AI_LIMITS_URL = "https://text2pen-backend.onrender.com/ai/limits"

INSTALL_DIR = os.path.join(os.environ["LOCALAPPDATA"], "Text2Pen")
UPDATE_EXE = os.path.join(INSTALL_DIR, "Update.exe")

def parse_text_with_tables(text):
    lines = text.split('\n')
    blocks = []
    current_table = []
    current_text = []

    def flush_text():
        nonlocal current_text
        if current_text:
            blocks.append({"type": "text", "content": "\n".join(current_text)})
            current_text = []

    def flush_table():
        nonlocal current_table
        if current_table:
            blocks.append({"type": "table", "rows": current_table})
            current_table = []

    for line in lines:
        if '\t' in line:
            flush_text()
            current_table.append(line.split('\t'))
        else:
            flush_table()
            current_text.append(line)

    flush_text()
    flush_table()
    return blocks

def is_update_running():
    for proc in psutil.process_iter(["name", "exe"]):
        try:
            if proc.info["exe"] and os.path.normcase(proc.info["exe"]) == os.path.normcase(UPDATE_EXE):
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return False

if is_update_running():
    print("Update is running, exiting Text2Pen to allow update...")
    time.sleep(2)
    sys.exit(0)

if os.path.exists(os.path.join(INSTALL_DIR, "Update.exe-newest")):
    os.replace(os.path.join(INSTALL_DIR, "Update.exe-newest"), os.path.join(INSTALL_DIR, "Update.exe"))

if len(sys.argv) > 1 and sys.argv[1] == "updaterStart":
    sys.exit(0)


class AIChatSidebar:
    """Side panel chat interface with model selection, image upload, and limit display."""

    def __init__(self, parent_app):
        self.app = parent_app
        self.root = parent_app.root
        self._parent_frame = parent_app.root  # overridden by _toggle_ai_sidebar
        self.sidebar = None
        self.chat_history = []
        self.attached_image_b64 = None
        self.attached_image_path = None
        self.attached_image_thumb = None
        self.is_open = False
        self.limits = {"chat_used": 0, "chat_limit": 10, "text_used": 0, "text_limit": 1000}

    # ------------------------------------------------------------------ open/close
    def toggle(self):
        if self.is_open:
            self.close()
        else:
            self.open()

    def open(self):
        if self.is_open:
            return
        self.is_open = True
        self._build()
        self._refresh_limits()

    def close(self):
        if not self.is_open:
            return
        self.is_open = False
        if self.sidebar and self.sidebar.winfo_exists():
            self.sidebar.destroy()
        self.sidebar = None
        # reset the toggle button label if accessible
        try:
            self.app.ai_btn.configure(text="✨ AI Chat")
        except Exception:
            pass

    # ------------------------------------------------------------------ build UI
    def _build(self):
        self.sidebar = ctk.CTkFrame(self._parent_frame, width=360, corner_radius=0)
        self.sidebar.pack(side=ctk.RIGHT, fill=ctk.Y, padx=0, pady=0)
        self.sidebar.pack_propagate(False)

        # ── Header ──────────────────────────────────────────────────────
        header = ctk.CTkFrame(self.sidebar, fg_color="#7719AA", corner_radius=0)
        header.pack(fill=ctk.X)

        ctk.CTkLabel(header, text="✨ Text2Pen AI Chat",
                     font=('Arial', 15, 'bold'), text_color="white").pack(side=ctk.LEFT, padx=12, pady=10)

        ctk.CTkButton(header, text="✕", width=32, height=32,
                      fg_color="transparent", text_color="white",
                      hover_color="#9b2cc7", font=('Arial', 14, 'bold'),
                      command=self.close).pack(side=ctk.RIGHT, padx=8, pady=8)

        # ── Limit bar ───────────────────────────────────────────────────
        self.limit_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.limit_frame.pack(fill=ctk.X, padx=10, pady=(6, 2))
        self.limit_label = ctk.CTkLabel(self.limit_frame, text="Loading limits...",
                                        font=('Arial', 10), text_color="gray")
        self.limit_label.pack(anchor='w')

        # ── Chat messages area ──────────────────────────────────────────
        self.msg_frame = ctk.CTkScrollableFrame(self.sidebar)
        self.msg_frame.pack(fill=ctk.BOTH, expand=True, padx=6, pady=4)

        # ── Image preview strip ─────────────────────────────────────────
        self.img_strip = ctk.CTkFrame(self.sidebar, height=60, fg_color="transparent")
        self.img_strip.pack(fill=ctk.X, padx=8)
        self.img_strip.pack_forget()   # hidden until image attached

        self.img_thumb_label = ctk.CTkLabel(self.img_strip, text="")
        self.img_thumb_label.pack(side=ctk.LEFT, padx=4)

        self.img_name_label = ctk.CTkLabel(self.img_strip, text="", font=('Arial', 10))
        self.img_name_label.pack(side=ctk.LEFT, padx=4)

        ctk.CTkButton(self.img_strip, text="✕", width=24, height=24,
                      fg_color="#e74c3c", text_color="white", font=('Arial', 10),
                      command=self._clear_image).pack(side=ctk.RIGHT, padx=4)

        # ── Input row ───────────────────────────────────────────────────
        input_frame = ctk.CTkFrame(self.sidebar, corner_radius=0)
        input_frame.pack(fill=ctk.X, padx=0, pady=0)

        # image upload button
        ctk.CTkButton(input_frame, text="🖼", width=36, height=36,
                      fg_color="transparent", text_color_disabled="gray",
                      hover_color="#3a3a3a", font=('Arial', 16),
                      command=self._attach_image).pack(side=ctk.LEFT, padx=(6, 2), pady=6)

        self.input_box = ctk.CTkTextbox(input_frame, height=60, wrap='word', font=('Arial', 12))
        self.input_box.pack(side=ctk.LEFT, fill=ctk.X, expand=True, padx=4, pady=6)
        self.input_box.bind("<Return>", self._on_enter)
        self.input_box.bind("<Shift-Return>", lambda e: None)  # allow newline with shift

        ctk.CTkButton(input_frame, text="➤", width=36, height=36,
                      fg_color="#7719AA", text_color="white", font=('Arial', 16),
                      command=self._send).pack(side=ctk.LEFT, padx=(2, 8), pady=6)

        # ── Welcome message ─────────────────────────────────────────────
        if not self.chat_history:
            self._add_bubble("assistant",
                "Hello! I'm your Text2Pen AI Assistant. 😊\n\n"
                "I can help you with:\n"
                "• Correcting or rewriting texts\n"
                "• Analyzing images (🖼 attach an image)\n"
                "• Answering questions\n\n"
                "If I suggest a text, you can insert it directly into the text field.",
                show_actions=False)

    # ------------------------------------------------------------------ limits
    def _refresh_limits(self):
        def fetch():
            try:
                resp = requests.get(AI_LIMITS_URL, timeout=10)
                if resp.status_code == 200:
                    data = resp.json()
                    self.limits = data
                    self.root.after(0, self._update_limit_label)
            except Exception:
                pass
        Thread(target=fetch, daemon=True).start()

    def _update_limit_label(self):
        try:
            if not self.sidebar or not self.sidebar.winfo_exists():
                return
            if not self.limit_label.winfo_exists():
                return
            chat_used = self.limits.get("chat_used", "?")
            chat_lim  = self.limits.get("chat_limit", 10)
            txt_used  = self.limits.get("text_used", "?")
            txt_lim   = self.limits.get("text_limit", 1000)
            self.limit_label.configure(
                text=f"💬 Chat: {chat_used}/{chat_lim} today   |   📝 Text-Actions: {txt_used}/{txt_lim} today"
            )
        except Exception:
            pass

    # ------------------------------------------------------------------ image handling
    def _attach_image(self):
        path = filedialog.askopenfilename(
            title="Select Image",
            filetypes=[("Images", "*.png *.jpg *.jpeg *.webp *.gif *.bmp"), ("All", "*.*")]
        )
        if not path:
            return
        try:
            with open(path, "rb") as f:
                raw = f.read()
            self.attached_image_b64 = base64.b64encode(raw).decode()
            self.attached_image_path = path

            # thumbnail
            img = Image.open(io.BytesIO(raw))
            img.thumbnail((48, 48))
            self.attached_image_thumb = ctk.CTkImage(light_image=img, dark_image=img, size=(img.width, img.height))
            self.img_thumb_label.configure(image=self.attached_image_thumb)
            self.img_name_label.configure(text=os.path.basename(path)[:28])
            self.img_strip.pack(fill=ctk.X, padx=8, before=self.input_box.master)
        except Exception as e:
            self._add_bubble("assistant", f"Error loading image: {e}", show_actions=False)

    def _clear_image(self):
        self.attached_image_b64 = None
        self.attached_image_path = None
        self.attached_image_thumb = None
        self.img_strip.pack_forget()

    # ------------------------------------------------------------------ send
    def _on_enter(self, event):
        # Shift+Enter = newline, Enter alone = send
        if not (event.state & 0x1):  # no shift
            self._send()
            return "break"

    def _send(self):
        text = self.input_box.get("1.0", "end-1c").strip()
        if not text and not self.attached_image_b64:
            return

        # show user bubble
        display = text or "(Image)"
        self._add_bubble("user", display, image_b64=self.attached_image_b64, show_actions=False)

        # build history entry
        user_entry = {"role": "user", "content": text}
        if self.attached_image_b64:
            user_entry["image_b64"] = self.attached_image_b64
        self.chat_history.append(user_entry)

        image_b64 = self.attached_image_b64
        has_image = image_b64 is not None

        # clear input
        self.input_box.delete("1.0", "end")
        self._clear_image()

        # typing indicator with animated dots
        typing_bubble = self._add_typing_indicator(has_image)

        def run():
            try:
                payload = {
                    "history": self.chat_history[:-1],
                    "message": text,
                    "image_b64": image_b64,
                }
                # timeout=(connect_timeout, read_timeout)
                # Connect must succeed in 10s (server is always up).
                # Read timeout: None = wait forever for Backend to respond.
                # The backend itself has a 90s hard limit, so it will always reply eventually.
                resp = requests.post(AI_CHAT_URL, json=payload, timeout=(10, None))
                if resp.status_code == 200:
                    data = resp.json()
                    reply = data.get("result", "")
                    queue_pos = data.get("queue_position")
                    self.root.after(0, lambda: self._on_reply(reply, typing_bubble, queue_pos))
                elif resp.status_code == 429:
                    self.root.after(0, lambda: self._on_error(
                        "⛔ Daily limit reached! Please try again tomorrow.", typing_bubble))
                elif resp.status_code == 503:
                    data = resp.json()
                    pos = data.get("queue_position", "?")
                    self.root.after(0, lambda: self._on_reply(
                        f"⏳ You are in the queue (position {pos}). Please wait...",
                        typing_bubble, pos))
                else:
                    self.root.after(0, lambda: self._on_error(
                        f"❌ Error: {resp.status_code}", typing_bubble))
            except requests.exceptions.Timeout:
                self.root.after(0, lambda: self._on_error(
                    "❌ Could not connect to the server (timeout after 10s).\n"
                    "Check your internet connection.",
                    typing_bubble))
            except requests.exceptions.ConnectionError:
                self.root.after(0, lambda: self._on_error(
                    "❌ Could not reach the server. Check your internet connection.",
                    typing_bubble))
            except Exception as e:
                self.root.after(0, lambda: self._on_error(
                    f"❌ Unexpected error: {e}", typing_bubble))

        Thread(target=run, daemon=True).start()

    def _on_reply(self, reply, typing_bubble, queue_pos=None):
        self._remove_bubble(typing_bubble)
        try:
            if not self.msg_frame or not self.msg_frame.winfo_exists():
                return
        except Exception:
            return
        self.chat_history.append({"role": "assistant", "content": reply})
        has_text_suggestion = self._detect_text_suggestion(reply)
        self._add_bubble("assistant", reply, show_actions=has_text_suggestion)
        self._refresh_limits()

    def _on_error(self, msg, typing_bubble):
        self._remove_bubble(typing_bubble)
        try:
            if not self.msg_frame or not self.msg_frame.winfo_exists():
                return
        except Exception:
            return
        self._add_bubble("assistant", msg, show_actions=False)

    # ------------------------------------------------------------------ text suggestion detection
    def _detect_text_suggestion(self, reply: str) -> bool:
        """Return True if reply looks like it contains a text suggestion."""
        keywords = ["here is", "here's", "suggestion:", "text:", "corrected:", "revised:",
                    "version:", "rewritten:", "```"]
        lower = reply.lower()
        return any(kw in lower for kw in keywords) or len(reply) > 80

    # ------------------------------------------------------------------ bubble helpers
    def _add_bubble(self, role, text, image_b64=None, show_actions=False):
        is_user = role == "user"
        bubble_color = "#7719AA" if is_user else ("gray20" if ctk.get_appearance_mode() == "Dark" else "#f0f0f0")
        text_color   = "white" if is_user else ("white" if ctk.get_appearance_mode() == "Dark" else "black")
        anchor = "e" if is_user else "w"

        outer = ctk.CTkFrame(self.msg_frame, fg_color="transparent")
        outer.pack(fill=ctk.X, pady=3, padx=4)

        bubble = ctk.CTkFrame(outer, fg_color=bubble_color, corner_radius=12)
        bubble.pack(anchor=anchor, padx=4)

        # image thumbnail inside bubble
        if image_b64 and is_user:
            try:
                raw = base64.b64decode(image_b64)
                img = Image.open(io.BytesIO(raw))
                img.thumbnail((120, 120))
                photo = ctk.CTkImage(light_image=img, dark_image=img, size=(img.width, img.height))
                lbl = ctk.CTkLabel(bubble, image=photo, text="")
                lbl.image = photo  # keep reference
                lbl.pack(padx=8, pady=(8, 2))
            except Exception:
                pass

        ctk.CTkLabel(bubble, text=text, font=('Arial', 12), text_color=text_color,
                     wraplength=270, justify='left').pack(padx=10, pady=(6, 4))

        if show_actions and not is_user:
            action_row = ctk.CTkFrame(bubble, fg_color="transparent")
            action_row.pack(fill=ctk.X, padx=8, pady=(0, 6))

            ctk.CTkButton(action_row, text="📋 Insert into text field",
                          width=150, height=26,
                          fg_color="#4ecdc4", text_color="white",
                          font=('Arial', 11),
                          command=lambda t=text: self._insert_to_textfield(t)
                          ).pack(side=ctk.LEFT, padx=2)

            ctk.CTkButton(action_row, text="❌ Dismiss",
                          width=90, height=26,
                          fg_color="#95a5a6", text_color="white",
                          font=('Arial', 11),
                          command=lambda f=outer: self._dismiss_bubble_actions(f)
                          ).pack(side=ctk.LEFT, padx=2)

        # scroll to bottom
        def _scroll():
            try:
                if self.msg_frame.winfo_exists():
                    self.msg_frame._parent_canvas.yview_moveto(1.0)
            except Exception:
                pass
        self.root.after(50, _scroll)
        return outer

    def _add_typing_indicator(self, has_image=False):
        outer = ctk.CTkFrame(self.msg_frame, fg_color="transparent")
        outer.pack(fill=ctk.X, pady=3, padx=4)
        bubble = ctk.CTkFrame(outer,
                              fg_color="gray25" if ctk.get_appearance_mode() == "Dark" else "#e8e8e8",
                              corner_radius=12)
        bubble.pack(anchor="w", padx=4)

        hint = "⏳ Analyzing image" if has_image else "⏳ Thinking"
        lbl = ctk.CTkLabel(bubble, text=hint + ".",
                           font=('Arial', 12, 'italic'), text_color="gray")
        lbl.pack(padx=10, pady=4)

        if has_image:
            ctk.CTkLabel(bubble,
                         text="(First request may take ~30s\nwhile the server wakes up)",
                         font=('Arial', 10), text_color="gray").pack(padx=10, pady=(0, 6))

        # Per-bubble animation state — avoids conflicts with multiple concurrent requests
        state = {"active": True, "dots": "."}

        def animate():
            if not state["active"]:
                return
            try:
                if not lbl.winfo_exists():
                    return
                state["dots"] = state["dots"] + "." if len(state["dots"]) < 3 else "."
                lbl.configure(text=hint + state["dots"])
                self.root.after(500, animate)
            except Exception:
                pass

        # Store stop function on the frame so _remove_bubble can stop the animation
        outer._stop_animation = lambda: state.update({"active": False})

        self.root.after(500, animate)

        def _scroll():
            try:
                if self.msg_frame.winfo_exists():
                    self.msg_frame._parent_canvas.yview_moveto(1.0)
            except Exception:
                pass
        self.root.after(50, _scroll)
        return outer

    def _remove_bubble(self, bubble_frame):
        try:
            if bubble_frame and bubble_frame.winfo_exists():
                stop_fn = getattr(bubble_frame, "_stop_animation", None)
                if stop_fn:
                    stop_fn()
                bubble_frame.destroy()
        except Exception:
            pass

    def _dismiss_bubble_actions(self, outer_frame):
        # Remove just the action buttons row (last child of bubble)
        try:
            bubble = outer_frame.winfo_children()[0]
            children = bubble.winfo_children()
            if children:
                children[-1].destroy()
        except Exception:
            pass

    def _insert_to_textfield(self, ai_text):
        """Replace the main text input with the AI-suggested text."""
        # Strip markdown-style code fences if present
        clean = ai_text
        if "```" in clean:
            parts = clean.split("```")
            # take the first code block content if present
            if len(parts) >= 3:
                clean = parts[1].strip()
                if '\n' in clean:
                    clean = clean[clean.index('\n')+1:]  # remove language tag line
            else:
                clean = clean.replace("```", "").strip()

        try:
            self.app.input_text.delete('1.0', 'end')
            self.app.input_text.insert('1.0', clean)
            # confirm in chat
            self._add_bubble("assistant", "✅ Text inserted into the text field!", show_actions=False)
        except AttributeError:
            self._add_bubble("assistant", "⚠️ Text field not available (still in learning mode).", show_actions=False)


class LetterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Text2Pen")
        
        self.letter_db = {}
        self.current_letter = 'a'
        self.learning_mode = True
        self.alphabet = "abcdefghijklmnopqrstuvwxyzäöüÄÖÜßABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.,!?;:-_—()[]{}<>\"'/*+=@#$%^&|~`"
        self.current_stroke = []
        self.stop_drawing = False
        
        self.db_file = 'letter_db.json'
        self.settings_file = 'settingsDB.json'
        self.text_file = 'saved_text.txt'
        
        self.load_letters()
        self.load_settings()
        
        if not os.path.exists(self.settings_file) or 'telemetry_opted_in' not in self.settings:
            self.show_telemetry_dialog()

        self.root.report_callback_exception = self.tk_exception_handler
        
        self.line_spacing_value = 60
        self.character_size_value = 0.1
        self.line_spacing = None
        self.characterSize = None
        
        if len(self.letter_db) == len(self.alphabet):
            self.learning_mode = False
        
        # AI chat sidebar instance
        self.ai_sidebar = AIChatSidebar(self)

        self.create_gui()
        self.root._state_before_windows_set_titlebar_color = 'zoomed'

    # ------------------------------------------------------------------ progress overlay
    def show_progress_overlay(self):
        screen_w = ctypes.windll.user32.GetSystemMetrics(0)
        width, height = 300, 30
        self.overlay = Toplevel(self.root)
        self.overlay.overrideredirect(True)
        self.overlay.attributes("-topmost", True)
        self.overlay.resizable(False, False)
        self.overlay.geometry(f"{width}x{height}+{screen_w - width}+0")
        self.overlay_bar = ctk.CTkProgressBar(self.overlay, width=width, height=height,
                                               corner_radius=0, progress_color="green")
        self.overlay_bar.pack()
        self.overlay_bar.set(0)

    def close_progress_overlay(self):
        if hasattr(self, 'overlay'):
            self.root.after(0, self.overlay.destroy)

    # ------------------------------------------------------------------ telemetry
    def sanitize_exception(self, e):
        msg = str(e)
        msg = msg.replace(os.environ.get("USERNAME", ""), "<user>")
        return msg[:500]

    def tk_exception_handler(self, exc, val, tb):
        print("TK EXCEPTION:", val)
        self.expeption_occured(val)

    def expeption_occured(self, exception):
        if not self.telemetry_opted_in:
            return
        Thread(target=self.send_error_event, args=(exception,), daemon=True).start()
    
    def send_error_event(self, exception):
        if not self.telemetry_opted_in:
            return
        event = {
            "event": "ERROR", "os": platform.system(),
            "timestamp": time.time(),
            "extra": {"Exact Error": self.sanitize_exception(exception)}
        }
        try:
            resp = requests.post(TELEMETRY_URL, json=event, timeout=60)
            return "success" if resp.status_code == 200 else "failed"
        except Exception:
            return "failed"

    # ------------------------------------------------------------------ text persistence
    def load_text(self):
        if os.path.exists(self.text_file):
            try:
                with open(self.text_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                self.input_text.delete('1.0', 'end')
                self.input_text.insert('1.0', content)
            except IOError:
                pass

    def save_text(self):
        try:
            content = self.input_text.get('1.0', 'end-1c')
            with open(self.text_file, 'w', encoding='utf-8') as f:
                f.write(content)
        except IOError:
            pass

    def on_closing(self):
        self.save_text()
        self.save_settings()
        self.root.destroy()

    # ------------------------------------------------------------------ old AI actions (text buttons)
    def call_ai(self, action, extra=None):
        text = self.input_text.get('1.0', 'end-1c')
        if not text.strip():
            return
        self.status_label.configure(text="🤖 AI processing...")

        def run():
            try:
                payload = {"action": action, "text": text}
                if extra:
                    payload["extra"] = extra
                resp = requests.post(AI_URL, json=payload, timeout=30)
                if resp.status_code == 200:
                    result = resp.json()["result"]
                    self.root.after(0, lambda: self._apply_ai_result(result))
                elif resp.status_code == 429:
                    self.root.after(0, lambda: self.status_label.configure(text="⛔ Daily limit reached!"))
                else:
                    self.root.after(0, lambda: self.status_label.configure(text="❌ AI error!"))
            except Exception:
                self.root.after(0, lambda: self.status_label.configure(text="❌ Connection error!"))

        Thread(target=run, daemon=True).start()

    def _apply_ai_result(self, result):
        self.input_text.delete('1.0', 'end')
        self.input_text.insert('1.0', result)
        self.status_label.configure(text="✅ AI done!")

    # ------------------------------------------------------------------ GUI creation
    def create_gui(self):
        if self.learning_mode:
            self.create_learning_gui()
        else:
            self.create_text_gui()
    
    def create_learning_gui(self):
        self.title_label = ctk.CTkLabel(self.root,
                                         text=f"Learning character: {self.current_letter}",
                                         font=('Arial', 24, 'bold'))
        self.title_label.pack(pady=20)
        
        ctk.CTkLabel(self.root, text="Draw the character shown in the background.",
                     font=('Arial', 14)).pack(pady=5)
        
        self.canvas = Canvas(self.root, width=700, height=450, bg='white',
                             highlightthickness=2, highlightbackground='#cccccc')
        self.canvas.pack(pady=15, padx=20, fill=ctk.BOTH, expand=True)
        self.draw_letter_template(self.current_letter)
        
        self.canvas.bind('<Button-1>', self.start_drawing)
        self.canvas.bind('<B1-Motion>', self.draw)
        self.canvas.bind('<ButtonRelease-1>', self.finish_stroke)
        
        button_frame = ctk.CTkFrame(self.root)
        button_frame.pack(pady=20)
        
        ctk.CTkButton(button_frame, text="Delete", command=self.delete_learning,
                      fg_color='#ff6b6b', text_color='white', font=('Arial', 14), width=120
                      ).pack(side=ctk.LEFT, padx=10)
        ctk.CTkButton(button_frame, text="Save character", command=self.save_letter,
                      fg_color='#4ecdc4', text_color='white', font=('Arial', 14), width=150
                      ).pack(side=ctk.LEFT, padx=10)
        ctk.CTkButton(button_frame, text="Skip", command=self.next_letter,
                      fg_color='#95a5a6', text_color='white', font=('Arial', 14), width=120
                      ).pack(side=ctk.LEFT, padx=10)
        
        self.progress_label = ctk.CTkLabel(
            self.root,
            text=f"Progress: {len(self.letter_db)}/{len(self.alphabet)} characters learned!",
            font=('Arial', 12))
        self.progress_label.pack(pady=10)
        self.strokes = []
    
    def create_text_gui(self):
        # ── Outer wrapper so sidebar sits alongside main content ────────
        self.main_wrapper = ctk.CTkFrame(self.root, fg_color="transparent", corner_radius=0)
        self.main_wrapper.pack(fill=ctk.BOTH, expand=True)

        # ── Content area (left of sidebar) ──────────────────────────────
        self.content_area = ctk.CTkFrame(self.main_wrapper, fg_color="transparent", corner_radius=0)
        self.content_area.pack(side=ctk.LEFT, fill=ctk.BOTH, expand=True)

        # ── Header ──────────────────────────────────────────────────────
        header_frame = ctk.CTkFrame(self.content_area)
        header_frame.pack(fill=ctk.X, padx=20, pady=15)
        
        ctk.CTkLabel(header_frame, text="Text2Pen - Write text in OneNote",
                     font=('Arial', 26, 'bold')).pack(side=ctk.LEFT, expand=True)

        ctk.CTkButton(header_frame, text="🌐 Website",
                      command=lambda: webbrowser.open("https://text2pen.onrender.com"),
                      fg_color="#7719AA", text_color="white",
                      font=('Arial', 14, 'bold'), width=130).pack(side=ctk.LEFT, padx=4, pady=10)

        ctk.CTkButton(header_frame, text="📊 Status",
                      command=lambda: webbrowser.open("https://m9rgm51d.status.cron-job.org"),
                      fg_color="#04FF00", text_color="black",
                      font=('Arial', 14, 'bold'), width=110).pack(side=ctk.LEFT, padx=4, pady=10)

        # AI Chat toggle button
        self.ai_btn = ctk.CTkButton(
            header_frame, text="✨ AI Chat",
            command=self._toggle_ai_sidebar,
            fg_color="#7719AA", text_color="white",
            font=('Arial', 14, 'bold'), width=120)
        self.ai_btn.pack(side=ctk.LEFT, padx=4, pady=10)

        ctk.CTkButton(header_frame, text="⚙️", command=self.open_settings,
                      fg_color='#95a5a6', text_color='white',
                      font=('Arial', 18), width=50).pack(side=ctk.RIGHT, padx=10)
        
        # ── Scrollable main content ──────────────────────────────────────
        main_frame = ctk.CTkScrollableFrame(self.content_area)
        main_frame.pack(fill=ctk.BOTH, expand=True, padx=20, pady=10)
        
        info_frame = ctk.CTkFrame(main_frame)
        info_frame.pack(pady=15, fill=ctk.X)
        ctk.CTkLabel(info_frame, text="Enter text below and press 'Draw in OneNote'!",
                     font=('Arial', 13)).pack(anchor='w')
        ctk.CTkLabel(info_frame, text="🛑 FAILSAFE: Move mouse to the top left corner to stop!",
                     font=('Arial', 11), text_color='red').pack(anchor='w', pady=5)
        
        # Text label + AI dropdown button row
        top_row = ctk.CTkFrame(main_frame, fg_color="transparent")
        top_row.pack(fill=ctk.X, pady=(15, 0))
        ctk.CTkLabel(top_row, text="Text:", font=('Arial', 13, 'bold')).pack(side=ctk.LEFT)

        # AI quick-action dropdown (old button menu)
        def show_ai_menu():
            menu = Toplevel(self.root)
            menu.overrideredirect(True)
            menu.attributes("-topmost", True)
            btn_x = quick_ai_btn.winfo_rootx()
            btn_y = quick_ai_btn.winfo_rooty() + quick_ai_btn.winfo_height()
            menu.geometry(f"210x290+{btn_x - 140}+{btn_y}")
            is_dark = ctk.get_appearance_mode() == "Dark"
            fg = "white" if is_dark else "black"
            hover = "#3d3d3d" if is_dark else "#d0d0d0"
            bg = "#2b2b2b" if is_dark else "#f0f0f0"
            menu.configure(bg=bg)
            actions = [
                ("Spell check",        lambda: (menu.destroy(), self.call_ai("correct"))),
                ("Shorten",            lambda: (menu.destroy(), self.call_ai("shorten"))),
                ("Make formal",        lambda: (menu.destroy(), self.call_ai("tone_formal"))),
                ("Make casual",        lambda: (menu.destroy(), self.call_ai("tone_casual"))),
                ("Generate text",      lambda: (menu.destroy(), self.call_ai("generate"))),
                ("Fix unknown chars",  lambda: (menu.destroy(), self.call_ai("replace_unknown", {
                    "known_chars": "".join(self.letter_db.keys())}))),
            ]
            for label, cmd in actions:
                ctk.CTkButton(menu, text=label, command=cmd,
                              fg_color="transparent", text_color=fg,
                              hover_color=hover, anchor="w",
                              font=('Arial', 12), height=32).pack(fill=ctk.X, padx=2)
            menu.bind("<FocusOut>", lambda e: menu.destroy())
            menu.focus_set()

        quick_ai_btn = ctk.CTkButton(top_row, text="✨ Quick AI",
                                      command=show_ai_menu,
                                      fg_color="#7719AA", text_color="white",
                                      font=('Arial', 11, 'bold'), width=110, height=26)
        quick_ai_btn.pack(side=ctk.RIGHT, padx=5)
        
        text_frame = ctk.CTkFrame(main_frame)
        text_frame.pack(fill=ctk.BOTH, expand=True, pady=6)
        
        self.input_text = Text(text_frame, font=('Arial', 12), wrap='word', height=15)
        self.input_text.pack(side=ctk.LEFT, fill=ctk.BOTH, expand=True, padx=5)
        scrollbar = Scrollbar(text_frame, command=self.input_text.yview)
        scrollbar.pack(side=ctk.LEFT, fill='y', padx=2)
        self.input_text.configure(yscrollcommand=scrollbar.set)
        
        action_frame = ctk.CTkFrame(main_frame)
        action_frame.pack(pady=20, fill=ctk.X)
        
        ctk.CTkButton(action_frame, text="Draw in OneNote",
                      command=self.draw_text_in_onenote,
                      fg_color='#7719AA', text_color='white',
                      font=('Arial', 14, 'bold'), width=180).pack(side=ctk.LEFT, padx=10)
        
        self.stop_button = ctk.CTkButton(action_frame, text="Stop!",
                                          command=self.stop_drawing_now,
                                          fg_color='#e74c3c', text_color='white',
                                          font=('Arial', 14, 'bold'), width=120, state='disabled')
        self.stop_button.pack(side=ctk.LEFT, padx=10)
        
        self.status_label = ctk.CTkLabel(main_frame, text="Ready!", font=('Arial', 12),
                                          text_color='#4ecdc4')
        self.status_label.pack(pady=10)
        
        options = ctk.CTkFrame(main_frame)
        options.pack(pady=15, fill=ctk.X)
        ctk.CTkButton(options, text="Relearn all characters",
                      command=self.reset_learning,
                      fg_color='#e74c3c', text_color='white',
                      font=('Arial', 11), width=170).pack(side=ctk.LEFT, padx=10)
        ctk.CTkButton(options, text="Change single character",
                      command=self.change_letter,
                      fg_color='#95a5a6', text_color='white',
                      font=('Arial', 11), width=170).pack(side=ctk.LEFT, padx=10)
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.load_text()

    # ------------------------------------------------------------------ sidebar toggle
    def _toggle_ai_sidebar(self):
        if self.ai_sidebar.is_open:
            self.ai_sidebar.close()
            self.ai_btn.configure(text="✨ AI Chat")
        else:
            self.ai_sidebar.root = self.root
            self.ai_sidebar._parent_frame = self.main_wrapper
            self.ai_sidebar.open()
            self.ai_btn.configure(text="✕ Close AI")

    # ------------------------------------------------------------------ settings
    def load_settings(self):
        default_settings = {
            'line_spacing': 60,
            'character_size': 0.1,
            'telemetry_opted_in': None
        }
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, 'r') as f:
                    self.settings = json.load(f)
                for key, value in default_settings.items():
                    if key not in self.settings:
                        self.settings[key] = value
            except (json.JSONDecodeError, IOError):
                self.settings = default_settings.copy()
        else:
            self.settings = default_settings.copy()
        self.line_spacing_value = self.settings['line_spacing']
        self.character_size_value = self.settings['character_size']
        self.telemetry_opted_in = self.settings['telemetry_opted_in']
    
    def save_settings(self):
        self.settings['line_spacing'] = self.line_spacing_value
        self.settings['character_size'] = self.character_size_value
        self.settings['telemetry_opted_in'] = self.telemetry_opted_in
        try:
            with open(self.settings_file, 'w') as f:
                json.dump(self.settings, f, indent=2)
        except IOError as e:
            print(f"Error saving settings: {e}")
    
    def show_telemetry_dialog(self):
        telemetry_win = Toplevel(self.root)
        telemetry_win.title("Welcome to Text2Pen")
        telemetry_win.resizable(False, False)
        telemetry_win.transient(self.root)
        telemetry_win.grab_set()
        
        ctk.CTkLabel(telemetry_win, text="Welcome to Text2Pen",
                     font=('Arial', 24, 'bold')).pack(pady=20)
        
        container_frame = ctk.CTkFrame(telemetry_win)
        container_frame.pack(fill=ctk.BOTH, expand=True, padx=40, pady=20)
        info_frame = ctk.CTkFrame(container_frame)
        info_frame.pack(fill=ctk.BOTH, expand=True)
        
        ctk.CTkLabel(info_frame,
                     text="Help us improve Text2Pen by sharing anonymous data.\n\n"
                          "We collect:\n"
                          "  • Crash reports and error logs\n"
                          "  • Basic usage statistics\n"
                          "  • Feature usage data\n\n"
                          "All data is completely anonymous.\n\n"
                          "Data may be processed outside the EU.",
                     font=('Arial', 13), justify='left', wraplength=500).pack(anchor='w', pady=15)
        
        telemetry_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(info_frame, text="I agree to share anonymous usage data",
                        variable=telemetry_var, font=('Arial', 12)).pack(anchor='w', pady=15)
        
        button_frame = ctk.CTkFrame(telemetry_win)
        button_frame.pack(pady=20)
        
        def confirm():
            self.telemetry_opted_in = telemetry_var.get()
            self.save_settings()
            telemetry_win.destroy()
        
        ctk.CTkButton(button_frame, text="Continue", command=confirm,
                      fg_color='#4ecdc4', text_color='white',
                      font=('Arial', 14), width=180).pack()
        
        telemetry_win.update_idletasks()
        x = (telemetry_win.winfo_screenwidth() // 2) - (telemetry_win.winfo_width() // 2)
        y = (telemetry_win.winfo_screenheight() // 2) - (telemetry_win.winfo_height() // 2)
        telemetry_win.geometry(f"+{x}+{y}")
    
    def open_settings(self):
        self.load_settings()
        settings_win = Toplevel(self.root)
        settings_win.title("Settings")
        settings_win.geometry("600x400")
        settings_win.resizable(False, False)
        settings_win.transient(self.root)
        settings_win.grab_set()
        
        ctk.CTkLabel(settings_win, text="Settings",
                     font=('Arial', 20, 'bold')).pack(pady=20)
        
        settings_frame = ctk.CTkScrollableFrame(settings_win)
        settings_frame.pack(fill=ctk.BOTH, expand=True, padx=20, pady=20)
        
        row1 = ctk.CTkFrame(settings_frame)
        row1.pack(fill=ctk.X, pady=15)
        ctk.CTkLabel(row1, text="Character Size:", font=('Arial', 12)).pack(side=ctk.LEFT, expand=True)
        self.characterSize = Spinbox(row1, from_=0.1, to=1.0, width=10, font=('Arial', 12), increment=0.1)
        self.characterSize.delete(0, 'end')
        self.characterSize.insert(0, str(self.character_size_value))
        self.characterSize.pack(side=ctk.RIGHT, padx=10)
        
        row2 = ctk.CTkFrame(settings_frame)
        row2.pack(fill=ctk.X, pady=15)
        ctk.CTkLabel(row2, text="Line Spacing:", font=('Arial', 12)).pack(side=ctk.LEFT, expand=True)
        self.line_spacing = Spinbox(row2, from_=10, to=100, width=10, font=('Arial', 12))
        self.line_spacing.delete(0, 'end')
        self.line_spacing.insert(0, str(self.line_spacing_value))
        self.line_spacing.pack(side=ctk.RIGHT, padx=10)
        
        row3 = ctk.CTkFrame(settings_frame)
        row3.pack(fill=ctk.X, pady=15)
        ctk.CTkLabel(row3, text="Anonymous Telemetry:", font=('Arial', 12)).pack(side=ctk.LEFT, expand=True)
        telemetry_var = ctk.BooleanVar(value=bool(self.telemetry_opted_in))
        ctk.CTkCheckBox(row3, text="Enable crash & error reports",
                        variable=telemetry_var, font=('Arial', 12)).pack(side=ctk.RIGHT, padx=10)

        # Limits display in settings too
        row4 = ctk.CTkFrame(settings_frame)
        row4.pack(fill=ctk.X, pady=15)
        ctk.CTkLabel(row4, text="API Limits (today):", font=('Arial', 12, 'bold')).pack(anchor='w', pady=4)
        self.settings_limit_label = ctk.CTkLabel(row4, text="Loading...", font=('Arial', 11), text_color="gray")
        self.settings_limit_label.pack(anchor='w')

        def fetch_limits_for_settings():
            try:
                resp = requests.get(AI_LIMITS_URL, timeout=10)
                if resp.status_code == 200:
                    d = resp.json()
                    txt = (f"💬 AI Chat: {d.get('chat_used',0)}/{d.get('chat_limit',10)} requests\n"
                           f"📝 Text Actions: {d.get('text_used',0)}/{d.get('text_limit',1000)} requests")
                    settings_win.after(0, lambda: self.settings_limit_label.configure(text=txt))
            except Exception:
                settings_win.after(0, lambda: self.settings_limit_label.configure(text="Not available"))
        Thread(target=fetch_limits_for_settings, daemon=True).start()
        
        button_frame = ctk.CTkFrame(settings_win)
        button_frame.pack(pady=20)
        
        def save_and_close():
            self.character_size_value = float(self.characterSize.get())
            self.line_spacing_value = int(self.line_spacing.get())
            self.telemetry_opted_in = bool(telemetry_var.get())
            self.save_settings()
            settings_win.destroy()
        
        settings_win.protocol("WM_DELETE_WINDOW", save_and_close)
        ctk.CTkButton(button_frame, text="Close", command=save_and_close,
                      fg_color='#4ecdc4', text_color='white',
                      font=('Arial', 12), width=120).pack()

    # ------------------------------------------------------------------ template
    def draw_letter_template(self, letter):
        self.canvas.delete('template')
        self.canvas.create_text(300, 200, text=letter,
                                font=("Arial", 260, "bold"), fill="#d0d0d0", tags="template")

    # ------------------------------------------------------------------ drawing
    def start_drawing(self, event):
        self.current_stroke = [(event.x, event.y)]
    
    def draw(self, event):
        if self.current_stroke:
            x1, y1 = self.current_stroke[-1]
            self.canvas.create_line(x1, y1, event.x, event.y, width=3, fill='black',
                                    capstyle=ROUND, smooth=True)
        self.current_stroke.append((event.x, event.y))
    
    def finish_stroke(self, event):
        if self.current_stroke:
            self.strokes.append(self.current_stroke.copy())
            self.current_stroke = []

    def delete_learning(self):
        self.canvas.delete('all')
        self.draw_letter_template(self.current_letter)
        self.strokes = []

    # ------------------------------------------------------------------ save letter
    def save_letter(self):
        if not self.strokes:
            return
        self.letter_db[self.current_letter] = self.strokes.copy()
        self.save_to_file()
        self.next_letter()
    
    def next_letter(self):
        idx = self.alphabet.index(self.current_letter)
        if idx < len(self.alphabet) - 1:
            self.current_letter = self.alphabet[idx + 1]
            self.canvas.delete('all')
            self.strokes = []
            self.draw_letter_template(self.current_letter)
            self.title_label.configure(text=f"Learning character: {self.current_letter}")
            self.progress_label.configure(
                text=f"Progress: {len(self.letter_db)}/{len(self.alphabet)} characters learned!")
        else:
            self.learning_mode = False
            for w in self.root.winfo_children():
                w.destroy()
            self.create_text_gui()

    # ------------------------------------------------------------------ failsafe
    def stop_drawing_now(self):
        self.stop_drawing = True
        self.status_label.configure(text="⏹ Drawing stopped!")
    
    def failsafe(self):
        x, y = win32api.GetCursorPos()
        if x < 10 and y < 10:
            self.stop_drawing = True
            return True
        return False

    # ------------------------------------------------------------------ onenote drawing
    def draw_text_in_onenote(self):
        text = self.input_text.get('1.0', 'end-1c')
        if not text:
            self.status_label.configure(text="Please enter some text!")
            return

        blocks = parse_text_with_tables(text)
        
        for block in blocks:
            to_check = block['content'] if block['type'] == 'text' else ''.join(''.join(row) for row in block['rows'])
            for ch in to_check:
                if ch and ch not in (' ', '\n', '\t') and ch not in self.letter_db:
                    self.status_label.configure(text=f"Character '{ch}' not learned!")
                    return
        
        self.stop_drawing = False
        self.stop_button.configure(state='normal')
        self.status_label.configure(text="Please open OneNote, starting in 4 seconds...")

        self.total_chars = sum(
            len([c for c in block['content'] if c not in (' ', '\n')])
            if block['type'] == 'text'
            else len([c for row in block['rows'] for c in ''.join(row) if c != ' '])
            for block in blocks
        )
        self.drawn_chars = 0
        Thread(target=self.onenote_thread, args=(blocks,)).start()

    def get_letter_bounds(self, ch):
        strokes = self.letter_db[ch]
        min_x = min_y = float('inf')
        max_x = max_y = float('-inf')
        for stroke in strokes:
            for x, y in stroke:
                if x < min_x: min_x = x
                if x > max_x: max_x = x
                if y < min_y: min_y = y
                if y > max_y: max_y = y
        return min_x, min_y, max_x, max_y

    def get_letter_width(self, ch):
        min_x, _, max_x, _ = self.get_letter_bounds(ch)
        return max_x - min_x

    def get_letter_height(self, ch):
        _, min_y, _, max_y = self.get_letter_bounds(ch)
        return max_y - min_y
    
    def onenote_thread(self, blocks):
        time.sleep(4)
        self.root.after(0, self.show_progress_overlay)
        
        if self.stop_drawing:
            self.root.after(0, lambda: self.stop_button.configure(state='disabled'))
            return
        
        hwnd = self.find_onenote_window()
        if not hwnd:
            self.root.after(0, lambda: self.status_label.configure(text="Error: OneNote not found!"))
            self.root.after(0, lambda: self.stop_button.configure(state='disabled'))
            self.root.after(0, self.close_progress_overlay)
            return
        
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.4)
        
        pt = win32gui.ClientToScreen(hwnd, (0, 0))
        client_left, client_top = pt

        self._draw_context = {
            "canvas_x": client_left,
            "canvas_y": client_top + 90,
            "line_spacing_px": self.line_spacing_value,
            "scale": self.character_size_value
        }

        offset_x, offset_y, current_lines = 50, 50, 0

        for block in blocks:
            if self.stop_drawing or self.failsafe():
                break
            if block['type'] == 'text':
                offset_x, offset_y, current_lines = self.draw_text_block(
                    block['content'], offset_x, offset_y, current_lines)
            elif block['type'] == 'table':
                offset_x, offset_y, current_lines = self.draw_table(
                    block['rows'], offset_x, offset_y, current_lines)
                offset_y += self.line_spacing_value
                current_lines += 1
        
        msg = "⏹ Stopped!" if self.stop_drawing else "✅ Done!"
        self.root.after(0, lambda: self.status_label.configure(text=msg))
        self.root.after(0, lambda: self.stop_button.configure(state='disabled'))
        self.root.after(0, self.close_progress_overlay)

    def _scroll_if_needed(self, current_lines):
        if current_lines >= 5:
            win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, -1000)
            time.sleep(1)
            return 0, 50
        return current_lines, None

    def draw_character(self, ch, offset_x, offset_y, scale, jitter=True):
        if ch not in self.letter_db:
            return 0
        
        self.drawn_chars += 1
        p = self.drawn_chars / self.total_chars
        self.root.after(0, lambda v=p: self.overlay_bar.set(v))

        raw_width = self.get_letter_width(ch)
        letter_scale = scale * random.uniform(0.85, 1.0) if jitter else scale
        effective_width = max(raw_width, 60)
        letter_spacing = max(
            int(effective_width * letter_scale * 1.23) + 8,
            int(50 * letter_scale) + 8
        )

        canvas_x = self._draw_context['canvas_x']
        canvas_y = self._draw_context['canvas_y']

        strokes = self.letter_db[ch]
        self.root.after(0, lambda c=ch: self.status_label.configure(text=f"Drawing '{c}'..."))
        offset_letter_y = random.randint(-8, 8) if jitter else 0

        for stroke in strokes:
            if self.stop_drawing or self.failsafe() or len(stroke) < 2:
                continue

            start_x, start_y = stroke[0]
            start_y += offset_letter_y
            sx = canvas_x + int(start_x * letter_scale) + int(offset_x)
            sy = canvas_y + int(start_y * letter_scale) + int(offset_y)

            win32api.SetCursorPos((int(sx), int(sy)))
            time.sleep(0.003)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
            time.sleep(0.0003)

            last_x, last_y = sx, sy
            for x, y in stroke[1::3]:
                if self.stop_drawing or self.failsafe():
                    break
                y_off = random.randint(-2, 2) if jitter else 0
                nx = canvas_x + int(x * letter_scale) + int(offset_x)
                ny = canvas_y + int((y + y_off) * letter_scale) + int(offset_y)
                if jitter:
                    nx += random.randint(-1, 1)
                    ny += random.randint(-1, 1)
                win32api.mouse_event(win32con.MOUSEEVENTF_MOVE, int(nx - last_x), int(ny - last_y))
                last_x, last_y = nx, ny

            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
            time.sleep(0.0003)

        return letter_spacing

    def draw_text_block(self, text, offset_x, offset_y, current_lines):
        scale = self._draw_context['scale']
        line_spacing_px = self._draw_context['line_spacing_px']
        chars_per_line = int(6 / scale)
        base_char_spacing_px = int(180 * scale)

        for line in text.split('\n'):
            if self.stop_drawing or self.failsafe():
                break
            offset_x = 50
            chars_in_line = 0

            for ch in line:
                if self.stop_drawing or self.failsafe():
                    break
                if chars_in_line >= chars_per_line and ch == ' ':
                    current_lines += 1
                    current_lines, new_y = self._scroll_if_needed(current_lines)
                    if new_y is not None:
                        offset_y = new_y
                    else:
                        offset_y += line_spacing_px
                    offset_x = 50
                    chars_in_line = 0
                    time.sleep(0.7)
                    continue
                if ch == ' ':
                    offset_x += base_char_spacing_px
                    chars_in_line += 1
                    continue
                spacing = self.draw_character(ch, offset_x, offset_y, scale, jitter=True)
                offset_x += spacing
                chars_in_line += 1

            offset_y += line_spacing_px
            current_lines += 1
            current_lines, new_y = self._scroll_if_needed(current_lines)
            if new_y is not None:
                offset_y = new_y

        return offset_x, offset_y, current_lines

    def draw_line_abs(self, screen_x1, screen_y1, screen_x2, screen_y2):
        steps = max(abs(screen_x2 - screen_x1), abs(screen_y2 - screen_y1), 1)
        coarse_steps = max(steps // 3, 1)
        win32api.SetCursorPos((int(screen_x1), int(screen_y1)))
        time.sleep(0.01)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
        time.sleep(0.005)
        last_x, last_y = screen_x1, screen_y1
        for i in range(1, coarse_steps + 1):
            nx = screen_x1 + int((screen_x2 - screen_x1) * i / coarse_steps)
            ny = screen_y1 + int((screen_y2 - screen_y1) * i / coarse_steps)
            if i < coarse_steps:
                if abs(screen_x2 - screen_x1) > abs(screen_y2 - screen_y1):
                    ny += random.randint(-2, 2)
                else:
                    nx += random.randint(-2, 2)
            win32api.mouse_event(win32con.MOUSEEVENTF_MOVE, int(nx - last_x), int(ny - last_y))
            last_x, last_y = nx, ny
            time.sleep(0.002)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
        time.sleep(0.005)

    def _measure_col_widths(self, padded_rows, scale):
        col_count = len(padded_rows[0])
        col_widths = []
        for col_idx in range(col_count):
            max_cell_width = 0
            for row in padded_rows:
                cell = str(row[col_idx])
                cell_px = 0
                for ch in cell:
                    if ch == ' ':
                        cell_px += int(180 * scale)
                    elif ch in self.letter_db:
                        raw_w = self.get_letter_width(ch)
                        effective_w = max(raw_w, 60)
                        cell_px += max(int(effective_w * scale * 1.23) + 8, int(50 * scale) + 8)
                if cell_px > max_cell_width:
                    max_cell_width = cell_px
            col_widths.append(max(max_cell_width + 40, 80))
        return col_widths

    def _measure_row_height(self, scale):
        heights = [self.get_letter_height(ch) for ch in self.letter_db if self.get_letter_height(ch) > 0]
        avg_letter_h = (sum(heights) / len(heights)) if heights else 200
        return max(int(avg_letter_h * scale) + 40, 60)

    def draw_table(self, table, offset_x, offset_y, current_lines):
        if not table:
            return offset_x, offset_y, current_lines
        scale = self._draw_context['scale']
        canvas_x = self._draw_context['canvas_x']
        canvas_y = self._draw_context['canvas_y']
        col_count = max(len(row) for row in table)
        padded_rows = [row + [''] * (col_count - len(row)) for row in table]
        col_widths = self._measure_col_widths(padded_rows, scale)
        row_height = self._measure_row_height(scale)
        table_width = sum(col_widths)
        origin_sx = canvas_x + int(offset_x)
        origin_sy = canvas_y + int(offset_y)

        for row_idx, row in enumerate(padded_rows):
            if self.stop_drawing or self.failsafe():
                break
            row_top_sy    = origin_sy + row_idx * row_height
            row_bottom_sy = row_top_sy + row_height
            if row_idx == 0:
                self.draw_line_abs(origin_sx, row_top_sy, origin_sx + table_width, row_top_sy)
            self.draw_line_abs(origin_sx, row_bottom_sy, origin_sx + table_width, row_bottom_sy)
            col_sx = origin_sx
            for col_idx, cell in enumerate(row):
                if self.stop_drawing or self.failsafe():
                    break
                col_width = col_widths[col_idx]
                self.draw_line_abs(col_sx, row_top_sy, col_sx, row_bottom_sy)
                char_offset_x = col_sx + 15 - canvas_x
                char_offset_y = row_top_sy + 12 - canvas_y
                for ch in str(cell):
                    if self.stop_drawing or self.failsafe():
                        break
                    if ch == ' ':
                        char_offset_x += int(180 * scale)
                    else:
                        spacing = self.draw_character(ch, char_offset_x, char_offset_y, scale, jitter=False)
                        char_offset_x += spacing
                col_sx += col_width
            self.draw_line_abs(origin_sx + table_width, row_top_sy, origin_sx + table_width, row_bottom_sy)
            current_lines += 1

        return 50, offset_y + len(padded_rows) * row_height, current_lines

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

    def save_to_file(self):
        with open(self.db_file, 'w') as f:
            json.dump(self.letter_db, f)
    
    def load_letters(self):
        if os.path.exists(self.db_file):
            with open(self.db_file, 'r') as f:
                self.letter_db = json.load(f)

    def reset_learning(self):
        self.letter_db = {}
        if os.path.exists(self.db_file):
            os.remove(self.db_file)
        self.current_letter = 'a'
        self.learning_mode = True
        for w in self.root.winfo_children():
            w.destroy()
        self.create_learning_gui()
    
    def change_letter(self):
        win = Toplevel(self.root)
        win.title("Change character")
        win.geometry("300x200")
        ctk.CTkLabel(win, text="Which character would you like to change?",
                     font=('Arial', 12)).pack(pady=20)
        var = StringVar(value='a')
        ttk.Combobox(win, textvariable=var, values=list(self.alphabet),
                     font=('Arial', 14)).pack(pady=10)
        def apply():
            self.current_letter = var.get()
            self.learning_mode = True
            win.destroy()
            for w in self.root.winfo_children():
                w.destroy()
            self.create_learning_gui()
        ctk.CTkButton(win, text="Change", command=apply,
                      fg_color='#4ecdc4', text_color='white',
                      font=('Arial', 12)).pack(pady=20)


if __name__ == "__main__":
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")
    root = ctk.CTk()
    root.title("Text2Pen")
    app = LetterApp(root)
    root.mainloop()
