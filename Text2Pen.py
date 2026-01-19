from tkinter import Canvas, Frame, Text, Scrollbar, Toplevel, StringVar, ROUND, Spinbox
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

BACKEND_URL = "https://text2pen-backend.onrender.com/telemetry"

INSTALL_DIR = os.path.join(os.environ["LOCALAPPDATA"], "Text2Pen")

UPDATE_EXE = os.path.join(INSTALL_DIR, "Update.exe")

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

#Update process
if os.path.exists(os.path.join(INSTALL_DIR, "Update-newest.exe")):
    os.replace(os.path.join(INSTALL_DIR, "Update-newest.exe"), os.path.join(INSTALL_DIR, "Update.exe"))

class LetterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Text2Pen")
        
        # Letter database
        self.letter_db = {}
        self.current_letter = 'a'
        self.learning_mode = True
        self.alphabet = "abcdefghijklmnopqrstuvwxyzäöüÄÖÜßABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.,!?;:-_—()[]{}<>\"'/*+=@#$%^&|~`"
        self.current_stroke = []
        self.stop_drawing = False
        
        # Storage files
        self.db_file = 'letter_db.json'
        self.settings_file = 'settingsDB.json'
        
        self.load_letters()
        self.load_settings()
        
        # Check if first launch (settings file didn't exist before)
        if not os.path.exists(self.settings_file) or 'telemetry_opted_in' not in self.settings:
            self.show_telemetry_dialog()

        self.root.report_callback_exception = self.tk_exception_handler
        
        # Default settings (store values, not widgets)
        self.line_spacing_value = 60
        self.character_size_value = 0.1
        
        # Spinbox widget references
        self.line_spacing = None
        self.characterSize = None
        
        # Switch to text mode if all letters exist
        if len(self.letter_db) == self.alphabet.__len__():
            self.learning_mode = False
        
        self.create_gui()

        self.root._state_before_windows_set_titlebar_color = 'zoomed'
        
    ## Telemetry functions
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
        
        Thread(
            target=self.send_error_event,
            args=(exception,),
            daemon=True
        ).start()
    
    def send_error_event(self, exception):
        if not self.telemetry_opted_in:
            return
        
        event = {
            "event": "ERROR",
            "os": platform.system(),
            "timestamp": time.time(),
            "extra": {"Exact Error": self.sanitize_exception(exception)}
        }

        try:
            resp = requests.post(BACKEND_URL, json=event, timeout=60)
            if resp.status_code == 200:
                return "sucess"
            else:
                print("Failed:", resp.status_code, resp.text)
                return "failed"
        except Exception as e:
            print("Error:", e)
            return "failed"
        
    
    def create_gui(self):
        if self.learning_mode:
            self.create_learning_gui()
        else:
            self.create_text_gui()
    
    def create_learning_gui(self):
        # Title
        self.title_label = ctk.CTkLabel(
            self.root,
            text=f"Learning character: {self.current_letter}",
            font=('Arial', 24, 'bold')
        )
        self.title_label.pack(pady=20)
        
        # Info
        ctk.CTkLabel(
            self.root,
            text="Draw the character shown in the background.",
            font=('Arial', 14)
        ).pack(pady=5)
        
        # Canvas
        self.canvas = Canvas(self.root, width=700, height=450, bg='white', highlightthickness=2, highlightbackground='#cccccc')
        self.canvas.pack(pady=15, padx=20, fill=ctk.BOTH, expand=True)
        
        # Draw template letter
        self.draw_letter_template(self.current_letter)
        
        # Mouse events
        self.canvas.bind('<Button-1>', self.start_drawing)
        self.canvas.bind('<B1-Motion>', self.draw)
        self.canvas.bind('<ButtonRelease-1>', self.finish_stroke)
        
        # Buttons
        button_frame = ctk.CTkFrame(self.root)
        button_frame.pack(pady=20)
        
        ctk.CTkButton(button_frame, text="Delete", command=self.delete_learning,
              fg_color='#ff6b6b', text_color='white', font=('Arial', 14), width=120).pack(side=ctk.LEFT, padx=10)
        
        ctk.CTkButton(button_frame, text="Save character", command=self.save_letter,
                  fg_color='#4ecdc4', text_color='white', font=('Arial', 14), width=150).pack(side=ctk.LEFT, padx=10)
        
        ctk.CTkButton(button_frame, text="Skip", command=self.next_letter,
                  fg_color='#95a5a6', text_color='white', font=('Arial', 14), width=120).pack(side=ctk.LEFT, padx=10)
        
        # Progress display
        self.progress_label = ctk.CTkLabel(
            self.root,
            text=f"Progression: {len(self.letter_db)}/{self.alphabet.__len__()} characters learned!",
            font=('Arial', 12)
        )
        self.progress_label.pack(pady=10)
        
        self.strokes = []
    
    def create_text_gui(self):
        # Top header frame with title and settings button
        header_frame = ctk.CTkFrame(self.root)
        header_frame.pack(fill=ctk.X, padx=20, pady=15)
        
        # Title
        ctk.CTkLabel(
            header_frame,
            text="Text2Pen - Write text in OneNote",
            font=('Arial', 26, 'bold')
        ).pack(side=ctk.LEFT, expand=True)
        
        # Settings button (gear icon)
        ctk.CTkButton(header_frame, text="⚙️", command=self.open_settings,
                  fg_color='#95a5a6', text_color='white', font=('Arial', 18), width=50).pack(side=ctk.RIGHT, padx=10)
        
        # Main scrollable frame
        main_frame = ctk.CTkScrollableFrame(self.root)
        main_frame.pack(fill=ctk.BOTH, expand=True, padx=20, pady=10)
        
        # Info frame
        info_frame = ctk.CTkFrame(main_frame)
        info_frame.pack(pady=15, fill=ctk.X)
        
        ctk.CTkLabel(info_frame, text="Input text below and press 'Draw in OneNote'!",
                 font=('Arial', 13)).pack(anchor='w')
        ctk.CTkLabel(info_frame, text="🛑 FAILSAFE: Move mouse to the top left corner to stop!",
                 font=('Arial', 11), text_color='red').pack(anchor='w', pady=5)
        
        # Text input frame
        text_label = ctk.CTkLabel(main_frame, text="Text:", font=('Arial', 13, 'bold'))
        text_label.pack(anchor='w', pady=(15, 5))
        
        text_frame = ctk.CTkFrame(main_frame)
        text_frame.pack(fill=ctk.BOTH, expand=True, pady=10)
        
        self.input_text = Text(text_frame, font=('Arial', 12), wrap='word', height=15)
        self.input_text.pack(side=ctk.LEFT, fill=ctk.BOTH, expand=True, padx=5)
        
        scrollbar = Scrollbar(text_frame, command=self.input_text.yview)
        scrollbar.pack(side=ctk.LEFT, fill='y', padx=2)
        self.input_text.configure(yscrollcommand=scrollbar.set)
        
        # Action buttons frame
        action_frame = ctk.CTkFrame(main_frame)
        action_frame.pack(pady=20, fill=ctk.X)
        
        ctk.CTkButton(action_frame, text="Draw in OneNote", command=self.draw_text_in_onenote,
                  fg_color='#7719AA', text_color='white', font=('Arial', 14, 'bold'), width=180).pack(side=ctk.LEFT, padx=10)
        
        self.stop_button = ctk.CTkButton(action_frame, text="Stop!", command=self.stop_drawing_now,
                  fg_color='#e74c3c', text_color='white', font=('Arial', 14, 'bold'), width=120,
                  state='disabled')
        self.stop_button.pack(side=ctk.LEFT, padx=10)
        
        # Status
        self.status_label = ctk.CTkLabel(main_frame, text="Ready to draw!", font=('Arial', 12), text_color='#4ecdc4')
        self.status_label.pack(pady=10)
        
        # Options frame
        options = ctk.CTkFrame(main_frame)
        options.pack(pady=15, fill=ctk.X)
        
        ctk.CTkButton(options, text="Relearn all characters", command=self.reset_learning,
                  fg_color='#e74c3c', text_color='white', font=('Arial', 11), width=160).pack(side=ctk.LEFT, padx=10)
        
        ctk.CTkButton(options, text="Change single character", command=self.change_letter,
                  fg_color='#95a5a6', text_color='white', font=('Arial', 11), width=160).pack(side=ctk.LEFT, padx=10)
    
    def load_settings(self):
        """Load settings from database, create with defaults if not exists"""
        default_settings = {
            'line_spacing': 60,
            'character_size': 0.1,
            'telemetry_opted_in': None
        }
        
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, 'r') as f:
                    self.settings = json.load(f)
                    # Merge with defaults for any missing keys
                    for key, value in default_settings.items():
                        if key not in self.settings:
                            self.settings[key] = value
            except (json.JSONDecodeError, IOError):
                self.settings = default_settings.copy()
        else:
            self.settings = default_settings.copy()
        
        # Store values as instance variables for easy access
        self.line_spacing_value = self.settings['line_spacing']
        self.character_size_value = self.settings['character_size']
        self.telemetry_opted_in = self.settings['telemetry_opted_in']
    
    def save_settings(self):
        """Save current settings to database"""
        self.settings['line_spacing'] = self.line_spacing_value
        self.settings['character_size'] = self.character_size_value
        self.settings['telemetry_opted_in'] = self.telemetry_opted_in
        
        try:
            with open(self.settings_file, 'w') as f:
                json.dump(self.settings, f, indent=2)
        except IOError as e:
            print(f"Error saving settings: {e}")
    
    def show_telemetry_dialog(self):
        """Show telemetry opt-in dialog on first launch"""
        telemetry_win = Toplevel(self.root)
        telemetry_win.title("Welcome to Text2Pen")
        telemetry_win.resizable(False, False)
        telemetry_win.transient(self.root)
        telemetry_win.grab_set()
        
        # Title
        ctk.CTkLabel(telemetry_win, text="Welcome to Text2Pen", font=('Arial', 24, 'bold')).pack(pady=20)
        
        # Main container frame
        container_frame = ctk.CTkFrame(telemetry_win)
        container_frame.pack(fill=ctk.BOTH, expand=True, padx=40, pady=20)
        
        # Info frame
        info_frame = ctk.CTkFrame(container_frame)
        info_frame.pack(fill=ctk.BOTH, expand=True)
        
        # Description text with proper wrapping
        description = ctk.CTkLabel(info_frame, 
            text="Help us improve Text2Pen by sharing anonymous data.\n\n"
                 "We collect:\n"
                 "  • Crash reports and error logs\n"
                 "  • Basic usage statistics\n"
                 "  • Feature usage data\n\n"
                 "All data is completely anonymous and helps us\n"
                 "make Text2Pen better for you.\n\n"
                 "Data could be processed outside of the EU.",
            font=('Arial', 13), justify='left', wraplength=500)
        description.pack(anchor='w', pady=15)
        
        # Checkbox for telemetry
        telemetry_var = ctk.BooleanVar(value=False)
        checkbox = ctk.CTkCheckBox(info_frame, text="I agree to share anonymous usage data",
                                   variable=telemetry_var, font=('Arial', 12))
        checkbox.pack(anchor='w', pady=15)
        
        # Button frame
        button_frame = ctk.CTkFrame(telemetry_win)
        button_frame.pack(pady=20)
        
        def confirm():
            self.telemetry_opted_in = telemetry_var.get()
            self.save_settings()
            telemetry_win.destroy()
        
        ctk.CTkButton(button_frame, text="Continue", command=confirm,
                  fg_color='#4ecdc4', text_color='white', font=('Arial', 14), width=180).pack()
        
        telemetry_win.update_idletasks()
        
        x = (telemetry_win.winfo_screenwidth() // 2) - (telemetry_win.winfo_width() // 2)
        y = (telemetry_win.winfo_screenheight() // 2) - (telemetry_win.winfo_height() // 2)
        telemetry_win.geometry(f"+{x}+{y}")
    
    def open_settings(self):
        # Reload settings from disk to get current state
        self.load_settings()
        
        settings_win = Toplevel(self.root)
        settings_win.title("Settings")
        settings_win.geometry("600x350")
        settings_win.resizable(False, False)
        
        # Make it modal
        settings_win.transient(self.root)
        settings_win.grab_set()
        
        # Title
        ctk.CTkLabel(settings_win, text="Settings", font=('Arial', 20, 'bold')).pack(pady=20)
        
        # Settings container frame
        settings_frame = ctk.CTkScrollableFrame(settings_win)
        settings_frame.pack(fill=ctk.BOTH, expand=True, padx=20, pady=20)

        
        # Character Size
        row1 = ctk.CTkFrame(settings_frame)
        row1.pack(fill=ctk.X, pady=15)
        ctk.CTkLabel(row1, text="Character Size:", font=('Arial', 12)).pack(side=ctk.LEFT, expand=True)
        self.characterSize = Spinbox(row1, from_=0.1, to=1.0, width=10, font=('Arial', 12), increment=0.1)
        self.characterSize.delete(0, 'end')
        self.characterSize.insert(0, str(self.character_size_value))
        self.characterSize.pack(side=ctk.RIGHT, padx=10)
        
        # Line spacing
        row2 = ctk.CTkFrame(settings_frame)
        row2.pack(fill=ctk.X, pady=15)
        ctk.CTkLabel(row2, text="Line Spacing:", font=('Arial', 12)).pack(side=ctk.LEFT, expand=True)
        self.line_spacing = Spinbox(row2, from_=10, to=100, width=10, font=('Arial', 12))
        self.line_spacing.delete(0, 'end')
        self.line_spacing.insert(0, str(self.line_spacing_value))
        self.line_spacing.pack(side=ctk.RIGHT, padx=10)

        # Telemetry toggle
        row3 = ctk.CTkFrame(settings_frame)
        row3.pack(fill=ctk.X, pady=15)

        ctk.CTkLabel(
            row3,
            text="Anonymous Telemetry:",
            font=('Arial', 12)
        ).pack(side=ctk.LEFT, expand=True)

        telemetry_var = ctk.BooleanVar(
            value=bool(self.telemetry_opted_in)
        )

        telemetry_checkbox = ctk.CTkCheckBox(
            row3,
            text="Enable crash & error reports",
            variable=telemetry_var,
            font=('Arial', 12)
        )
        telemetry_checkbox.pack(side=ctk.RIGHT, padx=10)

        
        # Close button
        button_frame = ctk.CTkFrame(settings_win)
        button_frame.pack(pady=20)
        
        def save_settings_and_close():
            self.character_size_value = float(self.characterSize.get())
            self.line_spacing_value = int(self.line_spacing.get())

            self.telemetry_opted_in = bool(telemetry_var.get())

            self.save_settings()
            settings_win.destroy()
        
        settings_win.protocol("WM_DELETE_WINDOW", save_settings_and_close)

        ctk.CTkButton(button_frame, text="Close", command=save_settings_and_close,
                  fg_color='#4ecdc4', text_color='white', font=('Arial', 12), width=120).pack()

    # ---------------------------------------------------------
    # TEMPLATE AS BIG GRAY TEXT LETTER
    # ---------------------------------------------------------
    def draw_letter_template(self, letter):
        self.canvas.delete('template')
        self.canvas.create_text(
            300, 200,
            text=letter,
            font=("Arial", 260, "bold"),
            fill="#d0d0d0",
            tags="template"
        )

    # ---------------------------------------------------------
    # DRAWING SYSTEM
    # ---------------------------------------------------------
    def start_drawing(self, event):
        self.current_stroke = [(event.x, event.y)]
    
    def draw(self, event):
        if self.current_stroke:
            x1, y1 = self.current_stroke[-1]
            x2, y2 = event.x, event.y
            self.canvas.create_line(x1, y1, x2, y2, width=3, fill='black',
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

    # ---------------------------------------------------------
    # SAVE LETTER
    # ---------------------------------------------------------
    def save_letter(self):
        if not self.strokes:
            return
        self.letter_db[self.current_letter] = self.strokes.copy()
        self.save_to_file()
        self.next_letter()
    
    def next_letter(self):
        idx = self.alphabet.index(self.current_letter)
        if idx < self.alphabet.__len__() - 1:
            self.current_letter = self.alphabet[idx + 1]
            self.canvas.delete('all')
            self.strokes = []
            self.draw_letter_template(self.current_letter)
            self.title_label.configure(text=f"Learning character: {self.current_letter}")
            self.progress_label.configure(text=f"Progression: {len(self.letter_db)}/{self.alphabet.__len__()} character learned!")
        else:
            self.learning_mode = False
            for w in self.root.winfo_children():
                w.destroy()
            self.create_text_gui()
    
    # ---------------------------------------------------------
    # FAILSAFE + STOP
    # ---------------------------------------------------------
    def stop_drawing_now(self):
        self.stop_drawing = True
        self.status_label.configure(text="⏹ stopped drawing!")
    
    def failsafe(self):
        x, y = win32api.GetCursorPos()
        if x < 10 and y < 10:
            self.stop_drawing = True
            return True
        return False

    # ---------------------------------------------------------
    # DRAW INTO ONENOTE
    # ---------------------------------------------------------
    def draw_text_in_onenote(self):
        text = self.input_text.get('1.0', 'end-1c')
        if not text:
            self.status_label.configure(text="Please input Text!")
            return
        
        # Check trained letters
        for ch in text:
            if ch and ch != " " and ch != '\n'and ch not in self.letter_db:
                self.status_label.configure(text=f"Character '{ch}' not learned!")
                return
        
        self.stop_drawing = False
        self.stop_button.configure(state='normal')
        self.status_label.configure(text="Please open OneNote now, starting in 4 seconds...")
        
        Thread(target=self.onenote_thread, args=(text,)).start()

    def get_letter_width(self, ch):
        strokes = self.letter_db[ch]
        min_x = float('inf')
        max_x = float('-inf')

        for stroke in strokes:
            for x, y in stroke:
                if x < min_x:
                    min_x = x
                if x > max_x:
                    max_x = x

        return (max_x - min_x)
    
    def onenote_thread(self, text):
        time.sleep(4)
        
        if self.stop_drawing:
            self.root.after(0, lambda: self.stop_button.configure(state='disabled'))
            return
        
        hwnd = self.find_onenote_window()
        if not hwnd:
            self.root.after(0, lambda: self.status_label.configure(text="Error: OneNote not found!"))
            self.root.after(0, lambda: self.stop_button.configure(state='disabled'))
            return
        
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.4)
        
        client = win32gui.GetClientRect(hwnd)
        pt = win32gui.ClientToScreen(hwnd, (0,0))

        client_left, client_top = pt
        client_right = pt[0] + client[2]
        client_bottom = pt[1] + client[3]

        canvas_x = client_left
        canvas_y = client_top + 90  # Offset for toolbar

        line_spacing_px = self.line_spacing_value
        scale = self.character_size_value

        chars_per_line = int(6 / scale)
        
        base_char_spacing_px = int(180 * scale)
        offset_x = 50
        offset_y = 50
        
        current_lines = 0
        
        lines = text.split('\n')
        
        for line in lines:
            if self.stop_drawing or self.failsafe():
                break
            
            offset_x = 50
            chars_in_line = 0
            
            for ch in line:
                if self.stop_drawing or self.failsafe():
                    break
                
                if chars_in_line >= chars_per_line and ch == ' ':
                    current_lines += 1
                    offset_y += line_spacing_px
                    offset_x = 50
                    chars_in_line = 0
                    time.sleep(0.7)
                    continue

                if ch == ' ':
                    offset_x += base_char_spacing_px
                    chars_in_line += 1
                    continue
                
                if current_lines >= 5:
                    win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, -1000)
                    current_lines = 0
                    offset_y = 50
                    time.sleep(1)
                
                if ch not in self.letter_db:
                    continue
                
                raw_width = self.get_letter_width(ch)

                MIN_WIDTH = 60

                effective_width = max(raw_width, MIN_WIDTH)

                letter_spacing = max(
                    int(effective_width * scale * 1.23) + 8,
                    int(50 * scale) + 8
                )

                strokes = self.letter_db[ch]
                self.root.after(0, lambda c=ch: self.status_label.configure(text=f"Drawing '{c}'..."))
                
                offset_letter_y = random.randint(-8, 8)

                for stroke in strokes:
                    if self.stop_drawing or self.failsafe():
                        break
                    
                    if len(stroke) < 2:
                        continue

                    start_x, start_y = stroke[0]
                    start_y += offset_letter_y
                    sx = canvas_x + int(start_x * scale) + offset_x
                    sy = canvas_y + int(start_y * scale) + offset_y
                    
                    win32api.SetCursorPos((int(sx), int(sy)))
                    time.sleep(0.003)
                    
                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
                    time.sleep(0.0003)
                    
                    last_x, last_y = sx, sy
                    
                    for x, y in stroke[1:]:
                        if self.stop_drawing or self.failsafe():
                            break
                        
                        nx = canvas_x + int(x * scale) + offset_x
                        ny = canvas_y + int(y * scale) + offset_y
                        
                        dx = nx - last_x
                        dy = ny - last_y
                        
                        win32api.mouse_event(win32con.MOUSEEVENTF_MOVE, int(dx), int(dy))
                        
                        last_x, last_y = nx, ny
                    
                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
                    time.sleep(0.0003)
                
                offset_x += letter_spacing

                chars_in_line += 1
            
            offset_y += line_spacing_px
            current_lines += 1
        
        if self.stop_drawing:
            self.root.after(0, lambda: self.status_label.configure(text="⏹ stopped drawing!"))
        else:
            self.root.after(0, lambda: self.status_label.configure(text="✅ Finished!"))
        
        self.root.after(0, lambda: self.stop_button.configure(state='disabled'))


    # ---------------------------------------------------------
    # WINDOW FINDER
    # ---------------------------------------------------------
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

    # ---------------------------------------------------------
    # SAVE + LOAD
    # ---------------------------------------------------------
    def save_to_file(self):
        with open(self.db_file, 'w') as f:
            json.dump(self.letter_db, f)
    
    def load_letters(self):
        if os.path.exists(self.db_file):
            with open(self.db_file, 'r') as f:
                self.letter_db = json.load(f)

    # ---------------------------------------------------------
    # RESET + CHANGE
    # ---------------------------------------------------------
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
        
        ctk.CTkLabel(win, text="What character would you like to change?", font=('Arial', 12)).pack(pady=20)
        
        var = StringVar(value='a')
        combo = ttk.Combobox(win, textvariable=var, values=list(self.alphabet), font=('Arial', 14))
        combo.pack(pady=10)
        
        def apply():
            self.current_letter = var.get()
            self.learning_mode = True
            win.destroy()
            for w in self.root.winfo_children():
                w.destroy()
            self.create_learning_gui()
        
        ctk.CTkButton(win, text="Change", command=apply,
                 fg_color='#4ecdc4', text_color='white', font=('Arial', 12)).pack(pady=20)


if __name__ == "__main__":
    ctk.set_appearance_mode("System")  # "Dark", "Light", "System"
    ctk.set_default_color_theme("blue") # optional
    root = ctk.CTk()
    root.title("Text2Pen")
    
    app = LetterApp(root)

    root.mainloop()

