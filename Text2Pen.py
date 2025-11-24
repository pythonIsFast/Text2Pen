import tkinter as tk
from tkinter import ttk
import subprocess
import time
from threading import Thread
import win32gui
import win32con
import win32api
import json
import os

class LetterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Buchstaben lernen und in OneNote schreiben")
        self.root.geometry("700x600")
        
        # Letter database
        self.letter_db = {}
        self.current_letter = 'a'
        self.learning_mode = True
        self.alphabet = "abcdefghijklmnopqrstuvwxyzäöüÄÖÜßABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.,!?;:-_—()[]{}<>\"'/*+=@#$%^&|~`"
        self.current_stroke = []
        self.stop_drawing = False
        
        # Storage file
        self.db_file = 'letter_db.json'
        self.load_letters()
        
        # Switch to text mode if all letters exist
        if len(self.letter_db) == self.alphabet.__len__():
            self.learning_mode = False
        
        self.create_gui()
    
    def create_gui(self):
        if self.learning_mode:
            self.create_learning_gui()
        else:
            self.create_text_gui()
    
    def create_learning_gui(self):
        # Title
        self.title_label = tk.Label(
            self.root,
            text=f"Lerne Buchstabe: {self.current_letter}",
            font=('Arial', 20, 'bold')
        )
        self.title_label.pack(pady=10)
        
        # Info
        tk.Label(
            self.root,
            text="Zeichne den Buchstaben über die graue Vorlage",
            font=('Arial', 12)
        ).pack()
        
        # Canvas
        self.canvas = tk.Canvas(self.root, width=600, height=400, bg='white')
        self.canvas.pack(pady=10)
        
        # Draw template letter
        self.draw_letter_template(self.current_letter)
        
        # Mouse events
        self.canvas.bind('<Button-1>', self.start_drawing)
        self.canvas.bind('<B1-Motion>', self.draw)
        self.canvas.bind('<ButtonRelease-1>', self.finish_stroke)
        
        # Buttons
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=10)
        
        tk.Button(button_frame, text="Löschen", command=self.delete_learning,
                  bg='#ff6b6b', fg='white', font=('Arial', 12), padx=20).pack(side=tk.LEFT, padx=5)
        
        tk.Button(button_frame, text="Buchstabe speichern", command=self.save_letter,
                  bg='#4ecdc4', fg='white', font=('Arial', 12), padx=20).pack(side=tk.LEFT, padx=5)
        
        tk.Button(button_frame, text="Überspringen", command=self.next_letter,
                  bg='#95a5a6', fg='white', font=('Arial', 12), padx=20).pack(side=tk.LEFT, padx=5)
        
        # Progress display
        self.progress_label = tk.Label(
            self.root,
            text=f"Fortschritt: {len(self.letter_db)}/{self.alphabet.__len__()} Buchstaben gelernt",
            font=('Arial', 10)
        )
        self.progress_label.pack(pady=5)
        
        self.strokes = []
    
    def create_text_gui(self):
        # Title
        tk.Label(
            self.root,
            text="Text in OneNote schreiben",
            font=('Arial', 20, 'bold')
        ).pack(pady=20)
        
        # Info
        info_frame = tk.Frame(self.root)
        info_frame.pack(pady=5)
        
        tk.Label(info_frame, text="Gib einen Text ein (\\n für manuellen Zeilenumbruch)",
                 font=('Arial', 12)).pack()
        tk.Label(info_frame, text="🛑 FAILSAFE: Bewege Maus in obere linke Ecke zum Stoppen!",
                 font=('Arial', 10), fg='red').pack()
        
        # Text input
        text_frame = tk.Frame(self.root)
        text_frame.pack(pady=10)
        
        tk.Label(text_frame, text="Text:", font=('Arial', 12)).pack(anchor='w', padx=5)
        
        inner_frame = tk.Frame(text_frame)
        inner_frame.pack()
        
        self.input_text = tk.Text(inner_frame, font=('Arial', 12), width=50, height=8, wrap='word')
        self.input_text.pack(side=tk.LEFT, padx=5)
        
        scrollbar = tk.Scrollbar(inner_frame, command=self.input_text.yview)
        scrollbar.pack(side=tk.LEFT, fill='y')
        self.input_text.config(yscrollcommand=scrollbar.set)
        
        # Settings
        settings = tk.Frame(self.root)
        settings.pack(pady=10)
        
        tk.Label(settings, text="Zeichen pro Zeile:", font=('Arial', 10)).grid(row=0, column=0, padx=5)
        self.chars_per_line = tk.Spinbox(settings, from_=10, to=50, width=10, font=('Arial', 10))
        self.chars_per_line.delete(0, 'end')
        self.chars_per_line.insert(0, '25')
        self.chars_per_line.grid(row=0, column=1, padx=5)
        
        tk.Label(settings, text="Buchstabegröße:", font=('Arial', 10)).grid(row=1, column=2, padx=5)
        self.characterSize = tk.Spinbox(settings, from_=0.1, to=1, increment=0.1, width=10, font=('Arial', 10))
        self.characterSize.delete(0, 'end')
        self.characterSize.insert(0, '0.2')
        self.characterSize.grid(row=1, column=3, padx=5)
        
        tk.Label(settings, text="Zeilenabstand:", font=('Arial', 10)).grid(row=1, column=0, padx=5)
        self.line_spacing = tk.Spinbox(settings, from_=50, to=400, increment=10, width=10, font=('Arial', 10))
        self.line_spacing.delete(0, 'end')
        self.line_spacing.insert(0, '60')
        self.line_spacing.grid(row=1, column=1, padx=5)
        
        # Draw button
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=20)
        
        tk.Button(button_frame, text="In OneNote zeichnen", command=self.draw_text_in_onenote,
                  bg='#7719AA', fg='white', font=('Arial', 14, 'bold'),
                  padx=30, pady=10).pack(side=tk.LEFT, padx=5)
        
        self.stop_button = tk.Button(button_frame, text="⏹ STOPP", command=self.stop_drawing_now,
                  bg='#e74c3c', fg='white', font=('Arial', 14, 'bold'),
                  padx=20, pady=10, state='disabled')
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        # Status
        self.status_label = tk.Label(self.root, text="Bereit zum Zeichnen!", font=('Arial', 10))
        self.status_label.pack(pady=10)
        
        # Options
        options = tk.Frame(self.root)
        options.pack(pady=10)
        
        tk.Button(options, text="Buchstaben neu lernen", command=self.reset_learning,
                  bg='#e74c3c', fg='white', font=('Arial', 10)).pack(side=tk.LEFT, padx=5)
        
        tk.Button(options, text="Einzelnen Buchstaben ändern", command=self.change_letter,
                  bg='#95a5a6', fg='white', font=('Arial', 10)).pack(side=tk.LEFT, padx=5)

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
                                    capstyle=tk.ROUND, smooth=True)
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
            self.title_label.config(text=f"Lerne Buchstabe: {self.current_letter.upper()}")
            self.progress_label.config(text=f"Fortschritt: {len(self.letter_db)}/{self.alphabet.__len__()} Buchstaben gelernt")
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
        self.status_label.config(text="⏹ Zeichnen gestoppt!")
    
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
            self.status_label.config(text="Bitte Text eingeben!")
            return
        
        # Check trained letters
        for ch in text:
            if ch and ch != " " and ch != '\n'and ch not in self.letter_db:
                self.status_label.config(text=f"Buchstabe '{ch}' wurde nicht gelernt!")
                return
        
        self.stop_drawing = False
        self.stop_button.config(state='normal')
        self.status_label.config(text="Öffne jetzt OneNote, 4 Sekunden Zeit!")
        
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
            self.root.after(0, lambda: self.stop_button.config(state='disabled'))
            return
        
        hwnd = self.find_onenote_window()
        if not hwnd:
            self.root.after(0, lambda: self.status_label.config(text="Fehler: OneNote nicht gefunden!"))
            self.root.after(0, lambda: self.stop_button.config(state='disabled'))
            return
        
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.4)
        
        rect = win32gui.GetWindowRect(hwnd)
        canvas_x = rect[0] + 20
        canvas_y = rect[1] + 100

        line_spacing_px = int(self.line_spacing.get())
        scale = float(self.characterSize.get())

        chars_per_line = int(6 / scale)
        
        base_char_spacing_px = int(150 * scale)
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
                
                # Manuelle Zeilenlogik
                if chars_in_line >= chars_per_line and ch == ' ':
                    current_lines += 1
                    offset_y += line_spacing_px
                    offset_x = 50
                    chars_in_line = 0
                    time.sleep(0.7)
                    continue
                
                # Leerzeichen
                if ch == ' ':
                    offset_x += base_char_spacing_px
                    chars_in_line += 1
                    continue
                
                # Überlauf → scrollen
                if current_lines >= 5:
                    win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, -1000)
                    current_lines = 0
                    offset_y = 50
                    time.sleep(1)
                
                # Buchstabe prüfen
                if ch not in self.letter_db:
                    continue
                
                raw_width = self.get_letter_width(ch)

                MIN_WIDTH = 60   # Mindestbreite für i, l, ! usw.

                effective_width = max(raw_width, MIN_WIDTH)

                letter_spacing = int(effective_width * scale * 1.15)

                strokes = self.letter_db[ch]
                self.root.after(0, lambda c=ch: self.status_label.config(text=f"Zeichne '{c}'..."))
                
                # 🔥 Zeichnen des Buchstabens
                for stroke in strokes:
                    if self.stop_drawing or self.failsafe():
                        break
                    
                    if len(stroke) < 2:
                        continue
                    
                    # Startpunkt setzen
                    start_x, start_y = stroke[0]
                    sx = canvas_x + int(start_x * scale) + offset_x
                    sy = canvas_y + int(start_y * scale) + offset_y
                    
                    win32api.SetCursorPos((int(sx), int(sy)))
                    time.sleep(0.003)
                    
                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
                    time.sleep(0.0003)
                    
                    last_x, last_y = sx, sy
                    
                    # Stroke Linien zeichnen
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
                
                # 🔥 Am Ende des Buchstabens den korrekten Abstand setzen
                offset_x += letter_spacing
                chars_in_line += 1
            
            # Neue Textzeile
            offset_y += line_spacing_px
            current_lines += 1
        
        # Abschlussmeldung
        if self.stop_drawing:
            self.root.after(0, lambda: self.status_label.config(text="⏹ Zeichnen wurde gestoppt!"))
        else:
            self.root.after(0, lambda: self.status_label.config(text="✅ Fertig!"))
        
        self.root.after(0, lambda: self.stop_button.config(state='disabled'))


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
        win = tk.Toplevel(self.root)
        win.title("Buchstabe ändern")
        win.geometry("300x200")
        
        tk.Label(win, text="Welchen Buchstaben ändern?", font=('Arial', 12)).pack(pady=20)
        
        var = tk.StringVar(value='a')
        combo = ttk.Combobox(win, textvariable=var, values=list(self.alphabet), font=('Arial', 14))
        combo.pack(pady=10)
        
        def apply():
            self.current_letter = var.get()
            self.learning_mode = True
            win.destroy()
            for w in self.root.winfo_children():
                w.destroy()
            self.create_learning_gui()
        
        tk.Button(win, text="Ändern", command=apply,
                 bg='#4ecdc4', fg='white', font=('Arial', 12)).pack(pady=20)


if __name__ == "__main__":
    root = tk.Tk()
    app = LetterApp(root)
    root.mainloop()
