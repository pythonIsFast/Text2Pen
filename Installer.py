import os
import requests
import tkinter as tk
from tkinter import messagebox, ttk
import shutil

APP_NAME = "Text2Pen"
DOWNLOAD_URL = "https://raw.githubusercontent.com/pythonIsFast/Text2Pen/main/Text2Pen.exe"

# Installationsverzeichnis (nur für den Benutzer)
INSTALL_DIR = os.path.join(os.environ["LOCALAPPDATA"], APP_NAME)
EXE_PATH = os.path.join(INSTALL_DIR, "Text2Pen.exe")


def download_and_install(progress_var, progress_bar):
    os.makedirs(INSTALL_DIR, exist_ok=True)

    try:
        r = requests.get(DOWNLOAD_URL, stream=True)
        r.raise_for_status()
        total_length = int(r.headers.get('content-length', 0))
        downloaded = 0

        with open(EXE_PATH, "wb") as f:
            for chunk in r.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
                    downloaded += len(chunk)
                    if total_length > 0:
                        progress = int(downloaded / total_length * 100)
                        progress_var.set(progress)
                        progress_bar.update()

        messagebox.showinfo("Erfolg", f"{APP_NAME} wurde erfolgreich installiert!")
    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Installieren:\n{e}")


def uninstall():
    if os.path.exists(INSTALL_DIR):
        shutil.rmtree(INSTALL_DIR)
    messagebox.showinfo("Deinstalliert", f"{APP_NAME} wurde entfernt.")


# GUI-Fenster mit Fortschrittsbalken
def main():
    root = tk.Tk()
    root.title("Text2Pen Installer")
    root.geometry("350x220")
    root.resizable(False, False)

    frame = tk.Frame(root)
    frame.pack(expand=True, padx=10, pady=10)

    tk.Label(frame, text="Text2Pen Installer", font=("Arial", 16)).pack(pady=10)

    progress_var = tk.IntVar()
    progress_bar = ttk.Progressbar(frame, orient="horizontal", length=300, mode="determinate", variable=progress_var)
    progress_bar.pack(pady=10)

    tk.Button(frame, text="Installieren", font=("Arial", 12),
              command=lambda: download_and_install(progress_var, progress_bar), width=25).pack(pady=5)
    tk.Button(frame, text="Deinstallieren", font=("Arial", 12),
              command=uninstall, width=25).pack(pady=5)
    tk.Button(frame, text="Beenden", font=("Arial", 12),
              command=root.quit, width=25).pack(pady=5)

    root.mainloop()


if __name__ == "__main__":
    main()
