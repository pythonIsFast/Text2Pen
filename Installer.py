import os
import requests
import tkinter as tk
from tkinter import messagebox, ttk
import shutil
import win32com.client

def create_shortcut(target, shortcut_path, working_dir):
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(shortcut_path)
    shortcut.TargetPath = target
    shortcut.WorkingDirectory = working_dir
    shortcut.IconLocation = target
    shortcut.save()

STARTUP_DIR = os.path.join(
    os.environ["APPDATA"],
    r"Microsoft\Windows\Start Menu\Programs\Startup"
)

APP_NAME = "Text2Pen"

INSTALL_DIR = os.path.join(os.environ["LOCALAPPDATA"], APP_NAME)
EXE_PATH = os.path.join(INSTALL_DIR, "Text2Pen.exe")

DOWNLOAD_URL = "https://github.com/pythonIsFast/Text2Pen/releases/latest/download/Text2Pen.exe"

UPDATE_URL = "https://github.com/pythonIsFast/Text2Pen/releases/latest/download/Update.exe"
UPDATE_PATH = os.path.join(INSTALL_DIR, "Update.exe")

def download_file(url, target_path, progress_callback=None):
    r = requests.get(url, stream=True)
    r.raise_for_status()

    total = int(r.headers.get("content-length", 0))
    downloaded = 0

    with open(target_path, "wb") as f:
        for chunk in r.iter_content(8192):
            if chunk:
                f.write(chunk)
                downloaded += len(chunk)
                if total and progress_callback:
                    progress_callback(int(downloaded / total * 100))

def create_startup_shortcut(target, name):
    os.makedirs(STARTUP_DIR, exist_ok=True)

    shortcut_path = os.path.join(STARTUP_DIR, f"{name}.lnk")

    create_shortcut(target, shortcut_path, os.path.dirname(target))

def download_and_install(progress_var, progress_bar):
    os.makedirs(INSTALL_DIR, exist_ok=True)

    try:
        download_file(
            DOWNLOAD_URL,
            EXE_PATH,
            lambda p: (progress_var.set(p), progress_bar.update())
        )

        download_file(UPDATE_URL, UPDATE_PATH)

        START_MENU = os.path.join(
            os.environ["APPDATA"],
            r"Microsoft\Windows\Start Menu\Programs"
        )

        create_shortcut(
            EXE_PATH,
            os.path.join(START_MENU, f"{APP_NAME}.lnk"),
            INSTALL_DIR
        )

        create_startup_shortcut(UPDATE_PATH, "Update")

        messagebox.showinfo(
            "Success!",
            f"{APP_NAME} installed successfully!"
        )

    except Exception as e:
        messagebox.showerror("Error", str(e))

def uninstall():
    #Remove automatic startup shortcut
    path = os.path.join(STARTUP_DIR, "Update.lnk")
    if os.path.exists(path):
        os.remove(path)

    if os.path.exists(INSTALL_DIR):
        shutil.rmtree(INSTALL_DIR)

    messagebox.showinfo("Uninstalled!", f"{APP_NAME} got removed.")

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

    tk.Button(frame, text="Install", font=("Arial", 12),
              command=lambda: download_and_install(progress_var, progress_bar), width=25).pack(pady=5)
    tk.Button(frame, text="Uninstall", font=("Arial", 12),
              command=uninstall, width=25).pack(pady=5)
    tk.Button(frame, text="Quit", font=("Arial", 12),
              command=root.quit, width=25).pack(pady=5)

    root.mainloop()


if __name__ == "__main__":
    main()
