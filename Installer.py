import os
import shutil
import tkinter as tk
from tkinter import messagebox, ttk

import requests

from backend.paths import (APP_DATA_DIR, IS_LINUX, IS_WINDOWS, ensure_app_data_dir,
                           exe_name, release_asset_name)

if IS_WINDOWS:
    import win32com.client

APP_NAME = "Text2Pen"

# Local filenames — no CPU-architecture tag; that only matters for picking
# the right download, below.
INSTALL_DIR = APP_DATA_DIR
EXE_NAME = exe_name("Text2Pen")
UPDATE_NAME = exe_name("Update")
EXE_PATH = os.path.join(INSTALL_DIR, EXE_NAME)
UPDATE_PATH = os.path.join(INSTALL_DIR, UPDATE_NAME)

# Release asset names are architecture-specific on Linux (a release holds an
# x86_64 and an arm64 binary side by side).
DOWNLOAD_URL = f"https://github.com/pythonIsFast/Text2Pen/releases/latest/download/{release_asset_name('Text2Pen')}"
UPDATE_URL = f"https://github.com/pythonIsFast/Text2Pen/releases/latest/download/{release_asset_name('Update')}"

if IS_WINDOWS:
    STARTUP_DIR = os.path.join(
        os.environ["APPDATA"], r"Microsoft\Windows\Start Menu\Programs\Startup"
    )
    START_MENU = os.path.join(os.environ["APPDATA"], r"Microsoft\Windows\Start Menu\Programs")
else:
    AUTOSTART_DIR = os.path.join(os.path.expanduser("~"), ".config", "autostart")
    APPLICATIONS_DIR = os.path.join(os.path.expanduser("~"), ".local", "share", "applications")


def create_shortcut(target, shortcut_path, working_dir):
    """Windows Start Menu / Startup entry (a .lnk file)."""
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(shortcut_path)
    shortcut.TargetPath = target
    shortcut.WorkingDirectory = working_dir
    shortcut.IconLocation = target
    shortcut.save()


def create_desktop_entry(target, path, name, hidden=False, autostart=False):
    """Linux equivalent of a shortcut: an XDG .desktop file."""
    lines = [
        "[Desktop Entry]",
        "Type=Application",
        f"Name={name}",
        f'Exec="{target}"',
        f"Path={os.path.dirname(target)}",
        "Terminal=false",
    ]
    if hidden:
        lines.append("NoDisplay=true")
    if autostart:
        lines.append("X-GNOME-Autostart-enabled=true")
    else:
        lines.append("Categories=Utility;")
    with open(path, "w") as f:
        f.write("\n".join(lines) + "\n")
    os.chmod(path, 0o755)


def create_launcher(target, name):
    """App menu entry (Start Menu shortcut / .desktop in applications/)."""
    if IS_WINDOWS:
        create_shortcut(target, os.path.join(START_MENU, f"{name}.lnk"), os.path.dirname(target))
    else:
        os.makedirs(APPLICATIONS_DIR, exist_ok=True)
        create_desktop_entry(target, os.path.join(APPLICATIONS_DIR, f"{name}.desktop"), name)


def create_autostart_entry(target, name):
    """Run-on-login entry for Update (Startup shortcut / autostart .desktop)."""
    if IS_WINDOWS:
        os.makedirs(STARTUP_DIR, exist_ok=True)
        create_shortcut(target, os.path.join(STARTUP_DIR, f"{name}.lnk"), os.path.dirname(target))
    else:
        os.makedirs(AUTOSTART_DIR, exist_ok=True)
        create_desktop_entry(target, os.path.join(AUTOSTART_DIR, f"{name}.desktop"), name,
                              hidden=True, autostart=True)


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

    if IS_LINUX:
        os.chmod(target_path, 0o755)


def download_and_install(progress_var, progress_bar):
    ensure_app_data_dir()

    try:
        download_file(
            DOWNLOAD_URL,
            EXE_PATH,
            lambda p: (progress_var.set(p), progress_bar.update())
        )

        download_file(UPDATE_URL, UPDATE_PATH)

        create_launcher(EXE_PATH, APP_NAME)
        create_autostart_entry(UPDATE_PATH, "Update")

        message = f"{APP_NAME} installed successfully!"
        if IS_LINUX:
            message += (
                "\n\nOne more step: Text2Pen writes strokes via /dev/uinput, "
                "which is root-owned by default. To use it without sudo, run:\n\n"
                'echo \'KERNEL=="uinput", GROUP="input", MODE="0660"\' | '
                "sudo tee /etc/udev/rules.d/99-uinput.rules\n"
                "sudo udevadm control --reload-rules && sudo udevadm trigger\n"
                'sudo usermod -aG input "$USER"\n\n'
                "Then log out and back in."
            )
        messagebox.showinfo("Success!", message)

    except Exception as e:
        messagebox.showerror("Error", str(e))


def uninstall():
    if IS_WINDOWS:
        shortcut_paths = [
            os.path.join(STARTUP_DIR, "Update.lnk"),
            os.path.join(START_MENU, f"{APP_NAME}.lnk"),
        ]
    else:
        shortcut_paths = [
            os.path.join(AUTOSTART_DIR, "Update.desktop"),
            os.path.join(APPLICATIONS_DIR, f"{APP_NAME}.desktop"),
        ]

    for path in shortcut_paths:
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
