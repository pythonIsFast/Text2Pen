import os
import requests
import subprocess
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import threading
import io
import hashlib

from backend.paths import APP_DATA_DIR, IS_LINUX, exe_name

APP_NAME = "Text2Pen"

# Installation paths
INSTALL_DIR = APP_DATA_DIR
TEXT2PEN_NAME = exe_name("Text2Pen")
UPDATE_NAME = exe_name("Update")
TEXT2PEN_PATH = os.path.join(INSTALL_DIR, TEXT2PEN_NAME)
UPDATE_PATH = os.path.join(INSTALL_DIR, UPDATE_NAME)

#Temp paths
UPDATE_TEMP = UPDATE_PATH + "-newest"

# URLs
TEXT2PEN_URL = f"https://github.com/pythonIsFast/Text2Pen/releases/latest/download/{TEXT2PEN_NAME}"
UPDATE_URL = f"https://github.com/pythonIsFast/Text2Pen/releases/latest/download/{UPDATE_NAME}"

class UpdateApp():
    def __init__(self, root):
        self.root = root
        self.root.title("Text2Pen Updater")
        self.root.geometry("400x150")
        self.root.resizable(False, False)

        tk.Label(root, text="Updating Text2Pen...", font=("Arial", 14)).pack(pady=10)
        
        self.progress_var = tk.IntVar()
        self.progress_bar = ttk.Progressbar(root, orient="horizontal", length=350,
                                            mode="determinate", variable=self.progress_var)
        self.progress_bar.pack(pady=20)

        self.status_label = tk.Label(root, text="Starting...", font=("Arial", 10))
        self.status_label.pack(pady=5)

        threading.Thread(target=self.main, daemon=True).start()

    def main(self):
        try:
            print("Checking GitHub release...")
            hashes = self.get_release_hashes()

            github_hash = hashes[TEXT2PEN_NAME]

            if os.path.exists(TEXT2PEN_PATH):
                local_hash = self.sha256_file(TEXT2PEN_PATH)

                if local_hash == github_hash:
                    print("Text2Pen already up to date")
                    new_hash = local_hash
                    new_data = None
                else:
                    self.status_label.config(text="Downloading Text2Pen...")
                    new_data = self.download_file_ram(TEXT2PEN_URL)
                    new_hash = self.sha256_bytes(new_data)

                    if new_hash != github_hash:
                        raise ValueError(f"Downloaded {TEXT2PEN_NAME} failed SHA256 verification")
            else:
                self.status_label.config(text="Downloading Text2Pen...")
                new_data = self.download_file_ram(TEXT2PEN_URL)
                new_hash = self.sha256_bytes(new_data)

                if new_hash != github_hash:
                    raise ValueError(f"Downloaded {TEXT2PEN_NAME} failed SHA256 verification")

            if os.path.exists(TEXT2PEN_PATH):
                if new_data is not None:
                    temp_path = TEXT2PEN_PATH + ".new"
                    self.write_file_once(temp_path, new_data)
                    self.replace_file(temp_path, TEXT2PEN_PATH)
            else:
                print("Adding file")
                temp_path = TEXT2PEN_PATH + ".new"
                self.write_file_once(temp_path, new_data)
                self.replace_file(temp_path, TEXT2PEN_PATH)
            
            github_updater_hash = hashes[UPDATE_NAME]

            if os.path.exists(UPDATE_PATH):
                local_updater_hash = self.sha256_file(UPDATE_PATH)

                if local_updater_hash == github_updater_hash:
                    updater_data = None
                else:
                    self.status_label.config(text="Downloading Updater...")
                    updater_data = self.download_file_ram(UPDATE_URL)

                    new_updater_hash = self.sha256_bytes(updater_data)

                    if new_updater_hash != github_updater_hash:
                        raise ValueError(f"Downloaded {UPDATE_NAME} failed SHA256 verification")
            else:
                self.status_label.config(text="Downloading Updater...")
                updater_data = self.download_file_ram(UPDATE_URL)

                new_updater_hash = self.sha256_bytes(updater_data)

                if new_updater_hash != github_updater_hash:
                    raise ValueError(f"Downloaded {UPDATE_NAME} failed SHA256 verification")

            if os.path.exists(UPDATE_PATH):
                if updater_data is not None:
                    self.write_file_once(UPDATE_TEMP, updater_data)
            else:
                self.write_file_once(UPDATE_TEMP, updater_data)

            self.status_label.config(text="Update finished!")
            self.progress_var.set(100)
            self.root.update_idletasks()

            print("Starting Text2Pen...")
            subprocess.Popen([TEXT2PEN_PATH, "updaterStart"], close_fds=True)
            self.root.destroy()

        except Exception as e:
            messagebox.showerror("Error", f"Update failed:\n{e}")
            self.root.destroy()

    def download_file_ram(self, url):
        r = requests.get(url, stream=True)
        r.raise_for_status()

        total = int(r.headers.get("content-length", 0))
        downloaded = 0

        buffer = io.BytesIO()  # RAM

        for chunk in r.iter_content(8192):
            if chunk:
                buffer.write(chunk)
                downloaded += len(chunk)

                if total:
                    percent = int(downloaded / total * 100)
                    self.progress_var.set(percent)
                    self.status_label.config(
                        text=f"Downloading {os.path.basename(url)}... {percent}%"
                    )
                    self.root.update_idletasks()

        buffer.seek(0)
        return buffer.getvalue()

    def sha256_bytes(self, data: bytes):
        return hashlib.sha256(data).hexdigest()

    def sha256_file(self, path):
        h = hashlib.sha256()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(8192), b""):
                h.update(chunk)
        return h.hexdigest()

    def write_file_once(self, path, data: bytes):
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "wb") as f:
            f.write(data)
        if IS_LINUX:
            os.chmod(path, 0o755)

    def replace_file(self, src, dst):
        try:
            os.replace(src, dst)
        except Exception as e:
            print(f"Error while replacing file! {dst}: {e}")

    def get_release_hashes(self):
        api = "https://api.github.com/repos/pythonIsFast/Text2Pen/releases/latest"
        r = requests.get(api, headers={"User-Agent": "Text2Pen-Updater"}, timeout=30)
        r.raise_for_status()

        release = r.json()

        hashes = {}
        for asset in release["assets"]:
            hashes[asset["name"]] = asset["digest"].replace("sha256:", "")

        return hashes

def on_close():
    messagebox.showinfo(
        "Update running",
        "Please wait for Text2Pen to finish the update."
    )

if __name__ == "__main__":
    root = tk.Tk()

    root.protocol("WM_DELETE_WINDOW", on_close)

    app = UpdateApp(root)
    root.mainloop()


