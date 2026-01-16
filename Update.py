import os
import sys
import requests
import shutil
import subprocess
import time
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk

APP_NAME = "Text2Pen"

# Installation paths
INSTALL_DIR = os.path.join(os.environ["LOCALAPPDATA"], APP_NAME)
TEXT2PEN_PATH = os.path.join(INSTALL_DIR, "Text2Pen.exe")
UPDATE_PATH = os.path.join(INSTALL_DIR, "Update.exe")

#Temp paths
UPDATE_TEMP = UPDATE_PATH + "-newest"

# URLs
TEXT2PEN_URL = "https://raw.githubusercontent.com/pythonIsFast/Text2Pen/main/Text2Pen.exe"
UPDATE_URL = "https://raw.githubusercontent.com/pythonIsFast/Text2Pen/main/Update.exe"

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

        self.root.after(100, self.main)

    def download_file(self, url, target_path):
        os.makedirs(os.path.dirname(target_path), exist_ok=True)
        r = requests.get(url, stream=True)
        r.raise_for_status()
        total = int(r.headers.get("content-length", 0))
        downloaded = 0
        with open(target_path, "wb") as f:
            for chunk in r.iter_content(8192):
                if chunk:
                    f.write(chunk)
                    downloaded += len(chunk)
                    if total:
                        percent = int(downloaded / total * 100)
                        print(f"\r{os.path.basename(target_path)}: {percent}%", end="")
                        self.progress_var.set(percent)
                        self.status_label.config(text=f"Downloading {os.path.basename(target_path)}... {percent}%")
                        self.root.update_idletasks()

        print("\Finished:", os.path.basename(target_path))

    def replace_file(self, src, dst):
        try:
            if os.path.exists(dst):
                os.remove(dst)
            shutil.move(src, dst)
        except Exception as e:
            print(f"Error while replacing file! {dst}: {e}")

    def main(self):
        #Download newest version
        try:
            print("Starting Download...")
            self.download_file(TEXT2PEN_URL, TEXT2PEN_PATH)
            self.download_file(UPDATE_URL, UPDATE_TEMP)

            self.status_label.config(text="Update finished!")
            self.progress_var.set(100)
            self.root.update_idletasks()

            #Starting Text2Pen
            print("Starting Text2Pen...")
            subprocess.Popen([TEXT2PEN_PATH])
            self.root.destroy()

        except Exception as e:
            messagebox.showerror("Error", f"Update failed:\n{e}")
            self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = UpdateApp(root)
    root.mainloop()