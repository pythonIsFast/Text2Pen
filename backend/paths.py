"""Platform detection and application paths."""

import os
import platform
import sys

IS_WINDOWS = platform.system() == "Windows"
IS_LINUX = platform.system() == "Linux"
IS_WAYLAND = bool(os.environ.get("WAYLAND_DISPLAY"))

APP_NAME = "Text2Pen"


def _real_home():
    """Home directory of the human user, even when started via sudo.

    Under `sudo` the HOME variable points at /root, which would hide the
    letter database the user actually trained. SUDO_USER lets us recover it.
    """
    if IS_LINUX and hasattr(os, "geteuid") and os.geteuid() == 0:
        sudo_user = os.environ.get("SUDO_USER")
        if sudo_user and sudo_user != "root":
            try:
                import pwd
                return pwd.getpwnam(sudo_user).pw_dir
            except (ImportError, KeyError):
                pass
    return os.path.expanduser("~")


def get_app_data_dir():
    if IS_WINDOWS:
        base = os.environ.get("LOCALAPPDATA") or _real_home()
        return os.path.join(base, APP_NAME)
    return os.path.join(_real_home(), ".local", "share", APP_NAME)


def exe_name(base):
    """`base` with the platform-appropriate executable suffix: Installer.py,
    Update.py and Text2Pen.py all ship as "<name>.exe" on Windows and a
    plain, extension-less "<name>" on Linux (matches the release asset
    names uploaded by .github/workflows/build.yml)."""
    return f"{base}.exe" if IS_WINDOWS else base


APP_DATA_DIR = get_app_data_dir()

LETTER_DB_FILE = os.path.join(APP_DATA_DIR, "letter_db.json")
SETTINGS_FILE = os.path.join(APP_DATA_DIR, "settingsDB.json")
TEXT_FILE = os.path.join(APP_DATA_DIR, "saved_text.txt")
UPDATE_EXE = os.path.join(APP_DATA_DIR, exe_name("Update"))


def resource_dir():
    """Directory holding bundled read-only resources (webui/).

    PyInstaller unpacks --add-data payloads into sys._MEIPASS.
    """
    bundled = getattr(sys, "_MEIPASS", None)
    if bundled:
        return bundled
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def webui_index():
    return os.path.join(resource_dir(), "webui", "index.html")


def ensure_app_data_dir():
    os.makedirs(APP_DATA_DIR, exist_ok=True)
    return APP_DATA_DIR


def is_root():
    return IS_LINUX and hasattr(os, "geteuid") and os.geteuid() == 0
