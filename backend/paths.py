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
    """`base` with the platform-appropriate executable suffix, for the LOCAL
    on-disk filename: Installer.py, Update.py and Text2Pen.py all install as
    "<name>.exe" on Windows and a plain, extension-less "<name>" on Linux,
    regardless of CPU architecture. For the GitHub release asset name (which
    does vary by architecture, since one release holds every arch's binary
    side by side), see release_asset_name()."""
    return f"{base}.exe" if IS_WINDOWS else base


_LINUX_ARCH_TAGS = {"x86_64": "x86_64", "aarch64": "arm64", "arm64": "arm64"}


def release_asset_name(base):
    """`base`'s filename as uploaded to a GitHub Release. Matches the
    `label` values in .github/workflows/build.yml's matrix, e.g.
    "Text2Pen-linux-x86_64" or "Text2Pen-linux-arm64" — used to build
    download URLs and to look up the right SHA256 from the release's asset
    list (see Update.py, Installer.py)."""
    if IS_WINDOWS:
        return f"{base}.exe"
    arch = platform.machine()
    arch_tag = _LINUX_ARCH_TAGS.get(arch, arch)
    return f"{base}-linux-{arch_tag}"


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
