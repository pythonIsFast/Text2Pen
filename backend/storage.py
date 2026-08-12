"""Persistence for the letter database, settings and the draft text.

Every file lives in APP_DATA_DIR so the app finds its data regardless of the
working directory it was launched from.
"""

import json
import os
import shutil

from .paths import LETTER_DB_FILE, SETTINGS_FILE, TEXT_FILE, ensure_app_data_dir

ALPHABET = (
    "abcdefghijklmnopqrstuvwxyzäöüÄÖÜß"
    "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    "0123456789"
    ".,!?;:-_—()[]{}<>\"'/*+=@#$%^&|~`"
)

DEFAULT_SETTINGS = {
    "line_spacing": 60,
    "character_size": 0.1,
    "telemetry_opted_in": None,
    "show_overlay": True,
    "start_delay": 4,
}


class Storage:
    def __init__(self):
        ensure_app_data_dir()

    # ---------------------------------------------------------------- letters
    def load_letters(self):
        """Return the trained strokes. A corrupt file is quarantined, not fatal."""
        if not os.path.exists(LETTER_DB_FILE):
            return {}
        try:
            with open(LETTER_DB_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, dict):
                raise ValueError("letter_db.json does not contain an object")
            return data
        except (json.JSONDecodeError, ValueError, OSError) as e:
            backup = LETTER_DB_FILE + ".corrupt"
            try:
                shutil.copy2(LETTER_DB_FILE, backup)
            except OSError:
                backup = "(backup failed)"
            print(f"letter_db.json unreadable ({e}); kept a copy at {backup}")
            return {}

    def save_letters(self, letter_db):
        self._write_json(LETTER_DB_FILE, letter_db)

    # --------------------------------------------------------------- settings
    def load_settings(self):
        settings = DEFAULT_SETTINGS.copy()
        first_run = not os.path.exists(SETTINGS_FILE)
        if not first_run:
            try:
                with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                    stored = json.load(f)
                if isinstance(stored, dict):
                    settings.update(stored)
            except (json.JSONDecodeError, OSError) as e:
                print(f"settingsDB.json unreadable ({e}); falling back to defaults")
        return settings, first_run

    def save_settings(self, settings):
        self._write_json(SETTINGS_FILE, settings, indent=2)

    # ------------------------------------------------------------------- text
    def load_text(self):
        if not os.path.exists(TEXT_FILE):
            return ""
        try:
            with open(TEXT_FILE, "r", encoding="utf-8") as f:
                return f.read()
        except OSError:
            return ""

    def save_text(self, content):
        try:
            with open(TEXT_FILE, "w", encoding="utf-8") as f:
                f.write(content)
            return True
        except OSError as e:
            print(f"Error saving text: {e}")
            return False

    # ---------------------------------------------------------------- helpers
    @staticmethod
    def _write_json(path, payload, indent=None):
        """Write atomically so a crash mid-write cannot truncate the original."""
        ensure_app_data_dir()
        tmp = path + ".tmp"
        try:
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(payload, f, indent=indent, ensure_ascii=False)
            os.replace(tmp, path)
            return True
        except OSError as e:
            print(f"Error writing {path}: {e}")
            try:
                if os.path.exists(tmp):
                    os.remove(tmp)
            except OSError:
                pass
            return False
