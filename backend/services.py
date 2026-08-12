"""HTTP clients for the Text2Pen backend: AI actions, chat, limits, telemetry."""

import os
import platform
import time

import requests

BASE_URL = "https://text2pen-backend.onrender.com"
AI_URL = f"{BASE_URL}/ai"
AI_CHAT_URL = f"{BASE_URL}/ai/chat"
AI_LIMITS_URL = f"{BASE_URL}/ai/limits"
TELEMETRY_URL = f"{BASE_URL}/telemetry"

WEBSITE_URL = "https://text2pen.onrender.com"
STATUS_URL = "https://m9rgm51d.status.cron-job.org"


class AIClient:
    """All calls return a dict with an `ok` flag so the UI can render failures."""

    def limits(self):
        try:
            resp = requests.get(AI_LIMITS_URL, timeout=10)
            if resp.status_code == 200:
                return {"ok": True, "limits": resp.json()}
            return {"ok": False, "error": f"HTTP {resp.status_code}"}
        except requests.RequestException as e:
            return {"ok": False, "error": str(e)}

    def action(self, action, text, extra=None):
        payload = {"action": action, "text": text}
        if extra:
            payload["extra"] = extra
        try:
            resp = requests.post(AI_URL, json=payload, timeout=30)
            if resp.status_code == 200:
                return {"ok": True, "result": resp.json().get("result", "")}
            if resp.status_code == 429:
                return {"ok": False, "error": "limit", "message": "Daily limit reached."}
            return {"ok": False, "error": f"HTTP {resp.status_code}"}
        except requests.exceptions.Timeout:
            return {"ok": False, "error": "timeout", "message": "The server did not respond in time."}
        except requests.exceptions.ConnectionError:
            return {"ok": False, "error": "offline", "message": "Could not reach the server."}
        except requests.RequestException as e:
            return {"ok": False, "error": str(e)}

    def chat(self, history, message, image_b64=None):
        payload = {"history": history, "message": message, "image_b64": image_b64}
        try:
            # Connect must succeed quickly; the read side waits, because the
            # backend enforces its own 90 s ceiling before it answers.
            resp = requests.post(AI_CHAT_URL, json=payload, timeout=(10, None))
            if resp.status_code == 200:
                data = resp.json()
                return {"ok": True, "result": data.get("result", ""),
                        "queue_position": data.get("queue_position")}
            if resp.status_code == 429:
                return {"ok": False, "error": "limit",
                        "message": "Daily limit reached. Please try again tomorrow."}
            if resp.status_code == 503:
                pos = resp.json().get("queue_position", "?")
                return {"ok": True, "queued": True, "queue_position": pos,
                        "result": f"You are in the queue (position {pos}). Please wait…"}
            return {"ok": False, "error": f"HTTP {resp.status_code}"}
        except requests.exceptions.Timeout:
            return {"ok": False, "error": "timeout",
                    "message": "Could not connect to the server. Check your internet connection."}
        except requests.exceptions.ConnectionError:
            return {"ok": False, "error": "offline",
                    "message": "Could not reach the server. Check your internet connection."}
        except requests.RequestException as e:
            return {"ok": False, "error": str(e)}


class Telemetry:
    def __init__(self, enabled=False):
        self.enabled = bool(enabled)

    @staticmethod
    def sanitize(message):
        msg = str(message)
        user = (os.environ.get("USERNAME") or os.environ.get("USER")
                or os.environ.get("LOGNAME") or "")
        if user:
            msg = msg.replace(user, "<user>")
        return msg[:500]

    def send_error(self, message):
        if not self.enabled:
            return False
        event = {
            "event": "ERROR",
            "os": platform.system(),
            "timestamp": time.time(),
            "extra": {"Exact Error": self.sanitize(message)},
        }
        try:
            resp = requests.post(TELEMETRY_URL, json=event, timeout=60)
            return resp.status_code == 200
        except requests.RequestException:
            return False
