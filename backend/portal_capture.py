"""Screen capture via the XDG ScreenCast portal, for picking a drawing origin
on Wayland where the real cursor position cannot be queried reliably (see
input_controller.py for the long version of why).

Signal/timeout callbacks from dbus-python and GLib always dispatch on the
*global default* main context regardless of which thread schedules them
(g_timeout_add_seconds and friends ignore g_main_context_push_thread_default).
pywebview's GTK backend already runs that global default context on the
main thread, so callbacks land there; this module just needs to block the
calling (worker) thread until they do, via a plain threading.Event.
"""
import base64
import threading

import dbus
import dbus.mainloop.glib
import gi
gi.require_version("Gst", "1.0")
from gi.repository import Gst

from .input_controller import debug_linux

_PORTAL_BUS_NAME = "org.freedesktop.portal.Desktop"
_PORTAL_PATH = "/org/freedesktop/portal/desktop"

SOURCE_TYPE_MONITOR = 1
CURSOR_MODE_EMBEDDED = 2
PERSIST_MODE_PERSISTENT = 2

_setup_lock = threading.Lock()
_glib_loop_ready = False

_token_lock = threading.Lock()
_token_counter = [0]


def _new_token(prefix):
    with _token_lock:
        _token_counter[0] += 1
        return f"{prefix}{_token_counter[0]}"


def _ensure_glib_mainloop():
    """Make dbus-python deliver signals via GLib's global default context.

    Must run before the first dbus.SessionBus() in the process. Safe to call
    more than once. If nothing else in the process pumps that context (no
    GTK window yet, e.g. under a bare test script), spin up a background
    thread that does — pywebview normally makes this redundant.
    """
    global _glib_loop_ready
    with _setup_lock:
        if _glib_loop_ready:
            return
        dbus.mainloop.glib.DBusGMainLoop(set_as_default=True)
        from gi.repository import GLib

        # Harmless if pywebview's GTK loop already pumps the default context:
        # GLib's context locking lets multiple threads cooperate over it.
        def pump():
            GLib.MainLoop().run()

        threading.Thread(target=pump, daemon=True, name="glib-mainloop-fallback").start()
        _glib_loop_ready = True


class _PendingCall:
    """Bridges a dbus-python async Response signal to a blocking wait."""

    def __init__(self, bus):
        self.bus = bus
        self.event = threading.Event()
        self.response = None
        self.results = None

    def _on_response(self, response, results):
        self.response = int(response)
        self.results = dict(results) if results is not None else {}
        self.event.set()

    def wait(self, request_path, timeout):
        sig = self.bus.add_signal_receiver(
            self._on_response, signal_name="Response",
            dbus_interface="org.freedesktop.portal.Request", path=request_path,
        )
        try:
            if not self.event.wait(timeout):
                return -1, {}
            return self.response, self.results
        finally:
            sig.remove()


def _call_and_wait(bus, method, options, *lead_args, timeout=120):
    """Call a portal method (lead_args..., options) -> Request path, then
    block until its Response signal fires. Returns (response_code, results)."""
    options = dict(options)
    options["handle_token"] = _new_token("t2p")
    pending = _PendingCall(bus)
    request_path = method(*(lead_args + (options,)))
    return pending.wait(request_path, timeout)


def _grab_one_frame(fd, node_id, timeout_s=15):
    """Blocking frame pull; uses Gst's own blocking APIs, no main-loop needed."""
    Gst.init(None)
    pipeline = Gst.parse_launch(
        f"pipewiresrc fd={fd} path={node_id} ! "
        "videoconvert ! pngenc ! appsink name=sink sync=false max-buffers=1"
    )
    sink = pipeline.get_by_name("sink")
    gbus = pipeline.get_bus()
    pipeline.set_state(Gst.State.PLAYING)

    try:
        remaining_ms = timeout_s * 1000
        while remaining_ms > 0:
            msg = gbus.timed_pop_filtered(
                50 * Gst.MSECOND, Gst.MessageType.ERROR | Gst.MessageType.EOS)
            if msg is not None:
                if msg.type == Gst.MessageType.ERROR:
                    err, dbg = msg.parse_error()
                    debug_linux(f"portal capture gst error: {err} | {dbg}")
                return None
            sample = sink.emit("try-pull-sample", 0)
            if sample is not None:
                buf = sample.get_buffer()
                ok, mapinfo = buf.map(Gst.MapFlags.READ)
                data = bytes(mapinfo.data)
                buf.unmap(mapinfo)
                return data
            remaining_ms -= 50
        return None
    finally:
        pipeline.set_state(Gst.State.NULL)


def capture_screen(restore_token=None):
    """Grab one screenshot of a user-picked monitor via the ScreenCast portal.

    Returns a dict:
      ok=True: {ok, image_b64 (PNG), width, height, offset_x, offset_y, restore_token}
      ok=False: {ok, error, message}
    `offset_x/offset_y` place the captured monitor within the virtual desktop
    (the same space uinput's ABS_X/ABS_Y already use), for multi-monitor setups.
    """
    try:
        _ensure_glib_mainloop()
        bus = dbus.SessionBus()
        portal = bus.get_object(_PORTAL_BUS_NAME, _PORTAL_PATH)
        screencast = dbus.Interface(portal, "org.freedesktop.portal.ScreenCast")

        resp, results = _call_and_wait(
            bus, screencast.CreateSession, {"session_handle_token": _new_token("session")})
        if resp != 0:
            return {"ok": False, "error": "cancelled" if resp == 1 else "portal_error",
                    "message": "Could not create a screen-sharing session."}
        session_handle = results["session_handle"]

        select_options = {
            "types": dbus.UInt32(SOURCE_TYPE_MONITOR),
            "multiple": False,
            "cursor_mode": dbus.UInt32(CURSOR_MODE_EMBEDDED),
            "persist_mode": dbus.UInt32(PERSIST_MODE_PERSISTENT),
        }
        if restore_token:
            select_options["restore_token"] = restore_token

        resp, results = _call_and_wait(
            bus, screencast.SelectSources, select_options, session_handle, timeout=180)
        if resp != 0:
            return {"ok": False, "error": "cancelled" if resp == 1 else "portal_error",
                    "message": "Screen selection was cancelled."}

        resp, results = _call_and_wait(
            bus, screencast.Start, {}, session_handle, "", timeout=180)
        if resp != 0:
            return {"ok": False, "error": "cancelled" if resp == 1 else "portal_error",
                    "message": "Screen sharing was cancelled."}

        streams = results.get("streams")
        if not streams:
            return {"ok": False, "error": "no_stream", "message": "No screen stream was offered."}
        node_id, stream_props = streams[0]
        stream_props = dict(stream_props)
        position = stream_props.get("position", (0, 0))
        size = stream_props.get("size", (0, 0))
        new_restore_token = results.get("restore_token")

        fd_obj = screencast.OpenPipeWireRemote(session_handle, {})
        fd = fd_obj.take()

        data = _grab_one_frame(fd, int(node_id))
        if data is None:
            return {"ok": False, "error": "capture_failed",
                    "message": "Could not read a frame from the screen stream."}

        return {
            "ok": True,
            "image_b64": base64.b64encode(data).decode("ascii"),
            "width": int(size[0]),
            "height": int(size[1]),
            "offset_x": int(position[0]),
            "offset_y": int(position[1]),
            "restore_token": str(new_restore_token) if new_restore_token else None,
        }
    except Exception as e:
        debug_linux(f"portal capture failed: {e}")
        return {"ok": False, "error": "exception", "message": str(e)}
