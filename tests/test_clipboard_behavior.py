import queue
import re
import threading
import time
import unittest

import sys
import types


def ensure_stub(name, value=None):
    """Register a lightweight stub module if missing."""
    if name in sys.modules:
        return sys.modules[name]
    if value is None:
        value = types.SimpleNamespace()
    sys.modules[name] = value
    return value


# Provide stubs for optional native deps during tests
pyperclip = ensure_stub("pyperclip", types.SimpleNamespace(copy=lambda *a, **k: None, paste=lambda: ""))
ensure_stub("pyautogui", types.SimpleNamespace())
ensure_stub("win32api", types.SimpleNamespace())
ensure_stub("win32con", types.SimpleNamespace())
win32com = ensure_stub("win32com", types.SimpleNamespace())
ensure_stub("win32com.client", types.SimpleNamespace(Dispatch=lambda *a, **k: None, GetActiveObject=lambda *a, **k: None))
ensure_stub("pythoncom", types.SimpleNamespace(CoInitialize=lambda *a, **k: None))
ensure_stub("winsound", types.SimpleNamespace(Beep=lambda *a, **k: None))

import autocopy.mixins.clipboard as clip


class DummyVar:
    def __init__(self, value):
        self.value = value

    def get(self):
        return self.value


class DummyText:
    def __init__(self):
        self._state = None
        self.deleted = []
        self.inserted = []

    def winfo_exists(self):
        return True

    def cget(self, key):
        if key == "state":
            return self._state
        return None

    def configure(self, **kwargs):  # matches Tk API shape
        if "state" in kwargs:
            self._state = kwargs["state"]

    def delete(self, start, end):
        self.deleted.append((start, end))

    def insert(self, index, value):
        self.inserted.append((index, value))


class DummyLabel:
    def __init__(self):
        self.config_calls = []

    def config(self, **kwargs):
        self.config_calls.append(kwargs)


class DummyButton:
    def __init__(self):
        self.config_calls = []

    def config(self, **kwargs):
        self.config_calls.append(kwargs)


class DummyConfirmation:
    def __init__(self):
        self.destroy_called = False

    def winfo_exists(self):
        return True

    def destroy(self):
        self.destroy_called = True


class DummyRoot:
    def __init__(self):
        self.after_calls = []
        self.after_cancel_calls = []
        self._job_counter = 0

    def after(self, delay, callback):
        # Simulate Tk's after by returning a job id without invoking callback
        self._job_counter += 1
        job_id = f"job{self._job_counter}"
        self.after_calls.append((delay, callback, job_id))
        return job_id

    def after_cancel(self, job_id):
        self.after_cancel_calls.append(job_id)


class DummyThread:
    def __init__(self, alive=True):
        self._alive = alive
        self.join_called = False

    def is_alive(self):
        return self._alive

    def join(self, timeout=None):
        self.join_called = True


class DummyClipboardApp(clip.ClipboardMixin):
    def __init__(self):
        # minimal attributes the mixin expects
        self.running = False
        self.clipboard_lock = threading.Lock()
        self.clipboard_content = ""
        self.duplicate_time_var = DummyVar("3")
        self.last_pasted_content = ""
        self.last_paste_time = 0
        self.ignore_initial_clipboard = False
        self.initial_clipboard_snapshot = ""
        self.message_queue = queue.Queue()
        self.clipboard_text = DummyText()
        self.clipboard_display_error_count = 0
        self.last_clipboard_display_error_time = 0
        self.clipboard_update_job = None
        self.root = DummyRoot()
        self.status_label = DummyLabel()
        self.start_button = DummyButton()
        self.stop_button = DummyButton()
        self.set_excel_button = DummyButton()
        self.confirmation_dialog = None
        self.monitor_thread = None
        self.logs = []

    def is_valid_format(self, text):
        # simple regex check similar to main app
        pattern = getattr(self, "format_pattern", r"^20\d{2}_\d{2}_\d{2}_\d{6}$")
        return bool(re.match(pattern, text or ""))

    def log(self, message, level=None):
        self.logs.append((level, message))


class DummyPyperclip:
    def __init__(self, content):
        self.content = content

    def paste(self):
        return self.content


class ClipboardBehaviorTests(unittest.TestCase):
    def setUp(self):
        self.original_pyperclip = clip.pyperclip

    def tearDown(self):
        clip.pyperclip = self.original_pyperclip

    def test_no_auto_paste_when_not_running(self):
        app = DummyClipboardApp()
        app.running = False
        app.format_pattern = r"^OK$"
        clip.pyperclip = DummyPyperclip("OK")

        app.update_clipboard_display()

        self.assertTrue(app.message_queue.empty(), "Should not enqueue paste when not running")
        self.assertIsNone(app.clipboard_update_job, "Should not schedule next call when stopped")

    def test_auto_paste_enqueued_when_running_and_valid(self):
        app = DummyClipboardApp()
        app.running = True
        app.format_pattern = r"^OK$"
        app.last_paste_time = time.time() - 10  # avoid duplicate guard
        clip.pyperclip = DummyPyperclip("OK")

        app.update_clipboard_display()

        self.assertFalse(app.message_queue.empty(), "Paste should be enqueued when running and valid")
        msg = app.message_queue.get_nowait()
        self.assertEqual(msg["type"], "paste_content")
        self.assertEqual(msg["content"], "OK")
        self.assertIsNotNone(app.clipboard_update_job, "Should schedule next polling callback when running")

    def test_stop_monitoring_cancels_callbacks_and_thread(self):
        app = DummyClipboardApp()
        app.running = True
        app.clipboard_update_job = "job123"
        app.monitor_thread = DummyThread(alive=True)
        dialog = DummyConfirmation()
        app.confirmation_dialog = dialog

        app.stop_monitoring()

        self.assertFalse(app.running, "Running flag should be cleared")
        self.assertIn("job123", app.root.after_cancel_calls, "Scheduled callback should be cancelled")
        self.assertIsNone(app.clipboard_update_job, "Callback id should be cleared")
        self.assertIsNone(app.monitor_thread, "Monitor thread reference should be cleared")
        self.assertTrue(dialog.destroy_called, "Notification dialog should be closed")
        self.assertIsNone(app.confirmation_dialog, "Dialog reference should be cleared")


if __name__ == "__main__":
    unittest.main()
