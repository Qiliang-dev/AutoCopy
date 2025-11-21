import sys
import threading
import types
import unittest


def ensure_stub(name, value=None):
    """Register a lightweight stub module if missing."""
    if name in sys.modules:
        return sys.modules[name]
    if value is None:
        value = types.SimpleNamespace()
    sys.modules[name] = value
    return value


# Provide stubs for optional native deps during tests
ensure_stub("pyperclip", types.SimpleNamespace(copy=lambda *a, **k: None, paste=lambda: ""))
ensure_stub("pyautogui", types.SimpleNamespace())
ensure_stub("win32api", types.SimpleNamespace())
ensure_stub("win32con", types.SimpleNamespace())
ensure_stub("win32com", types.SimpleNamespace())
ensure_stub("win32com.client", types.SimpleNamespace(Dispatch=lambda *a, **k: None, GetActiveObject=lambda *a, **k: None))
ensure_stub("pythoncom", types.SimpleNamespace(CoInitialize=lambda *a, **k: None))
ensure_stub("winsound", types.SimpleNamespace(Beep=lambda *a, **k: None))

import autocopy_main


class DummyRoot:
    def __init__(self):
        self.after_calls = []

    def after(self, delay, callback):
        # Record scheduling; do not invoke automatically to mimic Tk main loop behavior
        self.after_calls.append((delay, callback))
        return f"job-{len(self.after_calls)}"


class DummyText:
    def __init__(self):
        self.state = None
        self.inserted = []
        self.seen = None

    def configure(self, **kwargs):
        if "state" in kwargs:
            self.state = kwargs.get("state")

    def insert(self, index, value):
        self.inserted.append((index, value))

    def see(self, index):
        self.seen = index

    def winfo_exists(self):
        return True

    def cget(self, key):
        if key == "state":
            return self.state
        return None


class ThreadSafetyTests(unittest.TestCase):
    def setUp(self):
        self.app = autocopy_main.AutoCopyApp.__new__(autocopy_main.AutoCopyApp)
        self.app.root = DummyRoot()
        self.app.log_text = DummyText()

    def test_run_on_ui_thread_from_main_thread_executes_immediately(self):
        calls = []

        def mark():
            calls.append("main")

        # monkeypatch _execute_ui_task to record execution
        self.app._execute_ui_task = lambda task: (calls.append("exec"), task())
        self.app._run_on_ui_thread(mark)

        self.assertEqual(calls, ["exec", "main"])
        self.assertEqual(self.app.root.after_calls, [])

    def test_run_on_ui_thread_from_worker_schedules(self):
        calls = []

        def mark():
            calls.append("worker")

        self.app._execute_ui_task = lambda task: (calls.append("exec"), task())

        worker = threading.Thread(target=lambda: self.app._run_on_ui_thread(mark))
        worker.start()
        worker.join()

        # No direct execution on worker; only scheduling recorded
        self.assertEqual(calls, [])
        self.assertEqual(len(self.app.root.after_calls), 1)
        delay, callback = self.app.root.after_calls[0]
        self.assertEqual(delay, 0)

        # Simulate Tk invoking the scheduled callback
        callback()
        self.assertEqual(calls, ["exec", "worker"])

    def test_log_appends_text_and_newline(self):
        # Ensure logger object is absent to avoid file writes
        self.app.logger = None
        self.app.log("hello world")

        # two inserts: message + newline, with widget enabled/disabled around write
        self.assertEqual(len(self.app.log_text.inserted), 1)
        location, text = self.app.log_text.inserted[0]
        self.assertEqual(location, autocopy_main.tk.END)
        self.assertTrue(text.endswith("hello world\n"))
        self.assertIsNone(self.app.log_text.state)
        self.assertEqual(self.app.log_text.seen, autocopy_main.tk.END)


if __name__ == "__main__":
    unittest.main()
