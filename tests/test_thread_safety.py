import os
import threading
import unittest

import auto_copy_gui


class DummyRoot:
    def __init__(self):
        self.after_calls = []

    def after(self, delay, callback):
        self.after_calls.append((delay, callback))
        # Simulate Tk calling the callback immediately
        callback()


class DummyText:
    def __init__(self):
        self.state = None
        self.inserted = []
        self.seen = None

    def configure(self, **kwargs):
        state = kwargs.get("state")
        if state is not None:
            self.state = state

    def insert(self, index, value):
        self.inserted.append((index, value))

    def see(self, index):
        self.seen = index


class ThreadSafetyTests(unittest.TestCase):
    def setUp(self):
        self.app = auto_copy_gui.AutoCopyApp.__new__(auto_copy_gui.AutoCopyApp)
        self.app.root = DummyRoot()
        self.app.log_text = DummyText()

    def test_call_in_ui_thread_from_main_thread(self):
        calls = []

        def mark():
            calls.append("main")

        self.app.call_in_ui_thread(mark)

        self.assertEqual(calls, ["main"])
        self.assertEqual(self.app.root.after_calls, [])

    def test_call_in_ui_thread_from_worker_thread(self):
        calls = []

        def mark():
            calls.append("worker")

        worker = threading.Thread(target=lambda: self.app.call_in_ui_thread(mark))
        worker.start()
        worker.join()

        self.assertEqual(calls, ["worker"])
        self.assertEqual(len(self.app.root.after_calls), 1)
        delay, _ = self.app.root.after_calls[0]
        self.assertEqual(delay, 0)

    def test_append_log_uses_line_separator(self):
        self.app._append_log("hello")

        self.assertEqual(
            self.app.log_text.inserted,
            [
                (auto_copy_gui.tk.END, "hello"),
                (auto_copy_gui.tk.END, os.linesep),
            ],
        )
        self.assertEqual(self.app.log_text.state, auto_copy_gui.tk.DISABLED)
        self.assertEqual(self.app.log_text.seen, auto_copy_gui.tk.END)


if __name__ == "__main__":
    unittest.main()
