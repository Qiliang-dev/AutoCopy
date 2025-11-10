import ctypes
import time

import tkinter as tk
from ctypes import windll
from ctypes import wintypes

import win32api


class ActivityMixin:
    """Encapsulate activity monitoring and reminder dialog logic."""

    def start_activity_monitoring(self):
        """开始监控用户活动"""
        if not getattr(self, "reminder_enabled_var", None) or not self.reminder_enabled_var.get():
            self.stop_activity_monitoring()
            return

        self.activity_monitor_active = True
        self.last_activity_time = time.time()
        self.activity_detected = False

        cursor = wintypes.POINT()
        windll.user32.GetCursorPos(ctypes.byref(cursor))
        self.last_mouse_pos = (cursor.x, cursor.y)

        self._check_activity()

    def _check_activity(self):
        """检查活动状态"""
        if not self.activity_monitor_active:
            return
        if not self.reminder_enabled_var or not self.reminder_enabled_var.get():
            self.stop_activity_monitoring()
            return

        try:
            cursor = wintypes.POINT()
            windll.user32.GetCursorPos(ctypes.byref(cursor))
            current_pos = (cursor.x, cursor.y)

            if current_pos != self.last_mouse_pos:
                self.log("Mouse movement detected")
                self.activity_detected = True
                self.last_mouse_pos = current_pos

            for key in range(0x30, 0x5A):
                if windll.user32.GetAsyncKeyState(key) & 0x8000:
                    self.log("Keyboard activity detected")
                    self.activity_detected = True
                    break

            if self.activity_detected:
                self.log("Activity detected - stopping monitoring")
                self.stop_activity_monitoring()
                if self.reminder_dialog and self.reminder_dialog.winfo_exists():
                    self.reminder_dialog.destroy()
                    self.reminder_dialog = None
                return

            current_time = time.time()
            time_since_last_activity = current_time - self.last_activity_time

            try:
                reminder_time = int(self.reminder_time_var.get())
            except (ValueError, AttributeError):
                reminder_time = self.reminder_time

            if time_since_last_activity >= reminder_time:
                self.log("No activity detected for specified time - showing reminder")
                self.show_reminder_dialog()
                self.stop_activity_monitoring()
            else:
                self.root.after(100, self._check_activity)

        except Exception as e:
            self.log(f"Activity check error: {str(e)}")
            self.root.after(100, self._check_activity)

    def stop_activity_monitoring(self):
        """停止活动监控"""
        self.activity_monitor_active = False
        self.activity_detected = False
        if self.reminder_timer:
            self.root.after_cancel(self.reminder_timer)
            self.reminder_timer = None
        if self._reminder_flash_job and self.reminder_dialog:
            try:
                self.reminder_dialog.after_cancel(self._reminder_flash_job)
            except Exception:
                pass
        self._reminder_flash_job = None

    def show_reminder_dialog(self):
        if not self.reminder_enabled_var or not self.reminder_enabled_var.get():
            return
        if self.reminder_dialog and self.reminder_dialog.winfo_exists():
            return

        self.reminder_dialog = tk.Toplevel(self.root)
        self.reminder_dialog.title("Activity Reminder")
        self.reminder_dialog.attributes('-topmost', True)

        monitors = []
        try:
            def callback(monitor, dc, rect, data):
                monitors.append({
                    'left': rect.contents.left,
                    'top': rect.contents.top,
                    'right': rect.contents.right,
                    'bottom': rect.contents.bottom,
                    'width': rect.contents.right - rect.contents.left,
                    'height': rect.contents.bottom - rect.contents.top
                })
                return 1
            win32api.EnumDisplayMonitors(None, None, callback, 0)
        except Exception:
            monitors = [{
                'left': 0,
                'top': 0,
                'width': self.root.winfo_screenwidth(),
                'height': self.root.winfo_screenheight()
            }]

        target_monitor = monitors[0]

        window_width = 1200
        window_height = 800
        x = target_monitor['left'] + (target_monitor['width'] - window_width) // 2
        y = target_monitor['top'] + (target_monitor['height'] - window_height) // 2
        self.reminder_dialog.geometry(f"{window_width}x{window_height}+{x}+{y}")

        self._reminder_bg_colors = ["#FFE4E1", "#FF0000"]
        self._reminder_bg_index = 0
        self._reminder_flash_job = None

        self.reminder_dialog.configure(bg=self._reminder_bg_colors[self._reminder_bg_index])

        frame = tk.Frame(self.reminder_dialog, bg=self._reminder_bg_colors[self._reminder_bg_index], padx=30, pady=30)
        frame.pack(fill=tk.BOTH, expand=True)

        warning_label = tk.Label(frame, text="??", font=("Arial", 72), bg=self._reminder_bg_colors[self._reminder_bg_index])
        warning_label.pack(pady=(0, 20))

        message = "NO ACTIVITY DETECTED!\n\nPlease continue your work or close this window."
        text_label = tk.Label(
            frame,
            text=message,
            font=("Arial", 16, "bold"),
            justify=tk.CENTER,
            bg=self._reminder_bg_colors[self._reminder_bg_index],
            fg="#8B0000"
        )
        text_label.pack(pady=20)

        close_button = tk.Button(
            frame,
            text="CLOSE",
            command=self.reminder_dialog.destroy,
            font=("Arial", 12, "bold"),
            bg="#FF6B6B",
            fg="white",
            relief=tk.RAISED,
            padx=20,
            pady=10
        )
        close_button.pack(pady=20)

        self.reminder_dialog.bind("<Motion>", lambda e: self.reminder_dialog.destroy())
        self.reminder_dialog.bind("<Key>", lambda e: self.reminder_dialog.destroy())
        self.reminder_dialog.bind("<Button>", lambda e: self.reminder_dialog.destroy())

        self._reminder_flash_bg()

    def _reminder_flash_bg(self):
        if not self.reminder_enabled_var or not self.reminder_enabled_var.get():
            return
        if not (self.reminder_dialog and self.reminder_dialog.winfo_exists()):
            return
        self._reminder_bg_index = 1 - self._reminder_bg_index
        color = self._reminder_bg_colors[self._reminder_bg_index]
        self.reminder_dialog.configure(bg=color)
        for child in self.reminder_dialog.winfo_children():
            try:
                child.configure(bg=color)
                for subchild in child.winfo_children():
                    subchild.configure(bg=color)
            except Exception:
                pass
        self._reminder_flash_job = self.reminder_dialog.after(500, self._reminder_flash_bg)
