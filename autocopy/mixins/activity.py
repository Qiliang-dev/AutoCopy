import tkinter as tk
import tkinter.font as tkfont
import win32api


class ActivityMixin:
    """Reminder dialog logic (repurposed from the old inactivity checker)."""

    def start_activity_monitoring(self):
        """Show and keep updating the reminder dialog when enabled."""
        if not getattr(self, "reminder_enabled_var", None) or not self.reminder_enabled_var.get():
            self.stop_activity_monitoring()
            return

        self.activity_monitor_active = True
        self.show_reminder_dialog()
        self._schedule_reminder_refresh()

    def stop_activity_monitoring(self):
        """Stop reminder refresh loop and close dialog if reminder is off."""
        self.activity_monitor_active = False
        # Cancel periodic refresh
        try:
            if getattr(self, "_reminder_update_job", None) and self.reminder_dialog and self.reminder_dialog.winfo_exists():
                self.reminder_dialog.after_cancel(self._reminder_update_job)
        except Exception:
            pass
        self._reminder_update_job = None

        # Cancel flashing if any (legacy)
        if getattr(self, "_reminder_flash_job", None) and self.reminder_dialog:
            try:
                self.reminder_dialog.after_cancel(self._reminder_flash_job)
            except Exception:
                pass
        self._reminder_flash_job = None

        # Destroy dialog only if reminder is no longer enabled
        if not (self.reminder_enabled_var and self.reminder_enabled_var.get()):
            if self.reminder_dialog and self.reminder_dialog.winfo_exists():
                try:
                    self.reminder_dialog.destroy()
                except Exception:
                    pass
                self.reminder_dialog = None

    def _schedule_reminder_refresh(self):
        """Kick off refresh loop for reminder content."""
        if not self.activity_monitor_active:
            return
        if not self.reminder_enabled_var or not self.reminder_enabled_var.get():
            self.stop_activity_monitoring()
            return

        self._refresh_reminder_content()
        if self.reminder_dialog and self.reminder_dialog.winfo_exists():
            try:
                self._reminder_update_job = self.reminder_dialog.after(800, self._schedule_reminder_refresh)
            except Exception:
                pass

    def _refresh_reminder_content(self):
        """Render current column history into the reminder dialog."""
        if not (self.reminder_dialog and self.reminder_dialog.winfo_exists()):
            return

        content_text = ""
        try:
            summary = self.get_preceding_cells_summary()
            content_text = summary
        except Exception as exc:
            content_text = f"Failed to read Excel data: {exc}"
            self.log(content_text)

        widget = getattr(self, "reminder_content_widget", None)
        if widget and widget.winfo_exists():
            widget.configure(state=tk.NORMAL)
            widget.delete(1.0, tk.END)
            widget.insert(tk.END, content_text)
            widget.configure(state=tk.DISABLED)

    def show_reminder_dialog(self):
        """Create or lift the reminder dialog."""
        if not self.reminder_enabled_var or not self.reminder_enabled_var.get():
            return
        if self.reminder_dialog and self.reminder_dialog.winfo_exists():
            try:
                self.reminder_dialog.lift()
            except Exception:
                pass
            return

        self.reminder_dialog = tk.Toplevel(self.root)
        self.reminder_dialog.title("Reminder - Excel Context")
        self.reminder_dialog.attributes('-topmost', True)
        self.reminder_dialog.resizable(True, True)

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
        window_width = 900
        window_height = 600
        x = target_monitor['left'] + (target_monitor['width'] - window_width) // 2
        y = target_monitor['top'] + (target_monitor['height'] - window_height) // 2
        self.reminder_dialog.geometry(f"{window_width}x{window_height}+{x}+{y}")

        container = tk.Frame(self.reminder_dialog, bg="#1b1f38")
        container.pack(fill=tk.BOTH, expand=True)

        # Shared fonts that will scale with window resize
        self._reminder_fonts = {
            "header": tkfont.Font(family="Segoe UI", size=14, weight="bold"),
            "info": tkfont.Font(family="Segoe UI", size=10),
            "body": tkfont.Font(family="Consolas", size=10),
        }

        header = tk.Label(
            container,
            text="当前选中单元格的前置内容",
            font=self._reminder_fonts["header"],
            fg="#e4e7fb",
            bg="#1b1f38",
            pady=12
        )
        header.pack(fill=tk.X)

        info_label = tk.Label(
            container,
            text="实时读取选中列，向上展示到当前单元格之前的所有行（含合并单元格）。",
            font=self._reminder_fonts["info"],
            fg="#9aa2d4",
            bg="#1b1f38"
        )
        info_label.pack(fill=tk.X, padx=12)

        body = tk.Frame(container, bg="#1b1f38", padx=12, pady=12)
        body.pack(fill=tk.BOTH, expand=True)

        text_widget = tk.Text(
            body,
            height=20,
            wrap=tk.WORD,
            font=self._reminder_fonts["body"],
            bg="#0f1224",
            fg="#e4e7fb",
            insertbackground="#e4e7fb"
        )
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        text_widget.configure(state=tk.DISABLED)
        self.reminder_content_widget = text_widget

        scrollbar = tk.Scrollbar(body, command=text_widget.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget.configure(yscrollcommand=scrollbar.set)

        button_bar = tk.Frame(container, bg="#1b1f38")
        button_bar.pack(fill=tk.X, pady=(0, 10))

        close_btn = tk.Button(
            button_bar,
            text="关闭提醒",
            command=lambda: self._disable_reminder_from_dialog(),
            font=("Segoe UI", 10, "bold"),
            bg="#f26b38",
            fg="white",
            padx=12,
            pady=6,
            relief=tk.FLAT
        )
        close_btn.pack(side=tk.RIGHT, padx=12)

        self._refresh_reminder_content()
        self._start_reminder_flash(text_widget)
        # Bind resize to scale fonts
        self.reminder_dialog.bind("<Configure>", self._on_reminder_resize)

    def _disable_reminder_from_dialog(self):
        """Turn off reminder via dialog button."""
        try:
            if self.reminder_enabled_var:
                self.reminder_enabled_var.set(False)
        except Exception:
            pass
        self.stop_activity_monitoring()
        if self.reminder_dialog and self.reminder_dialog.winfo_exists():
            try:
                self.reminder_dialog.destroy()
            except Exception:
                pass
            self.reminder_dialog = None

    def _start_reminder_flash(self, widget: tk.Text):
        """Flash text color to make it noticeable."""
        try:
            if self._reminder_flash_job and self.reminder_dialog and self.reminder_dialog.winfo_exists():
                self.reminder_dialog.after_cancel(self._reminder_flash_job)
        except Exception:
            pass

        colors = ["#e4e7fb", "#ffc857"]
        state = {"idx": 0}

        def _toggle():
            if not (self.reminder_dialog and self.reminder_dialog.winfo_exists()):
                return
            try:
                state["idx"] = 1 - state["idx"]
                widget.configure(fg=colors[state["idx"]])
            except Exception:
                pass
            try:
                self._reminder_flash_job = self.reminder_dialog.after(650, _toggle)
            except Exception:
                pass

        _toggle()

    def _on_reminder_resize(self, event):
        """Adapt font sizes based on window height for better readability."""
        if not hasattr(self, "_reminder_fonts"):
            return
        try:
            h = max(event.height, 200)
            header_size = max(12, min(28, h // 24))
            info_size = max(10, min(20, h // 36))
            body_size = max(10, min(22, h // 32))
            self._reminder_fonts["header"].configure(size=header_size)
            self._reminder_fonts["info"].configure(size=info_size)
            self._reminder_fonts["body"].configure(size=body_size)
        except Exception:
            pass
