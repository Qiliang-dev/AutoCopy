import logging
import threading
import time
import traceback
import queue

import pyperclip
import tkinter as tk
from ctypes import windll
from tkinter import messagebox


class ClipboardMixin:
    """Encapsulate clipboard monitoring, auto-paste, and notification logic.
    
    THREAD SAFETY: This mixin uses a message queue to ensure all Excel operations
    happen on the main thread, avoiding COM threading issues.
    """

    def _get_clipboard_sequence_number(self):
        """Return Windows clipboard sequence number to track duplicate copy events."""
        try:
            return windll.user32.GetClipboardSequenceNumber()
        except Exception:
            return None

    def _init_message_queue(self):
        """Initialize the thread-safe message queue for Excel operations."""
        self.message_queue = queue.Queue()
        self._start_queue_processor()

    def _start_queue_processor(self):
        """Start processing messages from the queue on the main thread."""
        try:
            # Process all pending messages
            messages_processed = 0
            max_batch = 10  # Prevent blocking UI for too long
            
            while not self.message_queue.empty() and messages_processed < max_batch:
                try:
                    message = self.message_queue.get_nowait()
                    self._process_message(message)
                    messages_processed += 1
                except queue.Empty:
                    break
                except Exception as e:
                    self.log(f"Error processing message: {e}", level=logging.ERROR)
                    traceback.print_exc()
        except Exception as e:
            self.log(f"Queue processor error: {e}", level=logging.ERROR)
        finally:
            # Schedule next check
            if hasattr(self, 'root') and self.root:
                try:
                    self.root.after(100, self._start_queue_processor)
                except Exception:
                    pass

    def _process_message(self, message):
        """Process a single message from the queue (runs on main thread).
        
        Args:
            message: dict with 'type' and optional parameters
        """
        msg_type = message.get('type')
        
        try:
            # ✅ CRITICAL FIX: Set flag to prevent conflicts with schedule_cell_check
            if msg_type in ('paste_content', 'refresh_cell', 'move_next_row'):
                self._processing_excel_operation = True
            
            try:
                if msg_type == 'paste_content':
                    content = message.get('content', '')
                    self._handle_paste_request(content)
                    
                elif msg_type == 'refresh_cell':
                    self._handle_refresh_cell_request()
                    
                elif msg_type == 'update_clipboard_display':
                    force = message.get('force', False)
                    self._handle_update_clipboard_display(force)
                    
                elif msg_type == 'move_next_row':
                    skip_rows = message.get('skip_rows', 1)
                    self._handle_move_next_row(skip_rows)
                    
                else:
                    self.log(f"Unknown message type: {msg_type}", level=logging.WARNING)
            finally:
                # ✅ Always clear the flag
                if msg_type in ('paste_content', 'refresh_cell', 'move_next_row'):
                    self._processing_excel_operation = False
                
        except Exception as e:
            self.log(f"Error handling message {msg_type}: {e}", level=logging.ERROR)
            traceback.print_exc()
            self._processing_excel_operation = False  # Ensure flag is cleared

    def _handle_paste_request(self, content):
        """Handle paste request on main thread."""
        try:
            success = self.paste_to_excel(show_error_dialog=False)
            
            if success:
                self.show_success_notification(content)
            else:
                error_msg = "Paste failed. Please check Excel connection."
                self.log(error_msg)
                self.show_error_notification(error_msg)
            
            self.last_activity_time = time.time()
            self.start_activity_monitoring()
            
        except Exception as e:
            self.log(f"Paste request error: {e}", level=logging.ERROR)
            self.show_error_notification(f"Paste error: {e}")

    def _handle_refresh_cell_request(self):
        """Handle cell refresh request on main thread."""
        try:
            if self.excel_app:
                cell_address = self.excel_app.ActiveCell.Address
                if cell_address != self.current_cell:
                    self.current_cell = cell_address
                    self.cell_label.config(text=cell_address)
                    if self.running:
                        self.log(f"Cell selection changed: {cell_address}")
        except Exception as e:
            # Silently ignore Excel access errors during monitoring
            pass

    def _handle_update_clipboard_display(self, force=False):
        """Handle clipboard display update on main thread."""
        self.update_clipboard_display(force=force)

    def _handle_move_next_row(self, skip_rows):
        """Handle move to next row on main thread."""
        try:
            if self.excel_app:
                current_cell = self.excel_app.ActiveCell
                current_row = current_cell.Row
                current_column = current_cell.Column
                
                next_cell = self.excel_app.ActiveSheet.Cells(current_row + skip_rows, current_column)
                next_cell.Select()
                
                # Refresh cell info
                self._handle_refresh_cell_request()
                
                self.log(f"Automatically moved {skip_rows} rows down: {next_cell.Address}")
        except Exception as e:
            self.log(f"Error moving to next row: {e}")

    def update_clipboard_display(self, force: bool = False):
        """更新剪贴板显示 (runs on main thread)."""
        try:
            # ✅ CRITICAL FIX: Use lock to protect pyperclip access (not thread-safe!)
            with self.clipboard_lock:
                content = pyperclip.paste()

            # 重置错误计数（成功读取剪贴板）
            self.clipboard_display_error_count = 0

            # 检查内容是否有效以及是否发生变化
            content_has_value = content and content.strip() != ""
            content_changed = content != self.clipboard_content

            # 内容发生变化时才更新 UI，避免重复刷屏
            if content_has_value and (content_changed or force):
                # 记录更新时间
                current_time = time.time()

                # 判定是否重复内容
                is_duplicate = False
                if hasattr(self, "last_pasted_content") and hasattr(self, "last_paste_time"):
                    try:
                        duplicate_threshold = float(self.duplicate_time_var.get())
                    except (ValueError, AttributeError):
                        duplicate_threshold = 3.0  # 默认阈值

                    time_diff = current_time - self.last_paste_time
                    content_same = content == self.last_pasted_content

                    if content_same and time_diff < duplicate_threshold:
                        is_duplicate = True
                        if not getattr(self, "_duplicate_logged", False):
                            self.log(f"Duplicate content detected (within {duplicate_threshold}s), will skip until threshold passes")
                            self._duplicate_logged = True
                    elif content_same and time_diff >= duplicate_threshold:
                        is_duplicate = False
                        self._duplicate_logged = False
                    else:
                        self._duplicate_logged = False

                # 更新缓存
                self.clipboard_content = content

                # 更新 UI
                clip_widget = getattr(self, "clipboard_text", None)
                if clip_widget and clip_widget.winfo_exists():
                    previous_state = clip_widget.cget("state")
                    clip_widget.configure(state=tk.NORMAL)
                    clip_widget.delete(1.0, tk.END)

                    display_content = content[:500] + "... (content truncated)" if len(content) > 500 else content
                    clip_widget.insert(tk.END, display_content)
                    clip_widget.configure(state=previous_state)

                # 校验格式
                match_result = self.is_valid_format(content)
                if match_result:
                    skip_due_to_initial = False
                    if getattr(self, "ignore_initial_clipboard", False):
                        initial_snapshot = getattr(self, "initial_clipboard_snapshot", "")
                        if initial_snapshot and content == initial_snapshot:
                            skip_due_to_initial = True
                            self.log("Initial clipboard content ignored; waiting for the next clipboard change")

                    if not skip_due_to_initial and not is_duplicate:
                        # 有效内容，发送粘贴请求到队列（不直接操作Excel）
                        self.message_queue.put({'type': 'paste_content', 'content': content})
                elif self.running and content_changed:
                    self.log("Clipboard content does not match pattern")

            # 定时检查
            self.root.after(1000, self.update_clipboard_display)

        except Exception as e:
            error_str = str(e)
            is_clipboard_busy = (
                "WinError 0" in error_str or
                "OpenClipboard" in error_str or
                "clipboard" in error_str.lower()
            )

            if is_clipboard_busy:
                self.clipboard_display_error_count += 1
                current_time = time.time()

                should_log = (
                    self.clipboard_display_error_count == 1 or
                    self.clipboard_display_error_count == 20 or
                    self.clipboard_display_error_count % 50 == 0 or
                    (current_time - self.last_clipboard_display_error_time) > 60
                )

                if should_log:
                    self.log(
                        f"Clipboard temporarily busy (#{self.clipboard_display_error_count} attempts). This is normal when other apps are using the clipboard.",
                        level=logging.INFO
                    )
                    self.last_clipboard_display_error_time = current_time

                retry_delay = 1000
            else:
                self.clipboard_display_error_count += 1
                current_time = time.time()

                should_log = (
                    self.clipboard_display_error_count == 1 or
                    self.clipboard_display_error_count % 10 == 0 or
                    (current_time - self.last_clipboard_display_error_time) > 30
                )

                if should_log:
                    self.log(
                        f"Clipboard error (#{self.clipboard_display_error_count}): {error_str}",
                        level=logging.WARNING
                    )
                    self.last_clipboard_display_error_time = current_time

                retry_delay = min(1000 + (self.clipboard_display_error_count * 200), 5000)

            self.root.after(retry_delay, self.update_clipboard_display)

    def show_success_notification(self, content: str):
        """显示成功通知 (runs on main thread)."""
        if self.confirmation_dialog and self.confirmation_dialog.winfo_exists():
            self.confirmation_dialog.destroy()

        self.confirmation_dialog = tk.Toplevel(self.root)
        self.confirmation_dialog.overrideredirect(True)
        self.confirmation_dialog.attributes('-topmost', True)

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = 350
        window_height = 100
        x_position = screen_width - window_width - 20
        y_position = screen_height - window_height - 50
        self.confirmation_dialog.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

        self.confirmation_dialog.configure(bg="#D4EFDF")

        success_icon = "V"
        title_text = f"{success_icon} Content pasted to cell {self.current_cell}"
        title_label = tk.Label(
            self.confirmation_dialog,
            text=title_text,
            font=("Arial", 10, "bold"),
            bg="#D4EFDF",
            fg="#196F3D",
            padx=10, pady=5
        )
        title_label.pack(fill=tk.X)

        preview = content if len(content) < 40 else content[:37] + "..."
        content_label = tk.Label(
            self.confirmation_dialog,
            text=preview,
            font=("Consolas", 9),
            bg="#D4EFDF",
            fg="#1E8449",
            padx=10
        )
        content_label.pack(fill=tk.X)

        self._start_notification_timer(3)

        # Auto move to next row if enabled
        if self.auto_move_next:
            try:
                skip_rows = int(self.row_skip_var.get())
            except (ValueError, AttributeError):
                skip_rows = self.row_skip_count
            
            # Send message to queue instead of direct Excel access
            self.message_queue.put({'type': 'move_next_row', 'skip_rows': skip_rows})

        self.start_activity_monitoring()

    def show_error_notification(self, error_message: str):
        """显示错误通知 (runs on main thread)."""
        if self.confirmation_dialog and self.confirmation_dialog.winfo_exists():
            self.confirmation_dialog.destroy()

        self.confirmation_dialog = tk.Toplevel(self.root)
        self.confirmation_dialog.overrideredirect(True)
        self.confirmation_dialog.attributes('-topmost', True)

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = 350
        window_height = 100
        x_position = screen_width - window_width - 20
        y_position = screen_height - window_height - 50
        self.confirmation_dialog.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

        self.confirmation_dialog.configure(bg="#FADBD8")

        error_icon = "?"
        title_label = tk.Label(
            self.confirmation_dialog,
            text=f"{error_icon} Paste Failed",
            font=("Arial", 10, "bold"),
            bg="#FADBD8",
            fg="#943126",
            padx=10, pady=5
        )
        title_label.pack(fill=tk.X)

        error_label = tk.Label(
            self.confirmation_dialog,
            text=error_message,
            font=("Arial", 9),
            bg="#FADBD8",
            fg="#C0392B",
            wraplength=330,
            padx=10
        )
        error_label.pack(fill=tk.X)

        self._start_notification_timer(5)

        self.start_activity_monitoring()

    def _start_notification_timer(self, seconds: int):
        """启动通知自动关闭计时器"""
        self.root.after(seconds * 1000, self._close_notification)

    def _close_notification(self):
        """关闭通知窗口"""
        if self.confirmation_dialog and self.confirmation_dialog.winfo_exists():
            self.confirmation_dialog.destroy()
            self.confirmation_dialog = None

    def monitor_clipboard(self):
        """剪贴板监控线程 - 只检测变化，不访问Excel COM对象"""
        try:
            try:
                # ✅ CRITICAL FIX: Use lock to protect pyperclip access
                with self.clipboard_lock:
                    self.previous_content = pyperclip.paste()
            except Exception as initial_error:
                self.previous_content = ""
                self.log(f"Initial clipboard read failed: {initial_error}")
                traceback.print_exc()

            self.log("Clipboard monitoring started...")
            consecutive_clipboard_errors = 0
            consecutive_loop_errors = 0
            max_errors = 10

            while self.running:
                try:
                    try:
                        # ✅ CRITICAL FIX: Use lock to protect pyperclip access
                        with self.clipboard_lock:
                            current_content = pyperclip.paste()
                        consecutive_clipboard_errors = 0
                    except Exception as clipboard_error:
                        consecutive_clipboard_errors += 1
                        if consecutive_clipboard_errors <= 3:
                            self.log(f"Clipboard access error #{consecutive_clipboard_errors}: {clipboard_error}")
                        else:
                            self.log(f"Clipboard access error #{consecutive_clipboard_errors}")

                        if consecutive_clipboard_errors >= max_errors:
                            self.log("Too many consecutive clipboard errors; pausing before retry", level=logging.WARNING)
                            time.sleep(5.0)
                            consecutive_clipboard_errors = 0
                        else:
                            time.sleep(min(1.0 + 0.2 * consecutive_clipboard_errors, 2.0))
                        continue

                    sequence_number = self._get_clipboard_sequence_number()
                    sequence_changed = False
                    if sequence_number is not None:
                        if self.last_clipboard_sequence is None or sequence_number != self.last_clipboard_sequence:
                            sequence_changed = True
                            self.last_clipboard_sequence = sequence_number

                    content_changed = current_content != self.previous_content

                    if content_changed or sequence_changed:
                        if content_changed and current_content and current_content.strip():
                            content_preview = current_content[:30] + "..." if len(current_content) > 30 else current_content
                            self.log(f"New content detected: {content_preview}")

                        # ✅ CRITICAL FIX: Use message queue instead of direct Excel access
                        force_refresh = not content_changed and sequence_changed
                        self.message_queue.put({'type': 'update_clipboard_display', 'force': force_refresh})
                        
                        # ✅ CRITICAL FIX: Request cell refresh through queue, not direct access
                        if self.excel_app:
                            self.message_queue.put({'type': 'refresh_cell'})

                        self.previous_content = current_content

                    consecutive_loop_errors = 0
                    time.sleep(0.5)
                except Exception as loop_error:
                    consecutive_loop_errors += 1
                    self.log(f"Monitoring loop error #{consecutive_loop_errors}: {loop_error}")
                    traceback.print_exc()

                    if consecutive_loop_errors >= max_errors:
                        self.log("Too many monitoring loop errors; attempting automatic recovery", level=logging.WARNING)
                        time.sleep(5.0)
                        consecutive_loop_errors = 0
                    else:
                        time.sleep(min(1.0 + 0.2 * consecutive_loop_errors, 2.0))
        except Exception as e:
            self.log(f"Monitoring thread error: {str(e)}", level=logging.ERROR)
            traceback.print_exc()

    def start_monitoring(self):
        """开始监控"""
        try:
            if not self.excel_app:
                if not messagebox.askyesno("Warning", "Not connected to Excel. Do you want to connect now?"):
                    if not messagebox.askyesno("Warning", "Continue without Excel connection? The program will use keyboard shortcuts."):
                        return
                else:
                    if not self.connect_to_excel():
                        return

            # Initialize message queue
            if not hasattr(self, 'message_queue'):
                self._init_message_queue()

            self.running = True
            self.status_label.config(text="Running")
            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)

            self.set_excel_button.config(state=tk.DISABLED)

            self.log("Starting monitoring...")
            if self.target_excel != "Not specified":
                self.log(f"Target Excel: {self.target_excel}")

            if self.excel_app:
                # ✅ Safe: This runs on main thread
                self._handle_refresh_cell_request()
                self.log(f"Current selected cell: {self.current_cell}")

            self.ignore_initial_clipboard = False
            self.initial_clipboard_snapshot = ""
            try:
                # ✅ CRITICAL FIX: Use lock to protect pyperclip access
                with self.clipboard_lock:
                    baseline_clipboard = pyperclip.paste()
                if baseline_clipboard is None:
                    baseline_clipboard = ""
                self.previous_content = baseline_clipboard
                self.initial_clipboard_snapshot = baseline_clipboard
                if baseline_clipboard.strip():
                    try:
                        self.ignore_initial_clipboard = bool(self.is_valid_format(baseline_clipboard))
                    except Exception:
                        self.ignore_initial_clipboard = True
                else:
                    self.ignore_initial_clipboard = False
                self.log("Initial clipboard snapshot captured; waiting for next change")
            except Exception as snapshot_error:
                self.previous_content = ""
                self.initial_clipboard_snapshot = ""
                self.log(f"Initial clipboard snapshot failed, continuing with monitoring: {snapshot_error}")
                traceback.print_exc()

            self.last_clipboard_sequence = self._get_clipboard_sequence_number()

            self.update_clipboard_display()

            # Start monitoring thread - it won't access Excel COM objects
            self.monitor_thread = threading.Thread(target=self.monitor_clipboard, name="ClipboardMonitor")
            self.monitor_thread.daemon = True
            self.monitor_thread.start()

            self.log("Auto-paste mode: Content will be pasted automatically when detected")
            messagebox.showinfo(
                "Monitoring Started",
                "Auto-paste mode enabled.\n\n"
                "When matching content is detected, it will be automatically pasted to the current Excel cell with a notification.\n"
                "If the cell already has content, new content will be added on a new line.\n\n"
                "No action required - the process is fully automated."
            )

        except Exception as e:
            self.log(f"Start monitoring error: {str(e)}")
            messagebox.showerror("Error", f"Failed to start monitoring: {str(e)}")
            self.running = False

    def stop_monitoring(self):
        """停止监控"""
        try:
            self.running = False
            self.status_label.config(text="Stopped")
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)

            self.set_excel_button.config(state=tk.NORMAL)

            if self.confirmation_dialog and self.confirmation_dialog.winfo_exists():
                self.confirmation_dialog.destroy()
                self.confirmation_dialog = None

            self.log("Monitoring stopped")
            self.ignore_initial_clipboard = False
            self.initial_clipboard_snapshot = ""
        except Exception as e:
            self.log(f"Stop monitoring error: {str(e)}")
