import time
import traceback
import logging

import pythoncom
import win32com.client
import pyperclip
import tkinter as tk
from tkinter import messagebox, ttk


class ExcelMixin:
    """Encapsulate Excel interop related helpers.
    
    THREAD SAFETY: All methods in this mixin must be called from the main thread only.
    The schedule_cell_check method runs on the main thread using Tkinter's after() mechanism.
    """

    def schedule_cell_check(self):
        """轮询 Excel 活动单元格 (runs on main thread only).
        
        ⚠️ DISABLED: This was causing COM conflicts with message queue.
        Cell updates now happen only through message queue.
        """
        # ✅ CRITICAL FIX: Completely disabled to avoid COM conflicts
        # All cell checks now go through message queue
        
        # Schedule next check (but do nothing)
        try:
            self.excel_check_timer = self.root.after(5000, self.schedule_cell_check)
        except Exception:
            pass

    def refresh_current_cell(self):
        """刷新当前单元格地址 (must be called from main thread only)."""
        if self.excel_app:
            try:
                cell_address = self.excel_app.ActiveCell.Address
                if cell_address != self.current_cell:
                    self.current_cell = cell_address
                    self.cell_label.config(text=cell_address)
                    if self.running:
                        self.log(f"Cell selection changed: {cell_address}")
                return True
            except Exception as e:
                # Silently fail - Excel might be busy
                return False
        return False

    def connect_to_excel(self):
        """连接到 Excel (must be called from main thread)."""
        try:
            # ✅ CRITICAL: Initialize COM on main thread
            pythoncom.CoInitialize()

            is_reconnect = self.excel_app is not None
            
            # Try to get active Excel instance
            try:
                self.excel_app = win32com.client.GetActiveObject("Excel.Application")
            except Exception as e:
                self.log(f"Failed to get active Excel: {e}")
                messagebox.showerror(
                    "Excel Connection Error",
                    "Could not find an active Excel instance.\n\n"
                    "Please:\n"
                    "1. Open Excel\n"
                    "2. Open your workbook\n"
                    "3. Try connecting again"
                )
                return False

            # Verify Excel is responsive
            try:
                workbook_name = self.excel_app.ActiveWorkbook.Name
                sheet_name = self.excel_app.ActiveSheet.Name
            except Exception as e:
                self.log(f"Excel is not responsive: {e}")
                messagebox.showerror(
                    "Excel Error",
                    "Excel is running but not responsive.\n\n"
                    "Please make sure a workbook is open and try again."
                )
                self.excel_app = None
                return False

            self.target_excel = workbook_name
            self.excel_label.config(text=workbook_name)

            cell_address = self.excel_app.ActiveCell.Address
            self.current_cell = cell_address
            self.cell_label.config(text=cell_address)

            # Reset error counters
            self.cell_check_error_count = 0
            self.last_cell_check_error_time = 0

            if is_reconnect:
                self.log(f"Reconnected to Excel. Workbook: {workbook_name}, Sheet: {sheet_name}")
                messagebox.showinfo("Reconnected", f"Successfully reconnected to Excel!\nWorkbook: {workbook_name}")
            else:
                self.log(f"Connected to Excel. Workbook: {workbook_name}, Sheet: {sheet_name}")

            self.log(f"Current cell: {cell_address}")

            self.excel_cell_monitor_active = True

            self.paste_button.config(state=tk.NORMAL)

            return True

        except Exception as e:
            error_msg = str(e)
            self.log(f"Failed to connect to Excel: {error_msg}")
            
            # Provide helpful error messages
            if "0x800401E3" in error_msg or "Invalid class string" in error_msg:
                messagebox.showerror(
                    "Excel Not Found",
                    "Excel is not installed or not registered properly.\n\n"
                    "Please make sure Microsoft Excel is installed on this system."
                )
            elif "0x800401F0" in error_msg or "0x80010001" in error_msg:
                messagebox.showerror(
                    "Excel Connection Error",
                    "Could not connect to Excel due to COM error.\n\n"
                    "This might happen if:\n"
                    "- Excel is busy or unresponsive\n"
                    "- Excel is in protected mode\n"
                    "- Another program is blocking Excel\n\n"
                    "Try closing and reopening Excel, then try again."
                )
            else:
                messagebox.showerror(
                    "Excel Connection Error",
                    f"Failed to connect to Excel.\n\n"
                    f"Error: {error_msg}\n\n"
                    "Please make sure Excel is open with a workbook loaded."
                )
            
            traceback.print_exc()
            self.excel_app = None
            return False
        finally:
            # ✅ CRITICAL: Uninitialize COM after connection
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def monitor_excel_cell(self):
        """后台线程监控 Excel 单元格变动
        
        NOTE: This method is currently NOT USED to avoid COM threading issues.
        Cell monitoring is done on the main thread via schedule_cell_check().
        Keeping this method for potential future use with proper COM threading.
        """
        try:
            pythoncom.CoInitialize()

            last_cell = self.current_cell

            while self.running or self.excel_app:
                try:
                    if not self.excel_app:
                        break

                    current_cell = self.excel_app.ActiveCell.Address

                    if current_cell != last_cell:
                        self.current_cell = current_cell
                        self._run_on_ui_thread(self.cell_label.config, text=current_cell)
                        self.log(f"Cell selection changed: {current_cell}")
                        last_cell = current_cell

                except Exception as e:
                    self.log(f"Excel monitoring error: {str(e)}")
                    break

                time.sleep(0.1)

        except Exception as e:
            self.log(f"Excel monitoring thread error: {str(e)}")
        finally:
            pythoncom.CoUninitialize()

    def set_target_excel(self):
        """手动设置目标 Excel 名称"""
        try:
            dialog = tk.Toplevel(self.root)
            dialog.title("Set Target Excel")
            dialog.geometry("300x100")
            dialog.resizable(False, False)
            dialog.transient(self.root)
            dialog.grab_set()

            ttk.Label(dialog, text="Enter Excel filename:").pack(pady=(10, 5))
            excel_var = tk.StringVar()
            entry = ttk.Entry(dialog, textvariable=excel_var, width=30)
            entry.pack(pady=5, padx=10, fill=tk.X)
            entry.focus_set()

            def on_ok():
                value = excel_var.get().strip()
                if value:
                    self.target_excel = value
                    self.excel_label.config(text=value)
                    self.log(f"Target Excel set to: {value}")
                dialog.destroy()

            def on_cancel():
                dialog.destroy()

            button_frame = ttk.Frame(dialog)
            button_frame.pack(pady=5, fill=tk.X)

            ttk.Button(button_frame, text="OK", command=on_ok).pack(side=tk.RIGHT, padx=5)
            ttk.Button(button_frame, text="Cancel", command=on_cancel).pack(side=tk.RIGHT, padx=5)

            dialog.bind("<Return>", lambda event: on_ok())
            dialog.bind("<Escape>", lambda event: on_cancel())

            dialog.update_idletasks()
            x = self.root.winfo_x() + (self.root.winfo_width() - dialog.winfo_width()) // 2
            y = self.root.winfo_y() + (self.root.winfo_height() - dialog.winfo_height()) // 2
            dialog.geometry(f"+{x}+{y}")

            self.root.wait_window(dialog)
        except Exception as e:
            messagebox.showerror("Error", f"Error setting target Excel: {str(e)}")
            self.log(f"Error setting target Excel: {str(e)}")

    def paste_to_excel(self, show_error_dialog: bool = True):
        """粘贴数据到 Excel (must be called from main thread only)."""
        if not self.excel_app:
            if show_error_dialog:
                messagebox.showwarning("Warning", "Not connected to Excel. Please connect first.")
            return False

        try:
            # Refresh cell info before pasting
            self.refresh_current_cell()

            # ✅ CRITICAL FIX: Use lock to protect pyperclip access
            with self.clipboard_lock:
                content = pyperclip.paste()

            # ✅ Add retry logic for COM errors
            max_retries = 3
            retry_count = 0
            last_error = None

            while retry_count < max_retries:
                try:
                    current_value = self.excel_app.ActiveCell.Value
                    
                    if current_value:
                        self.log("Cell already has content, appending with new line")
                        new_value = f"{current_value}{chr(10)}{content}"
                        self.excel_app.ActiveCell.Value = new_value
                        self.log(f"Content appended to cell {self.current_cell}")
                    else:
                        self.excel_app.ActiveCell.Value = content
                        self.log(f"Content set to cell {self.current_cell}")

                    self.last_successful_paste_content = content
                    self.last_pasted_content = content
                    self.last_paste_time = time.time()

                    self.update_clipboard_display()
                    return True

                except Exception as e:
                    last_error = e
                    error_str = str(e)
                    
                    # Check if it's a COM error that might be temporary
                    if "0x800401F0" in error_str or "0x80010001" in error_str or "RPC" in error_str:
                        retry_count += 1
                        if retry_count < max_retries:
                            self.log(f"Paste attempt {retry_count} failed (COM error), retrying...", level=logging.WARNING)
                            time.sleep(0.5)  # Wait a bit before retry
                            continue
                    
                    # If not a retryable error, or retries exhausted, raise
                    raise

            # If we get here, all retries failed
            if last_error:
                raise last_error

        except Exception as e:
            error_msg = f"Failed to paste: {str(e)}"
            self.log(error_msg, level=logging.ERROR)
            
            # Provide helpful error messages
            error_str = str(e)
            if "0x800401F0" in error_str:
                error_msg = "Excel is not responding. Try:\n1. Click on Excel to activate it\n2. Wait a moment\n3. Try again"
            elif "0x80010001" in error_str:
                error_msg = "Excel is busy. Please wait for current operation to complete."
            elif "COM" in error_str or "RPC" in error_str:
                error_msg = "Excel connection lost. Please reconnect to Excel."
            
            if show_error_dialog:
                messagebox.showerror("Paste Error", error_msg)
            
            return False

    def reconnect_to_excel(self):
        """Reconnect to Excel after connection loss."""
        try:
            self.log("Attempting to reconnect to Excel...")
            success = self.connect_to_excel()
            
            if success:
                self.log("Reconnection successful")
                return True
            else:
                self.log("Reconnection failed")
                return False
        except Exception as e:
            self.log(f"Reconnection error: {e}", level=logging.ERROR)
            return False
