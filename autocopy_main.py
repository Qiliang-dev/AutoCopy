import pyperclip
import pyautogui
import logging
import atexit
import re
import time
import sys
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import functools
import win32com.client
import pythoncom
import traceback
import winsound
import ctypes
import win32api
import win32con
import queue
import json
from pathlib import Path
from ctypes import windll
from ctypes import wintypes

from autocopy.mixins.activity import ActivityMixin
from autocopy.mixins.clipboard import ClipboardMixin
from autocopy.mixins.excel import ExcelMixin
from autocopy.utils import resolve_resource_path, set_app_window_icon

class AutoCopyApp(ClipboardMixin, ActivityMixin, ExcelMixin):
    def __init__(self, root):
        # Initialize attributes first
        self.running = False
        self.monitor_thread = None
        self.previous_content = ""
        self.target_excel = "Not specified"  # Initialize before any method calls
        self.excel_app = None
        self.current_cell = "Not selected"
        self.ignore_initial_clipboard = False  # Skip auto-paste for baseline clipboard content
        self.excel_check_timer = None
        self.excel_monitor_thread = None
        self.excel_cell_monitor_active = False  # 新增标志，表示Excel单元格监控是否活跃
        self.clipboard_content = ""  # 当前剪贴板内容
        self.confirmation_dialog = None  # 确认对话框引用
        self.last_pasted_content = ""  # 上次粘贴的内容
        self.last_paste_time = 0
        self.last_clipboard_sequence = self._get_clipboard_sequence_number()  # 上次粘贴的时间戳
        self.auto_move_next = False  # 新增：是否自动移动到下一行
        self.row_skip_count = 1  # 新增：自动移动时跳过的行数
        self.reminder_dialog = None  # 提醒对话框
        self.reminder_enabled_var = None  # GUI toggle for reminder prompt
        self.reminder_spinbox = None  # deprecated placeholder
        self._reminder_flash_job = None
        self._reminder_update_job = None
        self.reminder_content_widget = None
        self._duplicate_logged = False
        self._faulthandler_file = None
        self.last_activity_time = 0  # 新增：最后活动时间
        self.activity_monitor_active = False  # 新增：活动监控状态
        self.global_hook_thread = None  # 新增：全局钩子线程
        self.activity_detected = False  # 新增：活动检测标志
        self.last_mouse_pos = None  # 新增：上次鼠标位置
        self.last_successful_paste_content = ""
        self.initial_clipboard_snapshot = ""  # 新增：上次成功粘贴的内容
        self.logs_dir = None
        self.log_file_path = None
        self.logger = None
        self.cell_check_error_count = 0  # 新增：单元格检查错误计数器
        self.last_cell_check_error_time = 0  # 新增：上次单元格检查错误时间
        self.clipboard_display_error_count = 0  # 剪贴板显示错误计数器
        self.last_clipboard_display_error_time = 0  # 上次剪贴板显示错误时间
        self.clipboard_update_job = None  # Track scheduled clipboard refresh callbacks
        
        # ✅ CRITICAL FIX: Initialize message queue for thread-safe Excel operations
        self.message_queue = queue.Queue()
        
        # ✅ CRITICAL FIX: Thread lock for pyperclip (not thread-safe!)
        self.clipboard_lock = threading.Lock()
        
        # ✅ CRITICAL FIX: Flag to prevent Excel operation conflicts
        self._processing_excel_operation = False
        
        # ✅ NEW FEATURE: Config file path
        self.config_dir = Path.cwd() / "config"
        self.config_file = self.config_dir / "settings.json"
        
        self.root = root
        self.root.title("AutoCopy Tool v1.9 - Stable Edition")
        self.root.geometry("650x550")  # 增加窗口尺寸，以容纳更多控件
        self.root.resizable(True, True)  # 允许用户调整窗口大小
        
        # 添加关闭窗口处理
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self._setup_logging()
        self._setup_exit_hooks()

        # Setup UI after initializing all attributes
        self.setup_ui()
        
        # ✅ CRITICAL FIX: Start message queue processor
        self._init_message_queue()
        
        # ✅ NEW FEATURE: Load saved settings
        self._load_settings()
        
    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 状态显示区域
        status_frame = ttk.LabelFrame(main_frame, text="Status", padding="10")
        status_frame.pack(fill=tk.X, pady=5)
        
        # Excel文件名区域 - 使用Grid布局以确保对齐
        ttk.Label(status_frame, text="Target Excel:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.excel_label = ttk.Label(status_frame, text=self.target_excel)
        self.excel_label.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        # Excel文件名设置按钮
        self.set_excel_button = ttk.Button(status_frame, text="Set Target", command=self.set_target_excel)
        self.set_excel_button.grid(row=0, column=2, padx=5, pady=5)
        
        # 连接Excel按钮
        self.connect_excel_button = ttk.Button(status_frame, text="Connect to Excel", command=self.connect_to_excel)
        self.connect_excel_button.grid(row=0, column=3, padx=5, pady=5)
        
        # Excel单元格信息
        ttk.Label(status_frame, text="Current Cell:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.cell_label = ttk.Label(status_frame, text=self.current_cell)
        self.cell_label.grid(row=1, column=1, sticky=tk.W, pady=5)
        
        # 刷新单元格按钮
        self.refresh_cell_button = ttk.Button(status_frame, text="Refresh Cell", command=self.refresh_current_cell)
        self.refresh_cell_button.grid(row=1, column=2, padx=5, pady=5)
        
        # 手动粘贴按钮 - 用于测试
        self.paste_button = ttk.Button(status_frame, text="Paste Now", command=self.paste_to_excel)
        self.paste_button.grid(row=1, column=3, padx=5, pady=5)
        
        # 运行状态
        ttk.Label(status_frame, text="Monitoring Status:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.status_label = ttk.Label(status_frame, text="Not Running")
        self.status_label.grid(row=2, column=1, sticky=tk.W, pady=5)
        
        # 控制按钮区域 - 分两行
        control_frame = ttk.LabelFrame(main_frame, text="Controls", padding="10")
        control_frame.pack(fill=tk.X, pady=5)
        
        # 第一行按钮
        self.start_button = ttk.Button(control_frame, text="Start Monitoring", command=self.start_monitoring)
        self.start_button.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.stop_button = ttk.Button(control_frame, text="Stop Monitoring", command=self.stop_monitoring, state=tk.DISABLED)
        self.stop_button.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        self.clear_log_button = ttk.Button(control_frame, text="Clear Log", command=self.clear_log)
        self.clear_log_button.grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        self.update_clipboard_button = ttk.Button(control_frame, text="Refresh Clipboard", command=self.update_clipboard_display)
        self.update_clipboard_button.grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)
        
        # 第二行按钮
        self.exit_button = ttk.Button(control_frame, text="Exit", command=self.on_closing)
        self.exit_button.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.auto_move_button = ttk.Button(control_frame, text="Auto Move Next: OFF", command=self.toggle_auto_move)
        self.auto_move_button.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 剪贴板内容显示区域
        clipboard_frame = ttk.LabelFrame(main_frame, text="Current Clipboard Content", padding="10")
        clipboard_frame.pack(fill=tk.X, pady=5)
        
        self.clipboard_text = scrolledtext.ScrolledText(clipboard_frame, height=4, width=70, wrap=tk.WORD)
        self.clipboard_text.pack(fill=tk.X, expand=True)
        self.clipboard_text.insert(tk.END, "(No content)")
        self.clipboard_text.config(state=tk.DISABLED)
        
        # Pattern Settings 区域
        format_frame = ttk.LabelFrame(main_frame, text="Pattern Settings", padding="10")
        format_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(format_frame, text="Pattern:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.format_var = tk.StringVar(value=r'^20\d{2}_\d{2}_\d{2}_\d{6}')
        format_entry = ttk.Entry(format_frame, textvariable=self.format_var, width=40)
        format_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5, columnspan=2)
        
        # 紧凑排列说明
        ttk.Label(format_frame, text="Duplicate Protection (s):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.duplicate_time_var = tk.StringVar(value="3")
        duplicate_entry = ttk.Spinbox(format_frame, from_=1, to=10, width=5, textvariable=self.duplicate_time_var)
        duplicate_entry.grid(row=1, column=1, sticky=tk.W, padx=2, pady=2)
        ttk.Label(format_frame, text="(No duplicate paste in seconds)", font=("Arial", 8)).grid(row=1, column=2, sticky=tk.W, padx=2, pady=2)
        
        self.reminder_enabled_var = tk.BooleanVar(value=False)
        reminder_checkbox = ttk.Checkbutton(
            format_frame,
            text="Enable Reminder",
            variable=self.reminder_enabled_var,
            command=self.on_reminder_toggle
        )
        reminder_checkbox.grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(format_frame, text="Row Skip Count:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=2)
        self.row_skip_var = tk.StringVar(value=str(self.row_skip_count))
        row_skip_entry = ttk.Spinbox(format_frame, from_=1, to=100, width=5, textvariable=self.row_skip_var)
        row_skip_entry.grid(row=3, column=1, sticky=tk.W, padx=2, pady=2)
        ttk.Label(format_frame, text="(Rows to skip when auto moving)", font=("Arial", 8)).grid(row=3, column=2, sticky=tk.W, padx=2, pady=2)
        
        # 匹配状态显示
        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="Log", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.log_text = tk.Text(log_frame, height=10, width=50, wrap=tk.WORD)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        # 添加初始日志消息
        self.log("Application started. Please connect to Excel and click 'Start Monitoring'")
        
        # 启动定时器，定期检查单元格和剪贴板
        self.schedule_cell_check()
        self.update_clipboard_display()
    
    def _execute_ui_task(self, task):
        """Run a callable and surface exceptions without touching Tk from worker threads."""
        try:
            task()
        except Exception:
            print("UI task raised an exception:")
            traceback.print_exc()

    def _run_on_ui_thread(self, callback, *args, **kwargs):
        """Ensure UI mutations run on Tk's main thread."""
        if not callable(callback):
            return

        task = functools.partial(callback, *args, **kwargs)
        if threading.current_thread() is threading.main_thread():
            self._execute_ui_task(task)
        else:
            try:
                self.root.after(0, lambda: self._execute_ui_task(task))
            except Exception:
                print("Failed to schedule UI callback:")
                traceback.print_exc()
    
    def _setup_logging(self):
        """Configure file logging and ensure logs directory exists."""
        try:
            from logging.handlers import RotatingFileHandler, TimedRotatingFileHandler

            self.logs_dir = Path.cwd() / "logs"
            self.logs_dir.mkdir(parents=True, exist_ok=True)
            self.log_file_path = self.logs_dir / "autocopy.log"

            logger = logging.getLogger("AutoCopy")
            logger.setLevel(logging.INFO)
            logger.propagate = False

            for handler in list(logger.handlers):
                logger.removeHandler(handler)

            # 改为按天分割日志：每天 0 点切分，默认保留 7 天
            handler = TimedRotatingFileHandler(
                self.log_file_path,
                when="midnight",
                interval=1,
                backupCount=7,
                encoding="utf-8",
                utc=False
            )
            handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
            logger.addHandler(handler)

            self.logger = logger
            self.logger.info("File logging initialized at %s", self.log_file_path)
            try:
                import faulthandler
                if faulthandler.is_enabled():
                    pass
                else:
                    self._faulthandler_file = open(self.log_file_path, "a", encoding="utf-8")
                    faulthandler.enable(self._faulthandler_file)
            except Exception as fh_exc:
                print(f"Faulthandler setup failed: {fh_exc}")

        except Exception as exc:
            self.logger = None
            print(f"Logging setup failed: {exc}")

    def _setup_exit_hooks(self):
        """Register hooks to capture normal exit and unhandled exceptions."""
        if getattr(self, "_exit_hooks_registered", False):
            return
        self._exit_hooks_registered = True

        def log_exit():
            try:
                target_logger = self.logger or logging.getLogger("AutoCopy")
                target_logger.info("Application exiting via atexit hook")
                try:
                    import faulthandler
                    if faulthandler.is_enabled():
                        faulthandler.disable()
                except Exception:
                    pass
                if getattr(self, "_faulthandler_file", None):
                    try:
                        self._faulthandler_file.close()
                    except Exception:
                        pass
                    self._faulthandler_file = None
            except Exception:
                pass

        previous_hook = sys.excepthook

        def log_exception(exc_type, exc_value, exc_traceback):
            exception_text = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
            try:
                target_logger = self.logger or logging.getLogger("AutoCopy")
                target_logger.error(
                    "Unhandled exception captured by sys.excepthook:\n%s",
                    exception_text
                )
            except Exception:
                print(f"Unhandled exception: {exception_text}")

            if previous_hook and previous_hook is not log_exception:
                previous_hook(exc_type, exc_value, exc_traceback)

        atexit.register(log_exit)
        sys.excepthook = log_exception

    def on_reminder_toggle(self):
        """Toggle reminder functionality without losing user settings."""
        self.stop_activity_monitoring()

        if self.reminder_enabled_var and self.reminder_enabled_var.get():
            # Immediately show/update reminder dialog when toggled on
            self.start_activity_monitoring()
        else:
            if self.reminder_dialog and self.reminder_dialog.winfo_exists():
                self.reminder_dialog.destroy()
                self.reminder_dialog = None

        state_text = 'enabled' if (self.reminder_enabled_var and self.reminder_enabled_var.get()) else 'disabled'
        self.log(f'Reminder feature {state_text}')
    def log(self, message, level=logging.INFO):
        """Write a message to both UI log and log file (if available)."""
        message_text = str(message)
        try:
            if getattr(self, "logger", None):
                self.logger.log(level, message_text)
        except Exception as exc:
            print(f"File logging error: {exc}")
        if hasattr(self, "root") and hasattr(self, "log_text"):
            def _write_to_ui():
                try:
                    widget = getattr(self, "log_text", None)
                    if not widget or not widget.winfo_exists():
                        return
                    previous_state = widget.cget("state")
                    widget.configure(state=tk.NORMAL)
                    widget.insert(tk.END, f"{time.strftime('%H:%M:%S')} - {message_text}\n")
                    widget.see(tk.END)
                    widget.configure(state=previous_state)
                except Exception as ui_error:
                    print(f"Log error: {ui_error}")
            self._run_on_ui_thread(_write_to_ui)

    def clear_log(self):
        """清除日志内容"""
        try:
            self.log_text.configure(state=tk.NORMAL)
            self.log_text.delete(1.0, tk.END)
            self.log_text.configure(state=tk.DISABLED)
            self.log("Log cleared")
        except Exception as e:
            print(f"Clear log error: {e}")
    
    def is_valid_format(self, text):
        """检查文本是否符合指定格式"""
        try:
            pattern = self.format_var.get()
            return bool(re.match(pattern, text))
        except re.error:
            self.log("Error: Invalid regular expression")
            return False
        except Exception as e:
            self.log(f"Format check error: {str(e)}")
            return False
    
    def _load_settings(self):
        """加载保存的设置"""
        try:
            # 确保配置目录存在
            self.config_dir.mkdir(parents=True, exist_ok=True)
            
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                
                # 加载设置到UI
                if 'row_skip_count' in settings:
                    self.row_skip_count = settings['row_skip_count']
                    self.row_skip_var.set(str(settings['row_skip_count']))
                
                if 'duplicate_time' in settings:
                    self.duplicate_time_var.set(str(settings['duplicate_time']))
                
                if 'reminder_enabled' in settings:
                    self.reminder_enabled_var.set(settings['reminder_enabled'])

                if 'auto_move_next' in settings:
                    self.auto_move_next = settings['auto_move_next']
                    button_text = "Auto Move Next: ON" if self.auto_move_next else "Auto Move Next: OFF"
                    self.auto_move_button.config(text=button_text)
                
                if 'pattern' in settings:
                    self.format_var.set(settings['pattern'])
                
                self.log("Settings loaded successfully")
            else:
                self.log("No saved settings found, using defaults")
                self._save_settings()  # 创建默认配置文件
                
        except Exception as e:
            self.log(f"Error loading settings: {e}")
            print(f"Error loading settings: {e}")
    
    def _save_settings(self):
        """保存当前设置"""
        try:
            # 确保配置目录存在
            self.config_dir.mkdir(parents=True, exist_ok=True)
            
            settings = {
                'row_skip_count': int(self.row_skip_var.get()) if hasattr(self, 'row_skip_var') else self.row_skip_count,
                'duplicate_time': float(self.duplicate_time_var.get()) if hasattr(self, 'duplicate_time_var') else 3,
                'reminder_enabled': self.reminder_enabled_var.get() if hasattr(self, 'reminder_enabled_var') and self.reminder_enabled_var else False,
                'auto_move_next': self.auto_move_next,
                'pattern': self.format_var.get() if hasattr(self, 'format_var') else r'^20\d{2}_\d{2}_\d{2}_\d{6}',
            }
            
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=4, ensure_ascii=False)
            
            self.log("Settings saved successfully")
            
        except Exception as e:
            self.log(f"Error saving settings: {e}")
            print(f"Error saving settings: {e}")
    
    def on_closing(self):
        """关闭窗口处理"""
        try:
            # ✅ NEW FEATURE: 保存设置
            self._save_settings()
            # 停止活动监控
            self.stop_activity_monitoring()
            
            # 尝试解绑所有全局快捷键
            for key in ("<Return>", "<KP_Enter>", "<Escape>"):
                try:
                    self.root.unbind_all(key)
                except:
                    pass
            
            if self.running:
                if messagebox.askokcancel("Exit", "Monitoring is still running. Are you sure you want to exit?"):
                    self.stop_monitoring()
                    self.running = False
                    
                    # 取消定时器
                    if self.excel_check_timer:
                        self.root.after_cancel(self.excel_check_timer)
                    
                    # 关闭确认对话框
                    if self.confirmation_dialog and self.confirmation_dialog.winfo_exists():
                        self.confirmation_dialog.destroy()
                    
                    # 释放Excel资源
                    if self.excel_app:
                        self.excel_app = None
                        
                    self.root.destroy()
                    
            else:
                # 取消定时器
                if self.excel_check_timer:
                    self.root.after_cancel(self.excel_check_timer)
                
                # 关闭确认对话框
                if self.confirmation_dialog and self.confirmation_dialog.winfo_exists():
                    self.confirmation_dialog.destroy()
                
                # 释放Excel资源
                if self.excel_app:
                    self.excel_app = None
                    
                self.root.destroy()
        except Exception as e:
            print(f"Error closing application: {str(e)}")
            self.root.destroy()

    def toggle_auto_move(self):
        """切换自动移动到下一行的功能"""
        self.auto_move_next = not self.auto_move_next
        button_text = "Auto Move Next: ON" if self.auto_move_next else "Auto Move Next: OFF"
        self.auto_move_button.config(text=button_text)
        self.log(f"Auto move to next row: {'Enabled' if self.auto_move_next else 'Disabled'}")
        # ✅ NEW FEATURE: 保存设置
        self._save_settings()

def main():
    try:
        root = tk.Tk()
        app = AutoCopyApp(root)
        set_app_window_icon(root)
        root.geometry("700x900")  # 启动时更大，确保所有控件显示
        root.minsize(600, 800)
        root.mainloop()
    except Exception as e:
        messagebox.showerror("Application Error", f"An error occurred: {str(e)}")
        print(f"Error: {str(e)}")
        traceback.print_exc()

if __name__ == "__main__":
    main()
