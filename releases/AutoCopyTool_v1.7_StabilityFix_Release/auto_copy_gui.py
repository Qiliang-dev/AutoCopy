import pyperclip
import pyautogui
import re
import time
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import win32com.client
import pythoncom
import traceback
import winsound
import ctypes
import win32api
import win32con
from ctypes import windll
from ctypes import wintypes
import requests
import logging
import os
import json
from datetime import datetime
import sys

class AutoCopyApp:
    def __init__(self, root):
        # Initialize logging first
        self.setup_logging()
        
        # Initialize attributes first
        self.running = False
        self.monitor_thread = None
        self.previous_content = ""
        self.target_excel = "Not specified"  # Initialize before any method calls
        self.excel_app = None
        self.current_cell = "Not selected"
        self.excel_check_timer = None
        self.excel_monitor_thread = None
        self.excel_cell_monitor_active = False  # 新增标志，表示Excel单元格监控是否活跃
        self.clipboard_content = ""  # 当前剪贴板内容
        self.confirmation_dialog = None  # 确认对话框引用
        self.last_pasted_content = ""  # 上次粘贴的内容
        self.last_paste_time = 0  # 上次粘贴的时间戳
        self.auto_move_next = False  # 新增：是否自动移动到下一行
        self.row_skip_count = 1  # 新增：自动移动时跳过的行数
        self.reminder_time = 20  # 新增：提醒等待时间（秒）
        self.reminder_timer = None  # 新增：提醒定时器
        self.reminder_dialog = None  # 新增：提醒对话框
        self.last_activity_time = 0  # 新增：最后活动时间
        self.activity_monitor_active = False  # 新增：活动监控状态
        self.global_hook_thread = None  # 新增：全局钩子线程
        self.activity_detected = False  # 新增：活动检测标志
        self.last_mouse_pos = None  # 新增：上次鼠标位置
        self.phone_notification_enabled = False  # 新增：手机通知开关，默认关闭
        self.last_successful_paste_content = ""  # 新增：上次成功粘贴的内容
        self.max_files_per_cell = None  # 默认：不限制单元格最大文件数
        self.settings_file = "autocopy_settings.json"  # 新增：设置文件路径
        self._cell_error_suppress_until = 0  # 新增：抑制Excel ActiveCell错误日志的时间戳
        self.ignore_max_move = True  # 新增：忽略最大数量，粘贴成功后总是下移（默认开启）
        
        self.root = root
        self.root.title("AutoCopy Tool")
        self.root.geometry("700x900")  # 增加窗口尺寸，以容纳更多控件
        self.root.resizable(True, True)  # 允许用户调整窗口大小
        
        # 添加关闭窗口处理
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 设置异常处理
        self.setup_exception_handling()
        
        # 加载保存的设置
        self.load_settings()
        
        # Setup UI after initializing all attributes
        self.setup_ui()
        
        # 更新UI按钮状态以反映加载的设置
        self.update_ui_from_settings()
        
        # 启动健康检查
        self.start_health_monitor()
        
        # 记录启动信息
        self.logger.info("AutoCopy application started successfully")
        self.log("Application started with crash protection and settings persistence")
    
    def setup_logging(self):
        """设置日志系统"""
        # 创建logs目录
        if not os.path.exists('logs'):
            os.makedirs('logs')
        
        # 配置日志格式
        log_filename = f"logs/autocopy_{datetime.now().strftime('%Y%m%d')}.log"
        
        # 配置logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_filename, encoding='utf-8'),
                logging.StreamHandler(sys.stdout)
            ]
        )
        
        self.logger = logging.getLogger(__name__)
        self.logger.info("Logging system initialized")
    
    def setup_exception_handling(self):
        """设置全局异常处理"""
        def handle_exception(exc_type, exc_value, exc_traceback):
            if issubclass(exc_type, KeyboardInterrupt):
                sys.__excepthook__(exc_type, exc_value, exc_traceback)
                return
            
            error_msg = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
            self.logger.critical(f"Uncaught exception: {error_msg}")
            
            # 显示用户友好的错误信息
            messagebox.showerror(
                "Application Error",
                f"An unexpected error occurred:\n\n{exc_type.__name__}: {str(exc_value)}\n\nCheck logs for details."
            )
        
        sys.excepthook = handle_exception
    
    def load_settings(self):
        """加载保存的设置"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                
                # 恢复设置值
                self.reminder_time = settings.get('reminder_time', 20)
                self.row_skip_count = settings.get('row_skip_count', 1)
                self.phone_notification_enabled = settings.get('phone_notification_enabled', False)
                self.max_files_per_cell = settings.get('max_files_per_cell', None)
                self.ignore_max_move = settings.get('ignore_max_move', True) # 加载忽略最大数量设置
                
                # 恢复其他设置
                if 'pattern' in settings:
                    self.pattern = settings['pattern']
                if 'duplicate_time' in settings:
                    self.duplicate_time = settings['duplicate_time']
                if 'ntfy_topic' in settings:
                    self.ntfy_topic = settings['ntfy_topic']
                
                self.logger.info("Settings loaded successfully")
            else:
                self.logger.info("No settings file found, using defaults")
                # 设置默认值
                self.pattern = r'^20\d{2}_\d{2}_\d{2}_\d{6}'
                self.duplicate_time = 3
                self.ntfy_topic = "autocopy_CPD"
                self.max_files_per_cell = None
                self.ignore_max_move = True # 默认忽略，总是下移
                
        except Exception as e:
            self.logger.error(f"Failed to load settings: {str(e)}")
            # 使用默认值
            self.pattern = r'^20\d{2}_\d{2}_\d{2}_\d{6}'
            self.duplicate_time = 3
            self.ntfy_topic = "autocopy_CPD"
            self.max_files_per_cell = None
            self.ignore_max_move = True # 默认忽略，总是下移
    
    def update_ui_from_settings(self):
        """根据加载的设置更新UI按钮状态"""
        try:
            # 更新Auto Move按钮
            if hasattr(self, 'auto_move_button'):
                self.auto_move_button.config(text="Auto Move Next: ON" if self.auto_move_next else "Auto Move Next: OFF")
            
            # 更新Ignore按钮
            if hasattr(self, 'ignore_max_move_button'):
                self.ignore_max_move_button.config(text="Ignore: ON" if self.ignore_max_move else "Ignore: OFF")
            
            # 更新Phone Notification按钮
            if hasattr(self, 'phone_notify_button'):
                self.phone_notify_button.config(text="Phone Alert: ON" if self.phone_notification_enabled else "Phone Alert: OFF")
            
            self.logger.info(f"UI updated from settings: ignore_max_move={self.ignore_max_move}, max_files={self.max_files_per_cell}")
        except Exception as e:
            self.logger.error(f"Error updating UI from settings: {str(e)}")
    
    def start_health_monitor(self):
        """启动程序健康状态监控"""
        self.health_check_counter = 0
        self.schedule_health_check()
    
    def schedule_health_check(self):
        """调度健康检查"""
        try:
            if hasattr(self, 'root') and self.root.winfo_exists():
                self.root.after(30000, self.perform_health_check)  # 每30秒检查一次
        except:
            pass  # 静默处理调度错误
    
    def perform_health_check(self):
        """执行健康检查"""
        try:
            self.health_check_counter += 1
            
            # 每10次健康检查（5分钟）记录一次状态
            if self.health_check_counter % 10 == 0:
                try:
                    status = f"Health check #{self.health_check_counter}: Running={self.running}, Excel={'Connected' if self.excel_app else 'Disconnected'}"
                    self.logger.info(status)
                except:
                    pass
            
            # 检查是否存在僵死的线程或资源泄露
            thread_count = threading.active_count()
            if thread_count > 10:  # 异常大量线程
                self.logger.warning(f"High thread count detected: {thread_count}")
            
            # 继续下次健康检查
            self.schedule_health_check()
            
        except Exception as e:
            try:
                self.logger.error(f"Health check error: {str(e)}")
            except:
                pass
            # 即使健康检查出错，也要继续调度
            self.schedule_health_check()
    
    def save_settings(self):
        """保存当前设置"""
        try:
            settings = {
                'reminder_time': self.reminder_time,
                'row_skip_count': self.row_skip_count,
                'phone_notification_enabled': self.phone_notification_enabled,
                'max_files_per_cell': self.max_files_per_cell,
                'ignore_max_move': self.ignore_max_move, # 保存忽略最大数量设置
                'pattern': getattr(self, 'format_var', None) and self.format_var.get() or self.pattern,
                'duplicate_time': getattr(self, 'duplicate_time_var', None) and self.duplicate_time_var.get() or self.duplicate_time,
                'ntfy_topic': getattr(self, 'ntfy_topic_var', None) and self.ntfy_topic_var.get() or self.ntfy_topic,
            }
            
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=2, ensure_ascii=False)
            
            self.logger.info("Settings saved successfully")
            
        except Exception as e:
            self.logger.error(f"Failed to save settings: {str(e)}")
            self.log(f"Error saving settings: {str(e)}")
    
    def count_files_in_cell(self, cell_content):
        """计算单元格中的文件数量（按文件名格式匹配，支持分号和换行分隔）"""
        if not cell_content:
            return 0
        
        try:
            # 获取当前使用的文件名格式pattern
            pattern = self.format_var.get()
            
            # 先按分号分割，再按换行分割，确保分号分隔的文件名也能被正确识别
            lines = []
            for raw in str(cell_content).replace(';', '\n').split('\n'):
                line = raw.strip()
                if line:
                    lines.append(line)
            
            # 统计匹配文件名格式的行数
            valid_file_count = 0
            for line in lines:
                try:
                    if re.match(pattern, line):
                        valid_file_count += 1
                except re.error:
                    # 如果正则表达式有错误，降级为按行计数
                    self.log(f"Warning: Invalid regex pattern, using line count instead")
                    return len(lines)
            
            return valid_file_count
            
        except Exception as e:
            self.log(f"Error counting files in cell: {str(e)}")
            # 出错时降级为按行计数
            lines = []
            for raw in str(cell_content).replace(';', '\n').split('\n'):
                line = raw.strip()
                if line:
                    lines.append(line)
            return len(lines)

    def parse_filenames_from_text(self, text):
        """从文本中提取符合当前文件名正则的条目（支持分号和换行分隔）。"""
        if not text:
            return []
        try:
            pattern = self.format_var.get()
        except Exception:
            pattern = r"^20\d{2}_\d{2}_\d{2}_\d{6}"
        items = []
        # 先按分号分割，再按换行分割，确保分号分隔的文件名也能被正确识别
        for raw in str(text).replace(';', '\n').split('\n'):
            line = raw.strip()
            if not line:
                continue
            try:
                if re.match(pattern, line):
                    items.append(line)
            except re.error:
                # 正则异常时，保守起见不匹配任何
                pass
        return items
    
    def check_cell_file_limit(self):
        """检查当前单元格是否超过文件数限制"""
        try:
            # 获取最大文件数限制
            max_files_str = self.max_files_var.get().strip()
            if not max_files_str:
                return True, 0  # 无限制
            
            try:
                max_files = int(max_files_str)
                if max_files <= 0:
                    return True, 0  # 无效值视为无限制
            except ValueError:
                self.log("Warning: Invalid max files value, treating as unlimited")
                return True, 0
            
            # 获取当前单元格内容
            if not self.excel_app:
                return True, 0
            
            pythoncom.CoInitialize()
            try:
                excel = win32com.client.GetActiveObject("Excel.Application")
                current_value = excel.ActiveCell.Value
                current_count = self.count_files_in_cell(current_value)
                
                if current_count >= max_files:
                    return False, current_count
                else:
                    return True, current_count
                    
            finally:
                try:
                    pythoncom.CoUninitialize()
                except:
                    pass
                
        except Exception as e:
            self.logger.error(f"Error checking cell file limit: {str(e)}")
            return True, 0  # 出错时允许粘贴
        
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
        
        # 当前单元格文件数量显示
        ttk.Label(status_frame, text="Files Count:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.cell_count_label = ttk.Label(status_frame, text="0", foreground="blue")
        self.cell_count_label.grid(row=2, column=1, sticky=tk.W, pady=5)
        
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
        # 添加手机提醒测试按钮
        self.test_phone_button = ttk.Button(control_frame, text="Test Phone Alert", command=self.test_phone_alert)
        self.test_phone_button.grid(row=1, column=2, padx=5, pady=5, sticky=tk.W)
        # 新增：忽略Max限制后，总是粘贴成功就下移
        self.ignore_max_move_button = ttk.Button(control_frame, text="Ignore: ON" if self.ignore_max_move else "Ignore: OFF", command=self.toggle_ignore_max_move)
        self.ignore_max_move_button.grid(row=1, column=3, padx=5, pady=5, sticky=tk.W)
        
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
        self.format_var = tk.StringVar(value=self.pattern)
        format_entry = ttk.Entry(format_frame, textvariable=self.format_var, width=40)
        format_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5, columnspan=2)
        
        # 紧凑排列说明
        ttk.Label(format_frame, text="Duplicate Protection (s):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.duplicate_time_var = tk.StringVar(value=str(self.duplicate_time))
        duplicate_entry = ttk.Spinbox(format_frame, from_=1, to=10, width=5, textvariable=self.duplicate_time_var)
        duplicate_entry.grid(row=1, column=1, sticky=tk.W, padx=2, pady=2)
        ttk.Label(format_frame, text="(No duplicate paste in seconds)", font=("Arial", 8)).grid(row=1, column=2, sticky=tk.W, padx=2, pady=2)
        
        ttk.Label(format_frame, text="Reminder Time (s):").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        self.reminder_time_var = tk.StringVar(value=str(self.reminder_time))
        reminder_entry = ttk.Spinbox(format_frame, from_=5, to=300, width=5, textvariable=self.reminder_time_var)
        reminder_entry.grid(row=2, column=1, sticky=tk.W, padx=2, pady=2)
        ttk.Label(format_frame, text="(Show reminder if no activity)", font=("Arial", 8)).grid(row=2, column=2, sticky=tk.W, padx=2, pady=2)
        
        ttk.Label(format_frame, text="Row Skip Count:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=2)
        self.row_skip_var = tk.StringVar(value=str(self.row_skip_count))
        row_skip_entry = ttk.Spinbox(format_frame, from_=1, to=100, width=5, textvariable=self.row_skip_var)
        row_skip_entry.grid(row=3, column=1, sticky=tk.W, padx=2, pady=2)
        ttk.Label(format_frame, text="(Rows to skip when auto moving)", font=("Arial", 8)).grid(row=3, column=2, sticky=tk.W, padx=2, pady=2)
        
        # 新增：单元格文件数限制
        ttk.Label(format_frame, text="Max Files Per Cell:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=2)
        self.max_files_var = tk.StringVar(value=str(self.max_files_per_cell) if self.max_files_per_cell else "")
        max_files_entry = ttk.Entry(format_frame, textvariable=self.max_files_var, width=5)
        max_files_entry.grid(row=4, column=1, sticky=tk.W, padx=2, pady=2)
        ttk.Label(format_frame, text="(Empty = unlimited, number = max files)", font=("Arial", 8)).grid(row=4, column=2, sticky=tk.W, padx=2, pady=2)
        
        # 手机通知设置
        ttk.Label(format_frame, text="Phone Notification:").grid(row=5, column=0, sticky=tk.W, padx=5, pady=2)
        self.phone_notify_button = ttk.Button(format_frame, text="Phone Alert: ON" if self.phone_notification_enabled else "Phone Alert: OFF", 
                                             command=self.toggle_phone_notification)
        self.phone_notify_button.grid(row=5, column=1, sticky=tk.W, padx=2, pady=2)
        
        self.ntfy_topic_var = tk.StringVar(value=self.ntfy_topic)
        topic_entry = ttk.Entry(format_frame, textvariable=self.ntfy_topic_var, width=20)
        topic_entry.grid(row=5, column=2, sticky=tk.W, padx=2, pady=2)
        ttk.Label(format_frame, text="(ntfy topic name)", font=("Arial", 8)).grid(row=5, column=3, sticky=tk.W, padx=2, pady=2)
        
        # 匹配状态显示
        ttk.Label(format_frame, text="Match Status:").grid(row=5, column=0, sticky=tk.W, padx=5, pady=5)
        self.match_status_label = ttk.Label(format_frame, text="Not checked")
        self.match_status_label.grid(row=5, column=1, sticky=tk.W, padx=5, pady=5)
        
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
    
    def paste_to_excel(self, show_error_dialog=True):
        """手动将剪贴板内容粘贴到Excel"""
        if not self.excel_app:
            if show_error_dialog:
                messagebox.showwarning("Warning", "Not connected to Excel. Please connect first.")
            return False
            
        try:
            # 确保 COM 已初始化
            pythoncom.CoInitialize()
            
            # 重新获取 Excel 应用程序实例
            excel = win32com.client.GetActiveObject("Excel.Application")
            
            # 刷新当前单元格
            self.refresh_current_cell()
            
            # 获取剪贴板内容
            content = pyperclip.paste()
            
            # 检查文件数限制（单元格是否已满）
            can_paste, current_count = self.check_cell_file_limit()
            max_files_str = self.max_files_var.get().strip()
            
            if max_files_str and not can_paste:
                # 超过文件数限制（当前格已满）
                try:
                    max_files = int(max_files_str)
                    error_msg = f"Cell file limit exceeded!\n\nCurrent files: {current_count}\nMax allowed: {max_files}\n\nPlease choose a different cell or increase the limit."
                    self.log(f"Paste blocked: Cell has {current_count} files, limit is {max_files}")
                    if show_error_dialog:
                        messagebox.showerror("File Limit Exceeded", error_msg)
                    return False
                except ValueError:
                    pass
            
            # 若设置了max，并且剪贴板中的文件名数量 > 可用空间，则直接报错，不移动、不粘贴
            if max_files_str:
                try:
                    max_files = int(max_files_str)
                    if max_files > 0:
                        clipboard_items = self.parse_filenames_from_text(content)
                        # 仅当剪贴板解析出>=1个有效文件名时才做严格校验
                        if clipboard_items:
                            available = max(0, max_files - current_count)
                            if len(clipboard_items) > available:
                                error_msg = (
                                    "Clipboard contains more items than remaining capacity in this cell.\n\n"
                                    f"Current files: {current_count}\n"
                                    f"Max allowed: {max_files}\n"
                                    f"Remaining capacity: {available}\n"
                                    f"Clipboard items (valid): {len(clipboard_items)}\n\n"
                                    "Please reduce clipboard items, increase the limit, or choose another cell."
                                )
                                self.log(
                                    f"Paste blocked: clipboard {len(clipboard_items)} > available {available} (limit {max_files})"
                                )
                                if show_error_dialog:
                                    messagebox.showerror("File Limit Exceeded", error_msg)
                                return False
                except ValueError:
                    pass
            
            # 正常粘贴逻辑
            current_value = excel.ActiveCell.Value
            if current_value:
                self.log("Cell already has content, appending with new line")
                new_value = f"{current_value}{chr(10)}{content}"
                excel.ActiveCell.Value = new_value
                self.log(f"Content appended to cell {self.current_cell}")
            else:
                excel.ActiveCell.Value = content
                self.log(f"Content set to cell {self.current_cell}")
            
            self.last_successful_paste_content = content
            self.update_clipboard_display()
            return True
        except Exception as e:
            error_msg = f"Failed to paste: {str(e)}"
            self.log(error_msg)
            self.logger.error(f"Paste error: {str(e)}")
            if show_error_dialog:
                messagebox.showerror("Paste Error", error_msg)
            return False
        finally:
            pythoncom.CoUninitialize()
    
    def update_clipboard_display(self):
        """更新剪贴板内容显示"""
        try:
            # 增强异常处理，防止剪贴板访问失败导致崩溃
            content = None
        try:
            content = pyperclip.paste()
            except Exception as clipboard_error:
                self.logger.debug(f"Clipboard access failed: {str(clipboard_error)}")
                # 剪贴板访问失败时跳过这次检查，继续下次循环
                self.root.after(1000, self.update_clipboard_display)
                return
            
            # 防止None或异常内容
            if content is None:
                self.root.after(1000, self.update_clipboard_display)
                return
            
            # 只有在内容变化时更新 - 使用更严格的比较并防止重复触发
            if content != self.clipboard_content and content.strip() != "":
                self.log(f"New content detected: {content}")
                
                # 记录上一次粘贴的内容和时间戳，用于防止重复处理
                current_time = time.time()
                
                # 检查是否是在短时间内尝试粘贴相同内容（防止重复）
                is_duplicate = False
                if hasattr(self, 'last_pasted_content') and hasattr(self, 'last_paste_time'):
                    # 获取用户设置的防重复时间间隔（秒）
                    try:
                        duplicate_threshold = float(self.duplicate_time_var.get())
                    except (ValueError, AttributeError):
                        duplicate_threshold = 3.0  # 默认值
                    
                    time_diff = current_time - self.last_paste_time
                    content_same = content == self.last_pasted_content
                    
                    # 如果在设定时间内尝试粘贴相同内容，视为重复操作
                    if content_same and time_diff < duplicate_threshold:
                        is_duplicate = True
                        self.log(f"Ignored duplicate paste attempt (within {time_diff:.1f}s, threshold: {duplicate_threshold}s)")
                
                # 更新记录的剪贴板内容
                self.clipboard_content = content
                
                # 安全更新文本显示
                try:
                self.clipboard_text.config(state=tk.NORMAL)
                self.clipboard_text.delete(1.0, tk.END)
                
                # 限制显示长度，防止过长内容
                if len(content) > 500:
                    display_content = content[:500] + "... (content truncated)"
                else:
                    display_content = content
                    
                self.clipboard_text.insert(tk.END, display_content)
                self.clipboard_text.config(state=tk.DISABLED)
                except Exception as ui_error:
                    self.logger.debug(f"UI update failed: {str(ui_error)}")
                
                # 检查是否匹配格式
                try:
                match_result = self.is_valid_format(content)
                if match_result:
                    self.match_status_label.config(text="Matches Pattern", foreground="green")
                    
                        # 检查是否与上次成功粘贴的内容相同
                        is_same_as_last_pasted = hasattr(self, 'last_successful_paste_content') and content == self.last_successful_paste_content
                        
                        # 如果正在监控并且匹配成功且不是重复内容且不与上次粘贴内容相同，显示通知并自动粘贴
                        if self.running and not is_duplicate and not is_same_as_last_pasted:
                        # 保存此次操作数据，用于防止重复
                        self.last_pasted_content = content
                        self.last_paste_time = current_time
                        
                        # 延迟一小段时间后执行粘贴操作，给UI时间更新
                        self.root.after(100, lambda: self.auto_paste_with_notification(content))
                        elif is_same_as_last_pasted:
                            self.log(f"Content same as last successful paste, skipping auto-paste")
                        elif is_duplicate:
                            self.log(f"Duplicate content detected within time threshold, skipping auto-paste")
                else:
                    self.match_status_label.config(text="Does Not Match", foreground="red")
                except Exception as format_error:
                    self.logger.debug(f"Format checking failed: {str(format_error)}")
            
            # 设置剪贴板检查定时器 - 每秒更新一次
            if hasattr(self, 'root') and self.root.winfo_exists():
            self.root.after(1000, self.update_clipboard_display)
            
        except Exception as e:
            try:
                self.logger.error(f"Critical error in clipboard monitoring: {str(e)}")
                # 只有在UI仍然存在时才尝试记录日志
                if hasattr(self, 'root') and self.root.winfo_exists():
                    self.log(f"Clipboard monitoring error: {str(e)}")
            except:
                pass  # 静默处理UI错误
                
            # 继续检查，但要确保根窗口仍然存在
            try:
                if hasattr(self, 'root') and self.root.winfo_exists():
                    self.root.after(2000, self.update_clipboard_display)
            except:
                # 如果UI已经不存在，停止监控
                return
    
    def auto_paste_with_notification(self, content):
        """自动粘贴并显示简短通知"""
        try:
            # 执行粘贴操作 - 传入False禁用错误对话框
            success = self.paste_to_excel(show_error_dialog=False)
            
            # 显示结果通知
            if success:
                self.show_success_notification(content)
            else:
                error_msg = "Paste failed. Please check Excel connection."
                self.log(error_msg)
                self.show_error_notification(error_msg)
            
            # 重置活动监控
            self.last_activity_time = time.time()
            # 启动活动监控
            self.start_activity_monitoring()
                
        except Exception as e:
            self.log(f"Auto paste error: {str(e)}")
            self.show_error_notification(f"Paste error: {str(e)}")
    
    def show_success_notification(self, content):
        """显示成功粘贴的通知"""
        # 关闭之前的通知
        if self.confirmation_dialog and self.confirmation_dialog.winfo_exists():
            self.confirmation_dialog.destroy()
        
        # 创建无边框通知窗口
        self.confirmation_dialog = tk.Toplevel(self.root)
        self.confirmation_dialog.overrideredirect(True)  # 移除所有窗口装饰
        self.confirmation_dialog.attributes('-topmost', True)  # 保持在最前面
        
        # 固定位置 - 右下角
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = 350
        window_height = 100
        x_position = screen_width - window_width - 20
        y_position = screen_height - window_height - 50
        self.confirmation_dialog.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
        
        # 设置绿色背景
        self.confirmation_dialog.configure(bg="#D4EFDF")
        
        # 创建标签
        success_icon = "✓"  # 成功图标
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
        
        # 显示简短的内容预览
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
        
        # 自动关闭通知的倒计时
        self._start_notification_timer(3)  # 3秒后自动关闭
        
        # 在显示通知后，如果启用了自动移动到下一行，则检查是否需要移动
        if self.auto_move_next and self.excel_app:
            try:
                # 检查是否需要移动
                should_move = False
                max_files_str = self.max_files_var.get().strip()
                
                if self.ignore_max_move:
                    # 忽略Max限制：只要粘贴成功就移动
                    should_move = True
                    self.log("Ignore max for move enabled: moving to next row after paste")
                elif max_files_str:
                    try:
                        max_files = int(max_files_str)
                        if max_files > 0:
                            # 获取当前单元格内容并计算文件数
                            current_value = self.excel_app.ActiveCell.Value
                            current_count = self.count_files_in_cell(current_value)
                            
                            # 如果当前单元格已满，则需要移动
                            if current_count >= max_files:
                                should_move = True
                                self.log(f"Cell is full ({current_count}/{max_files} valid files), moving to next row")
                            else:
                                self.log(f"Cell not full ({current_count}/{max_files} valid files), staying in current cell")
                    except ValueError:
                        # 如果max_files不是有效数字，按原来的逻辑移动
                        should_move = True
                        self.log("Invalid max files setting, using default auto-move behavior")
                else:
                    # 如果没有设置文件数限制，按原来的逻辑移动
                    should_move = True
                    self.log("No file limit set, using default auto-move behavior")
                
                # 只有在需要移动时才移动
                if should_move:
                    # 获取当前单元格的行和列
                    current_cell = self.excel_app.ActiveCell
                    current_row = current_cell.Row
                    current_column = current_cell.Column
                    
                    # 获取要跳过的行数
                    try:
                        skip_rows = int(self.row_skip_var.get())
                    except (ValueError, AttributeError):
                        skip_rows = self.row_skip_count
                    
                    # 移动到指定行数后的单元格，保持列不变
                    next_cell = self.excel_app.ActiveSheet.Cells(current_row + skip_rows, current_column)
                    next_cell.Select()
                    
                    # 更新当前单元格显示
                    self.refresh_current_cell()
                    
                    # 记录日志
                    self.log(f"Automatically moved {skip_rows} rows down: {next_cell.Address}")
                    
            except Exception as e:
                self.log(f"Error in smart auto-move logic: {str(e)}")
                self.logger.error(f"Smart auto-move error: {str(e)}", exc_info=True)
        
        # 启动活动监控
        self.start_activity_monitoring()
    
    def show_error_notification(self, error_message):
        """显示错误通知"""
        # 关闭之前的通知
        if self.confirmation_dialog and self.confirmation_dialog.winfo_exists():
            self.confirmation_dialog.destroy()
        
        # 创建无边框通知窗口
        self.confirmation_dialog = tk.Toplevel(self.root)
        self.confirmation_dialog.overrideredirect(True)  # 移除所有窗口装饰
        self.confirmation_dialog.attributes('-topmost', True)  # 保持在最前面
        
        # 固定位置 - 右下角
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = 350
        window_height = 100
        x_position = screen_width - window_width - 20
        y_position = screen_height - window_height - 50
        self.confirmation_dialog.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
        
        # 设置红色背景
        self.confirmation_dialog.configure(bg="#FADBD8")
        
        # 创建标签
        error_icon = "✗"  # 错误图标
        title_label = tk.Label(
            self.confirmation_dialog,
            text=f"{error_icon} Paste Failed",
            font=("Arial", 10, "bold"),
            bg="#FADBD8",
            fg="#943126",
            padx=10, pady=5
        )
        title_label.pack(fill=tk.X)
        
        # 显示错误信息
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
        
        # 自动关闭通知的倒计时
        self._start_notification_timer(5)  # 5秒后自动关闭
        
        # 启动活动监控
        self.start_activity_monitoring()
    
    def _start_notification_timer(self, seconds):
        """启动通知自动关闭倒计时"""
        self.root.after(seconds * 1000, self._close_notification)
        
    def _close_notification(self):
        """关闭通知窗口"""
        if self.confirmation_dialog and self.confirmation_dialog.winfo_exists():
            self.confirmation_dialog.destroy()
            self.confirmation_dialog = None
    
    def schedule_cell_check(self):
        """定期检查Excel单元格"""
        if self.excel_app:
            try:
                try:
                    pythoncom.CoInitialize()
                except Exception:
                    pass
                # 直接复用已连接的实例，避免每次GetActiveObject
                cell_address = self.excel_app.ActiveCell.Address
                if cell_address != self.current_cell:
                    self.current_cell = cell_address
                    self.cell_label.config(text=cell_address)
                    if self.running:
                        self.log(f"Cell selection changed: {cell_address}")
            except Exception as e:
                import time as _t
                now = _t.time()
                if now >= getattr(self, '_cell_error_suppress_until', 0):
                    # 仅写到文件debug，避免UI刷屏
                    try:
                        self.logger.debug(f"Error checking cell: {e}")
                    except Exception:
                        pass
                    self._cell_error_suppress_until = now + 1.0
            finally:
                try:
                    pythoncom.CoUninitialize()
                except:
                    pass
        
        # 每250毫秒检查一次单元格，降低频率
        self.excel_check_timer = self.root.after(250, self.schedule_cell_check)
    
    def refresh_current_cell(self):
        """刷新当前选中的单元格"""
        if self.excel_app:
            try:
                # 确保 COM 已初始化
                pythoncom.CoInitialize()
                # 重新获取 Excel 应用程序实例
                excel = win32com.client.GetActiveObject("Excel.Application")
                cell_address = excel.ActiveCell.Address
                if cell_address != self.current_cell:
                    def _apply_update():
                    self.current_cell = cell_address
                    self.cell_label.config(text=cell_address)
                    if self.running:
                        self.log(f"Cell selection changed: {cell_address}")
                        # 安全更新当前单元格中的文件数量
                        self._update_cell_count_display(excel)
                    if threading.current_thread() is threading.main_thread():
                        _apply_update()
                    else:
                        self.root.after(0, _apply_update)
                else:
                    # 即使地址相同，也更新文件数量（内容可能变化）
                    def _update_count_only():
                        self._update_cell_count_display(excel)
                    if threading.current_thread() is threading.main_thread():
                        _update_count_only()
                    else:
                        self.root.after(0, _update_count_only)
                return True
            except Exception as e:
                import time as _t
                now = _t.time()
                if now >= getattr(self, '_cell_error_suppress_until', 0):
                    try:
                        self.logger.debug(f"Error refreshing cell: {e}")
                    except Exception:
                        pass
                    self._cell_error_suppress_until = now + 1.0
                return False
            finally:
                # 安全清理 COM
                try:
                    pythoncom.CoUninitialize()
                except:
                    pass  # 静默处理COM清理错误
        return False
    
    def _update_cell_count_display(self, excel):
        """安全更新当前单元格文件数量显示"""
        try:
            current_value = excel.ActiveCell.Value
            count = self.count_files_in_cell(current_value)
            
            # 记录当前单元格文件数量到日志
            if self.running:  # 只在监控时记录
                cell_address = excel.ActiveCell.Address
                self.log(f"Current cell {cell_address} contains {count} files")
            
            # 安全更新UI
            def _ui_update():
                if hasattr(self, 'cell_count_label'):
                    if self.max_files_per_cell:
                        self.cell_count_label.config(text=f"{count}/{self.max_files_per_cell}")
                        if count >= self.max_files_per_cell:
                            self.cell_count_label.config(foreground="red")
                        else:
                            self.cell_count_label.config(foreground="blue")
                    else:
                        self.cell_count_label.config(text=str(count), foreground="blue")
            
            if threading.current_thread() is threading.main_thread():
                _ui_update()
            else:
                self.root.after(0, _ui_update)
                
        except Exception as e:
            # 静默处理错误，不影响主要功能
            try:
                self.logger.debug(f"Error updating cell count display: {e}")
            except:
                pass
    
    def connect_to_excel(self):
        """连接到Excel应用程序"""
        try:
            self.logger.info("Attempting to connect to Excel")
            
            # 初始化COM线程 - 主UI线程
            pythoncom.CoInitialize()
            
            # 获取Excel应用程序实例
            self.excel_app = win32com.client.GetActiveObject("Excel.Application")
            self.logger.info("Successfully obtained Excel application object")
            
            # 获取当前活动工作簿和工作表
            workbook_name = self.excel_app.ActiveWorkbook.Name
            sheet_name = self.excel_app.ActiveSheet.Name
            
            # 更新Excel文件名
            self.target_excel = workbook_name
            self.excel_label.config(text=workbook_name)
            
            # 获取当前选中的单元格
            cell_address = self.excel_app.ActiveCell.Address
            self.current_cell = cell_address
            self.cell_label.config(text=cell_address)
            
            self.log(f"Connected to Excel. Workbook: {workbook_name}, Sheet: {sheet_name}")
            self.log(f"Current cell: {cell_address}")
            self.logger.info(f"Excel connection successful - Workbook: {workbook_name}, Sheet: {sheet_name}, Cell: {cell_address}")
            
            # 启动单元格监控 - 现在使用定时器代替线程
            self.excel_cell_monitor_active = True
            
            # 确保可以粘贴
            self.paste_button.config(state=tk.NORMAL)
            
            return True
            
        except Exception as e:
            error_msg = f"Failed to connect to Excel: {str(e)}"
            self.log(error_msg)
            self.logger.error(error_msg, exc_info=True)
            messagebox.showerror("Excel Connection Error", 
                                 "Failed to connect to Excel. Please make sure Excel is open and try again.")
            return False
        finally:
            pythoncom.CoUninitialize()
    
    def monitor_excel_cell(self):
        """监控Excel单元格变化 - 此方法现已弃用，使用定时器代替"""
        try:
            # 初始化COM线程
            pythoncom.CoInitialize()
            
            last_cell = self.current_cell
            
            while self.running:
                try:
                    # 检查Excel是否还在运行 - 更安全的检查
                    if not self.excel_app:
                        break
                        
                    # 尝试获取当前单元格地址 - 增强异常处理
                    try:
                    current_cell = self.excel_app.ActiveCell.Address
                        excel_available = True
                    except Exception as excel_error:
                        # Excel不可用 - 可能已关闭
                        excel_available = False
                        current_time = time.time()
                        if not hasattr(self, '_last_excel_error_time'):
                            self._last_excel_error_time = 0
                        if (current_time - self._last_excel_error_time) >= 5.0:  # 5秒记录一次
                            self.logger.debug(f"Excel access error: {str(excel_error)}")
                            self._last_excel_error_time = current_time
                        
                        # 等待更长时间再重试
                        time.sleep(1.0)
                        continue
                    
                    if not excel_available:
                        break
                        
                    # 检查单元格变化
                    if current_cell != last_cell:
                        self.current_cell = current_cell
                        
                        # 安全地更新UI
                        try:
                            if hasattr(self, 'root') and self.root.winfo_exists():
                        self.root.after(0, lambda: self.cell_label.config(text=current_cell))
                        self.root.after(0, lambda: self.log(f"Cell selection changed: {current_cell}"))
                        except Exception as ui_error:
                            self.logger.debug(f"UI update error in cell check: {str(ui_error)}")
                        
                        last_cell = current_cell
                        
                except Exception as e:
                    # 如果出错，可能是Excel已关闭
                    self.root.after(0, lambda: self.log(f"Excel monitoring error: {str(e)}"))
                    break
                
                # 短暂暂停以减少CPU使用
                time.sleep(0.1)
                
        except Exception as e:
            self.root.after(0, lambda: self.log(f"Excel monitoring thread error: {str(e)}"))
        finally:
            # 结束COM线程
            pythoncom.CoUninitialize()
    
    def set_target_excel(self):
        """设置目标Excel文件名"""
        try:
            # Create dialog to enter Excel name
            dialog = tk.Toplevel(self.root)
            dialog.title("Set Target Excel")
            dialog.geometry("300x100")
            dialog.resizable(False, False)
            dialog.transient(self.root)
            dialog.grab_set()
            
            ttk.Label(dialog, text="Enter Excel filename:").pack(pady=(10,5))
            
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
            
            # Handle Enter key
            dialog.bind("<Return>", lambda event: on_ok())
            dialog.bind("<Escape>", lambda event: on_cancel())
            
            # Center the dialog on the main window
            dialog.update_idletasks()
            x = self.root.winfo_x() + (self.root.winfo_width() - dialog.winfo_width()) // 2
            y = self.root.winfo_y() + (self.root.winfo_height() - dialog.winfo_height()) // 2
            dialog.geometry(f"+{x}+{y}")
            
            # Wait for dialog to close
            self.root.wait_window(dialog)
        except Exception as e:
            messagebox.showerror("Error", f"Error setting target Excel: {str(e)}")
            self.log(f"Error setting target Excel: {str(e)}")
    
    def log(self, message):
        """线程安全的日志记录"""
        def _log():
        try:
                # 同时记录到UI和文件日志
            self.log_text.configure(state=tk.NORMAL)
            self.log_text.insert(tk.END, f"{time.strftime('%H:%M:%S')} - {message}\n")
            self.log_text.see(tk.END)
            self.log_text.configure(state=tk.DISABLED)
                
                # 记录到文件日志
                self.logger.info(message)
        except Exception as e:
            print(f"Log error: {e}")
                # 至少保证文件日志记录
                try:
                    self.logger.error(f"UI log error: {e}, original message: {message}")
                except:
                    print(f"Critical logging failure: {e}")
        
        # 确保在主线程中执行UI操作
        if threading.current_thread() is threading.main_thread():
            _log()
        else:
            # 后台线程：安排到主线程执行
            self.root.after(0, _log)
    
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
    
    def monitor_clipboard(self):
        """监控剪贴板内容 - 增强稳定性版本"""
        self.log("Clipboard monitoring thread started...")
        
        try:
            # 初始化剪贴板内容
        try:
            self.previous_content = pyperclip.paste()
            except Exception as e:
                self.log(f"Initial clipboard read failed: {str(e)}")
                self.previous_content = ""
            
            consecutive_errors = 0
            max_consecutive_errors = 10
            
            while self.running:
                try:
                    # 检查UI是否还存在
                    if not hasattr(self, 'root') or not self.root.winfo_exists():
                        self.log("UI no longer exists, stopping clipboard monitoring")
                        break
                    
                    # 尝试读取剪贴板
                try:
                    current_content = pyperclip.paste()
                        consecutive_errors = 0  # 重置错误计数
                    except Exception as clipboard_error:
                        consecutive_errors += 1
                        if consecutive_errors <= 3:  # 只记录前几次错误
                            self.log(f"Clipboard access error #{consecutive_errors}: {str(clipboard_error)}")
                        
                        if consecutive_errors >= max_consecutive_errors:
                            self.log(f"Too many consecutive clipboard errors ({consecutive_errors}), stopping monitoring")
                            break
                            
                        time.sleep(1.0)  # 等待更长时间再重试
                        continue
                    
                    # 检查内容是否有变化
                    if current_content != self.previous_content and current_content is not None:
                        try:
                        content_preview = current_content[:30] + "..." if len(current_content) > 30 else current_content
                        self.log(f"New content detected: {content_preview}")
                        
                            # 安全地更新剪贴板显示
                            try:
                                if hasattr(self, 'root') and self.root.winfo_exists():
                        self.root.after(0, self.update_clipboard_display)
                            except Exception as ui_error:
                                self.log(f"UI update error: {str(ui_error)}")
                        
                            # 安全地刷新Excel单元格信息
                            try:
                                if self.excel_app and hasattr(self, 'excel_app'):
                            self.refresh_current_cell()
                            except Exception as excel_error:
                                self.log(f"Excel refresh error: {str(excel_error)}")
                        
                        self.previous_content = current_content
                        
                        except Exception as processing_error:
                            self.log(f"Content processing error: {str(processing_error)}")
                        
                except Exception as loop_error:
                    consecutive_errors += 1
                    self.log(f"Monitoring loop error #{consecutive_errors}: {str(loop_error)}")
                    
                    if consecutive_errors >= max_consecutive_errors:
                        self.log("Too many loop errors, stopping monitoring")
                        break
                
                # 短暂暂停，但要检查运行状态
                for _ in range(10):  # 0.5秒总计，但分成小块检查
                    if not self.running:
                        break
                    time.sleep(0.05)
                    
        except Exception as e:
            self.log(f"Critical monitoring thread error: {str(e)}")
        finally:
            self.log("Clipboard monitoring thread ended")
            # 确保线程标记为停止
            try:
                if hasattr(self, 'monitor_thread'):
                    self.monitor_thread = None
            except:
                pass
    
    def start_monitoring(self):
        """开始监控"""
        try:
            self.logger.info("Starting monitoring process")
            
            # 检查是否连接到Excel
            if not self.excel_app:
                self.logger.warning("Excel not connected when starting monitoring")
                if not messagebox.askyesno("Warning", "Not connected to Excel. Do you want to connect now?"):
                    if not messagebox.askyesno("Warning", "Continue without Excel connection? The program will use keyboard shortcuts."):
                        self.logger.info("User cancelled monitoring start - no Excel connection")
                        return
                else:
                    if not self.connect_to_excel():
                        self.logger.error("Failed to connect to Excel during monitoring start")
                        return
            
            self.running = True
            self.status_label.config(text="Running")
            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
            
            # Disable Excel target setting while monitoring
            self.set_excel_button.config(state=tk.DISABLED)
            
            self.log("Starting monitoring...")
            self.logger.info(f"Monitoring started with settings: reminder_time={self.reminder_time}, row_skip={self.row_skip_count}, max_files={self.max_files_per_cell}")
            
            if self.target_excel != "Not specified":
                self.log(f"Target Excel: {self.target_excel}")
                
            # 显示当前选中的单元格
            if self.excel_app:
                self.refresh_current_cell()
                self.log(f"Current selected cell: {self.current_cell}")
                self.logger.info(f"Current cell: {self.current_cell}")
                
            # 更新剪贴板显示
            self.update_clipboard_display()
            
            # 启动监控线程
            self.monitor_thread = threading.Thread(target=self.monitor_clipboard)
            self.monitor_thread.daemon = True
            self.monitor_thread.start()
            self.logger.info("Clipboard monitoring thread started")
            
            # 提示用户操作方法
            self.log("Auto-paste mode: Content will be pasted automatically when detected")
            messagebox.showinfo("Monitoring Started", 
                               "Auto-paste mode enabled.\n\n"
                               "When matching content is detected, it will be automatically pasted to the current Excel cell with a notification.\n"
                               "If the cell already has content, new content will be added on a new line.\n\n"
                               "No action required - the process is fully automated.")
            
        except Exception as e:
            error_msg = f"Start monitoring error: {str(e)}"
            self.log(error_msg)
            self.logger.error(error_msg, exc_info=True)
            messagebox.showerror("Error", f"Failed to start monitoring: {str(e)}")
            self.running = False
    
    def stop_monitoring(self):
        """停止监控 - 增强资源清理"""
        try:
            self.log("Stopping monitoring...")
            self.running = False
            
            # 等待监控线程结束
            if hasattr(self, 'monitor_thread') and self.monitor_thread and self.monitor_thread.is_alive():
                try:
                    self.log("Waiting for monitoring thread to stop...")
                    self.monitor_thread.join(timeout=2.0)  # 最多等待2秒
                    if self.monitor_thread.is_alive():
                        self.log("Warning: Monitoring thread did not stop gracefully")
                    else:
                        self.log("Monitoring thread stopped successfully")
                except Exception as e:
                    self.log(f"Error stopping monitoring thread: {str(e)}")
            
            # 清理线程引用
            if hasattr(self, 'monitor_thread'):
                self.monitor_thread = None
            
            # 停止单元格检查线程
            if hasattr(self, 'cell_check_thread') and self.cell_check_thread and self.cell_check_thread.is_alive():
                try:
                    self.log("Waiting for cell check thread to stop...")
                    self.cell_check_thread.join(timeout=2.0)
                except Exception as e:
                    self.log(f"Error stopping cell check thread: {str(e)}")
            
            # 清理线程引用
            if hasattr(self, 'cell_check_thread'):
                self.cell_check_thread = None
            
            # 停止健康检查
            if hasattr(self, '_health_check_job') and self._health_check_job:
                try:
                    self.root.after_cancel(self._health_check_job)
                    self._health_check_job = None
                except:
                    pass
            
            # 安全停止COM对象
            try:
                pythoncom.CoUninitialize()
            except Exception as e:
                self.log(f"COM cleanup warning: {str(e)}")
            
            # 更新界面
            self.status_label.config(text="Stopped")
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            
            # Re-enable Excel target setting
            self.set_excel_button.config(state=tk.NORMAL)
            
            # 关闭任何打开的通知窗口
            if self.confirmation_dialog and self.confirmation_dialog.winfo_exists():
                self.confirmation_dialog.destroy()
                self.confirmation_dialog = None
                
            self.logger.info("Monitoring stopped with enhanced cleanup")
            self.log("Monitoring stopped successfully")
        except Exception as e:
            self.log(f"Stop monitoring error: {str(e)}")
    
    def on_closing(self):
        """关闭窗口处理"""
        try:
            # 保存当前设置
            self.save_current_settings()
            
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
            self.logger.error(f"Error closing application: {str(e)}")
            self.root.destroy()
    
    def save_current_settings(self):
        """保存当前UI中的设置"""
        try:
            # 更新当前设置值
            self.reminder_time = int(self.reminder_time_var.get())
            self.row_skip_count = int(self.row_skip_var.get())
            
            # 更新max_files_per_cell
            max_files_str = self.max_files_var.get().strip()
            if max_files_str:
                try:
                    self.max_files_per_cell = int(max_files_str)
                except ValueError:
                    self.max_files_per_cell = None
            else:
                self.max_files_per_cell = None
            
            # 更新忽略最大数量设置
            self.ignore_max_move = self.ignore_max_move_button.cget("text").endswith("ON")

            # 保存设置
            self.save_settings()
            self.log("Settings saved on exit")
            
        except Exception as e:
            self.logger.error(f"Error saving settings on exit: {str(e)}")
            self.log(f"Warning: Failed to save settings: {str(e)}")

    def toggle_auto_move(self):
        """切换自动移动到下一行的功能"""
        self.auto_move_next = not self.auto_move_next
        button_text = "Auto Move Next: ON" if self.auto_move_next else "Auto Move Next: OFF"
        self.auto_move_button.config(text=button_text)
        self.log(f"Auto move to next row: {'Enabled' if self.auto_move_next else 'Disabled'}")

    def toggle_phone_notification(self):
        """切换手机通知功能"""
        self.phone_notification_enabled = not self.phone_notification_enabled
        button_text = "Phone Alert: ON" if self.phone_notification_enabled else "Phone Alert: OFF"
        self.phone_notify_button.config(text=button_text)
        self.log(f"Phone notification: {'Enabled' if self.phone_notification_enabled else 'Disabled'}")

    def toggle_ignore_max_move(self):
        """切换忽略最大数量后总是下移的行为（仅影响移动，不影响粘贴容量校验）"""
        try:
            self.ignore_max_move = not self.ignore_max_move
            self.ignore_max_move_button.config(text=("Ignore: ON" if self.ignore_max_move else "Ignore: OFF"))
            self.log(f"Ignore max for move: {'Enabled' if self.ignore_max_move else 'Disabled'}")
        except Exception:
            pass

    def send_phone_alert(self):
        """发送手机震动提醒"""
        if not self.phone_notification_enabled:
            return
            
        topic = self.ntfy_topic_var.get().strip()
        if not topic:
            self.log("Phone alert failed: No topic specified")
            return
        
        def send_async():
            try:
                url = f"https://ntfy.sh/{topic}"
                
                # 发送3条连续通知以增强震动效果
                messages = [
                    ("URGENT: AutoCopy Alert! Please continue working immediately!", "WORK REMINDER"),
                    ("ATTENTION: Still no activity detected!", "SECOND ALERT"),
                    ("FINAL WARNING: Resume work now!", "LAST NOTICE")
                ]
                
                success_count = 0
                for i, (message, title) in enumerate(messages):
                    response = requests.post(
                        url, 
                        data=message,
                        headers={
                            "Title": title,
                            "Priority": "urgent",
                            "Tags": "warning"
                        },
                        timeout=5
                    )
                    if response.status_code == 200:
                        success_count += 1
                    
                    # 每条消息间隔0.2秒，更快更密集
                    if i < len(messages) - 1:
                        time.sleep(0.2)
                
                if success_count > 0:
                    self.log(f"Phone alert sent successfully ({success_count}/3 messages)")
                else:
                    self.log("All phone alerts failed")
                    
            except Exception as e:
                self.log(f"Phone alert failed: {str(e)}")
        
        # 异步发送，不阻塞主线程
        threading.Thread(target=send_async, daemon=True).start()

    def test_phone_alert(self):
        """测试手机提醒功能"""
        if not self.phone_notification_enabled:
            messagebox.showinfo("Info", "Please enable phone notification and set topic name first")
            return
            
        topic = self.ntfy_topic_var.get().strip()
        if not topic:
            messagebox.showinfo("Info", "Please set ntfy topic name first")
            return
            
        self.send_phone_alert()
        self.log("Phone alert test sent")

    def start_activity_monitoring(self):
        """开始监控用户活动"""
        self.activity_monitor_active = True
        self.last_activity_time = time.time()
        self.activity_detected = False
        
        # 获取初始鼠标位置
        cursor = wintypes.POINT()
        windll.user32.GetCursorPos(ctypes.byref(cursor))
        self.last_mouse_pos = (cursor.x, cursor.y)
        
        # 启动活动检查定时器
        self._check_activity()

    def _check_activity(self):
        """检查活动状态"""
        if not self.activity_monitor_active:
            return
            
        try:
            # 检查鼠标位置
            cursor = wintypes.POINT()
            windll.user32.GetCursorPos(ctypes.byref(cursor))
            current_pos = (cursor.x, cursor.y)
            
            # 如果鼠标位置发生变化
            if current_pos != self.last_mouse_pos:
                self.log("Mouse movement detected")
                self.activity_detected = True
                self.last_mouse_pos = current_pos
            
            # 检查键盘状态
            for key in range(0x30, 0x5A):  # 检查常用按键
                if windll.user32.GetAsyncKeyState(key) & 0x8000:
                    self.log("Keyboard activity detected")
                    self.activity_detected = True
                    break
            
            if self.activity_detected:
                self.log("Activity detected - stopping monitoring")
                self.stop_activity_monitoring()
                # 关闭弹窗（如果存在）
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
                # 显示提醒后停止监控
                self.stop_activity_monitoring()
            else:
                # 继续检查
                self.root.after(100, self._check_activity)
                
        except Exception as e:
            self.log(f"Activity check error: {str(e)}")
            # 继续检查
            self.root.after(100, self._check_activity)

    def stop_activity_monitoring(self):
        """停止监控用户活动"""
        self.activity_monitor_active = False
        self.activity_detected = False
        # 取消定时器
        if self.reminder_timer:
            self.root.after_cancel(self.reminder_timer)
            self.reminder_timer = None

    def show_reminder_dialog(self):
        if self.reminder_dialog and self.reminder_dialog.winfo_exists():
            return

        # 发送手机提醒
        self.send_phone_alert()

        self.reminder_dialog = tk.Toplevel(self.root)
        self.reminder_dialog.title("Activity Reminder")
        self.reminder_dialog.attributes('-topmost', True)
        
        # 获取所有显示器的信息
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
        except:
            # 如果获取显示器信息失败，使用主显示器
            monitors = [{
                'left': 0,
                'top': 0,
                'width': self.root.winfo_screenwidth(),
                'height': self.root.winfo_screenheight()
            }]

        # 选择要显示提醒的显示器（默认使用主显示器）
        target_monitor = monitors[0]
        
        # 设置窗口大小和位置
        window_width = 1200
        window_height = 800
        x = target_monitor['left'] + (target_monitor['width'] - window_width) // 2
        y = target_monitor['top'] + (target_monitor['height'] - window_height) // 2
        self.reminder_dialog.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # 初始背景色
        self._reminder_bg_colors = ["#FFE4E1", "#FF0000"]
        self._reminder_bg_index = 0
        self._reminder_flash_job = None

        self.reminder_dialog.configure(bg=self._reminder_bg_colors[self._reminder_bg_index])

        frame = tk.Frame(self.reminder_dialog, bg=self._reminder_bg_colors[self._reminder_bg_index], padx=30, pady=30)
        frame.pack(fill=tk.BOTH, expand=True)

        warning_label = tk.Label(frame, text="⚠️", font=("Arial", 72), bg=self._reminder_bg_colors[self._reminder_bg_index])
        warning_label.pack(pady=(0, 20))

        message = "NO ACTIVITY DETECTED!\n\nPlease continue your work or close this window."
        text_label = tk.Label(frame, text=message, font=("Arial", 16, "bold"),
                              justify=tk.CENTER, bg=self._reminder_bg_colors[self._reminder_bg_index], fg="#8B0000")
        text_label.pack(pady=20)

        close_button = tk.Button(frame, text="CLOSE", command=self.reminder_dialog.destroy,
                                 font=("Arial", 12, "bold"), bg="#FF6B6B", fg="white",
                                 relief=tk.RAISED, padx=20, pady=10)
        close_button.pack(pady=20)

        # 绑定活动事件到弹窗
        self.reminder_dialog.bind("<Motion>", lambda e: self.reminder_dialog.destroy())
        self.reminder_dialog.bind("<Key>", lambda e: self.reminder_dialog.destroy())
        self.reminder_dialog.bind("<Button>", lambda e: self.reminder_dialog.destroy())

        # 启动闪烁
        self._reminder_flash_bg()

    def _reminder_flash_bg(self):
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
            except:
                pass
        self._reminder_flash_job = self.reminder_dialog.after(500, self._reminder_flash_bg)

def main():
    try:
        root = tk.Tk()
        app = AutoCopyApp(root)
        # 设置图标
        try:
            root.iconbitmap("clipboard.ico")
        except:
            pass
        root.geometry("700x900")  # 启动时更大，确保所有控件显示
        root.minsize(600, 800)
        
        # 设置全局异常处理器
        def handle_exception(exc_type, exc_value, exc_traceback):
            """全局异常处理器"""
            if issubclass(exc_type, KeyboardInterrupt):
                # 允许Ctrl+C退出
                sys.__excepthook__(exc_type, exc_value, exc_traceback)
                return
            
            error_msg = f"Unhandled exception: {exc_type.__name__}: {str(exc_value)}"
            try:
                app.logger.critical(error_msg, exc_info=(exc_type, exc_value, exc_traceback))
                app.log(f"CRITICAL ERROR: {error_msg}")
            except:
                print(f"CRITICAL ERROR: {error_msg}")
                traceback.print_exception(exc_type, exc_value, exc_traceback)
        
        # 只在没有调试器时设置全局异常处理
        import sys
        if not sys.gettrace():
            sys.excepthook = handle_exception
        
        # 增强的主循环错误处理
        loop_error_count = 0
        max_loop_errors = 5
        
        while True:
            try:
                root.mainloop()
                break  # 正常退出
            except Exception as e:
                loop_error_count += 1
                error_msg = f"UI Error #{loop_error_count}: {str(e)}"
                
                try:
                    app.logger.error(f"Main loop error #{loop_error_count}: {str(e)}", exc_info=True)
                    app.log(error_msg)
                except:
                    print(error_msg)
                    traceback.print_exc()
                
                # 如果连续错误太多，强制退出
                if loop_error_count >= max_loop_errors:
                    try:
                        messagebox.showerror("Too Many Errors", 
                                           f"Too many consecutive UI errors ({loop_error_count}). The application will close.")
                    except:
                        pass
                    break
                
                # 显示错误但继续运行
                try:
                    response = messagebox.askyesno("Application Error", 
                                                 f"{error_msg}\n\nThe application will attempt to continue.\n\nDo you want to continue running?")
                    if not response:
                        break
                except:
                    # 如果连对话框都无法显示，等待一下再尝试
                    import time
                    time.sleep(1)
                    continue
                    
    except Exception as e:
        error_msg = f"Critical Application Error: {str(e)}"
        try:
            # 尝试记录到文件
            with open("crash_log.txt", "a", encoding="utf-8") as f:
                f.write(f"\n{time.strftime('%Y-%m-%d %H:%M:%S')} - CRASH: {error_msg}\n")
                traceback.print_exc(file=f)
        except:
            pass
        
        try:
            messagebox.showerror("Critical Application Error", 
                               f"{error_msg}\n\nThe application will now close.\n\nCheck crash_log.txt for details.")
        except:
            pass
        
        print(f"Critical Error: {error_msg}")
        traceback.print_exc()

if __name__ == "__main__":
    main()