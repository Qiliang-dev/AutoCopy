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

class AutoCopyApp:
    def __init__(self, root):
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
        
        self.root = root
        self.root.title("AutoCopy Tool")
        self.root.geometry("650x550")  # 增加窗口尺寸，以容纳更多控件
        self.root.resizable(True, True)  # 允许用户调整窗口大小
        
        # 添加关闭窗口处理
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Setup UI after initializing all attributes
        self.setup_ui()
        
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
        
        # 控制按钮区域 - 移到状态区域下方，确保可见
        control_frame = ttk.LabelFrame(main_frame, text="Controls", padding="10")
        control_frame.pack(fill=tk.X, pady=5)
        
        # 使用Grid布局代替Pack，确保按钮正确排列
        self.start_button = ttk.Button(control_frame, text="Start Monitoring", command=self.start_monitoring)
        self.start_button.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        
        self.stop_button = ttk.Button(control_frame, text="Stop Monitoring", command=self.stop_monitoring, state=tk.DISABLED)
        self.stop_button.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 清除日志按钮
        self.clear_log_button = ttk.Button(control_frame, text="Clear Log", command=self.clear_log)
        self.clear_log_button.grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        
        # 更新剪贴板按钮
        self.update_clipboard_button = ttk.Button(control_frame, text="Refresh Clipboard", command=self.update_clipboard_display)
        self.update_clipboard_button.grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)
        
        # 退出按钮
        self.exit_button = ttk.Button(control_frame, text="Exit", command=self.on_closing)
        self.exit_button.grid(row=0, column=4, padx=5, pady=5, sticky=tk.E)
        
        # 剪贴板内容显示区域
        clipboard_frame = ttk.LabelFrame(main_frame, text="Current Clipboard Content", padding="10")
        clipboard_frame.pack(fill=tk.X, pady=5)
        
        self.clipboard_text = scrolledtext.ScrolledText(clipboard_frame, height=4, width=70, wrap=tk.WORD)
        self.clipboard_text.pack(fill=tk.X, expand=True)
        self.clipboard_text.insert(tk.END, "(No content)")
        self.clipboard_text.config(state=tk.DISABLED)
        
        # 匹配格式控制
        format_frame = ttk.LabelFrame(main_frame, text="Pattern Settings", padding="10")
        format_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(format_frame, text="Pattern:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.format_var = tk.StringVar(value=r'^20\d{2}_\d{2}_\d{2}_\d{6}')
        format_entry = ttk.Entry(format_frame, textvariable=self.format_var, width=40)
        format_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 添加防重复粘贴时间间隔设置
        ttk.Label(format_frame, text="Duplicate Protection (s):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.duplicate_time_var = tk.StringVar(value="3")  # 默认3秒
        duplicate_entry = ttk.Spinbox(format_frame, from_=1, to=10, width=5, textvariable=self.duplicate_time_var)
        duplicate_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Label(format_frame, text="Prevention of duplicate pasting within the specified seconds", 
                  font=("Arial", 8)).grid(row=1, column=2, sticky=tk.W, padx=5, pady=5)
        
        # 匹配状态显示
        ttk.Label(format_frame, text="Match Status:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.match_status_label = ttk.Label(format_frame, text="Not checked")
        self.match_status_label.grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)
        
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
            # 刷新当前单元格
            self.refresh_current_cell()
            
            # 获取剪贴板内容
            content = pyperclip.paste()
            
            # 检查当前单元格是否已有内容
            current_value = self.excel_app.ActiveCell.Value
            if current_value:
                # 如果单元格已有内容，则在现有内容后添加换行和新内容
                self.log("Cell already has content, appending with new line")
                # Excel中的换行符是Chr(10)
                new_value = f"{current_value}{chr(10)}{content}"
                self.excel_app.ActiveCell.Value = new_value
                self.log(f"Content appended to cell {self.current_cell}")
            else:
                # 单元格为空，直接设置值
                self.excel_app.ActiveCell.Value = content
                self.log(f"Content set to cell {self.current_cell}")
            
            # 更新显示
            self.update_clipboard_display()
            return True  # 返回成功状态
        except Exception as e:
            error_msg = f"Failed to paste: {str(e)}"
            self.log(error_msg)
            if show_error_dialog:
                messagebox.showerror("Paste Error", error_msg)
            return False  # 返回失败状态
    
    def update_clipboard_display(self):
        """更新剪贴板内容显示"""
        try:
            content = pyperclip.paste()
            
            # 只有在内容变化时更新 - 使用更严格的比较并防止重复触发
            if content != self.clipboard_content and content.strip() != "":
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
                
                # 更新文本显示
                self.clipboard_text.config(state=tk.NORMAL)
                self.clipboard_text.delete(1.0, tk.END)
                
                # 限制显示长度，防止过长内容
                if len(content) > 500:
                    display_content = content[:500] + "... (content truncated)"
                else:
                    display_content = content
                    
                self.clipboard_text.insert(tk.END, display_content)
                self.clipboard_text.config(state=tk.DISABLED)
                
                # 检查是否匹配格式
                match_result = self.is_valid_format(content)
                if match_result:
                    self.match_status_label.config(text="Matches Pattern", foreground="green")
                    
                    # 如果正在监控并且匹配成功且不是重复内容，显示通知并自动粘贴
                    if self.running and not is_duplicate:
                        # 保存此次操作数据，用于防止重复
                        self.last_pasted_content = content
                        self.last_paste_time = current_time
                        
                        # 延迟一小段时间后执行粘贴操作，给UI时间更新
                        self.root.after(100, lambda: self.auto_paste_with_notification(content))
                else:
                    self.match_status_label.config(text="Does Not Match", foreground="red")
            
            # 设置剪贴板检查定时器 - 每秒更新一次
            self.root.after(1000, self.update_clipboard_display)
            
        except Exception as e:
            self.log(f"Error updating clipboard display: {str(e)}")
            # 继续检查，即使有错误
            self.root.after(1000, self.update_clipboard_display)
    
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
            self.refresh_current_cell()
        
        # 每100毫秒检查一次单元格
        self.excel_check_timer = self.root.after(100, self.schedule_cell_check)
    
    def refresh_current_cell(self):
        """刷新当前选中的单元格"""
        if self.excel_app:
            try:
                cell_address = self.excel_app.ActiveCell.Address
                if cell_address != self.current_cell:
                    self.current_cell = cell_address
                    self.cell_label.config(text=cell_address)
                    # 仅在监控状态下记录单元格变化
                    if self.running:
                        self.log(f"Cell selection changed: {cell_address}")
                return True
            except Exception as e:
                self.log(f"Error refreshing cell: {str(e)}")
                return False
        return False
    
    def connect_to_excel(self):
        """连接到Excel应用程序"""
        try:
            # 初始化COM线程 - 主UI线程
            pythoncom.CoInitialize()
            
            # 获取Excel应用程序实例
            self.excel_app = win32com.client.GetActiveObject("Excel.Application")
            
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
            
            # 启动单元格监控 - 现在使用定时器代替线程
            self.excel_cell_monitor_active = True
            
            # 确保可以粘贴
            self.paste_button.config(state=tk.NORMAL)
            
            return True
            
        except Exception as e:
            self.log(f"Failed to connect to Excel: {str(e)}")
            messagebox.showerror("Excel Connection Error", 
                                 "Failed to connect to Excel. Please make sure Excel is open and try again.")
            traceback.print_exc()
            return False
    
    def monitor_excel_cell(self):
        """监控Excel单元格变化 - 此方法现已弃用，使用定时器代替"""
        try:
            # 初始化COM线程
            pythoncom.CoInitialize()
            
            last_cell = self.current_cell
            
            while self.running or self.excel_app:
                try:
                    # 检查Excel是否还在运行
                    if not self.excel_app:
                        break
                        
                    # 获取当前单元格
                    current_cell = self.excel_app.ActiveCell.Address
                    
                    # 如果单元格改变了，更新显示
                    if current_cell != last_cell:
                        self.current_cell = current_cell
                        
                        # 使用线程安全的方式更新UI
                        self.root.after(0, lambda: self.cell_label.config(text=current_cell))
                        self.root.after(0, lambda: self.log(f"Cell selection changed: {current_cell}"))
                        
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
        """添加消息到日志区域"""
        try:
            self.log_text.configure(state=tk.NORMAL)
            self.log_text.insert(tk.END, f"{time.strftime('%H:%M:%S')} - {message}\n")
            self.log_text.see(tk.END)
            self.log_text.configure(state=tk.DISABLED)
        except Exception as e:
            print(f"Log error: {e}")
    
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
        """监控剪贴板内容"""
        try:
            self.previous_content = pyperclip.paste()
            self.log("Clipboard monitoring started...")
            
            while self.running:
                try:
                    current_content = pyperclip.paste()
                    
                    # 检查剪贴板内容是否有变化
                    if current_content != self.previous_content:
                        content_preview = current_content[:30] + "..." if len(current_content) > 30 else current_content
                        self.log(f"New content detected: {content_preview}")
                        
                        # 更新剪贴板显示
                        self.root.after(0, self.update_clipboard_display)
                        
                        # 每次粘贴前刷新当前单元格信息
                        if self.excel_app:
                            self.refresh_current_cell()
                        
                        # 注意：现在使用update_clipboard_display方法显示确认对话框，不再在此处执行自动粘贴
                        self.previous_content = current_content
                        
                except Exception as e:
                    self.log(f"Monitoring loop error: {str(e)}")
                
                # 短暂暂停以减少CPU使用
                time.sleep(0.5)
        except Exception as e:
            self.log(f"Monitoring thread error: {str(e)}")
    
    def start_monitoring(self):
        """开始监控"""
        try:
            # 检查是否连接到Excel
            if not self.excel_app:
                if not messagebox.askyesno("Warning", "Not connected to Excel. Do you want to connect now?"):
                    if not messagebox.askyesno("Warning", "Continue without Excel connection? The program will use keyboard shortcuts."):
                        return
                else:
                    if not self.connect_to_excel():
                        return
            
            self.running = True
            self.status_label.config(text="Running")
            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
            
            # Disable Excel target setting while monitoring
            self.set_excel_button.config(state=tk.DISABLED)
            
            self.log("Starting monitoring...")
            if self.target_excel != "Not specified":
                self.log(f"Target Excel: {self.target_excel}")
                
            # 显示当前选中的单元格
            if self.excel_app:
                self.refresh_current_cell()
                self.log(f"Current selected cell: {self.current_cell}")
                
            # 更新剪贴板显示
            self.update_clipboard_display()
            
            # 启动监控线程
            self.monitor_thread = threading.Thread(target=self.monitor_clipboard)
            self.monitor_thread.daemon = True
            self.monitor_thread.start()
            
            # 提示用户操作方法
            self.log("Auto-paste mode: Content will be pasted automatically when detected")
            messagebox.showinfo("Monitoring Started", 
                               "Auto-paste mode enabled.\n\n"
                               "When matching content is detected, it will be automatically pasted to the current Excel cell with a notification.\n"
                               "If the cell already has content, new content will be added on a new line.\n\n"
                               "No action required - the process is fully automated.")
            
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
            
            # Re-enable Excel target setting
            self.set_excel_button.config(state=tk.NORMAL)
            
            # 关闭任何打开的通知窗口
            if self.confirmation_dialog and self.confirmation_dialog.winfo_exists():
                self.confirmation_dialog.destroy()
                self.confirmation_dialog = None
                
            self.log("Monitoring stopped")
        except Exception as e:
            self.log(f"Stop monitoring error: {str(e)}")
    
    def on_closing(self):
        """关闭窗口处理"""
        try:
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

def main():
    try:
        root = tk.Tk()
        app = AutoCopyApp(root)
        # 设置图标
        try:
            # 如果有图标文件可以使用
            root.iconbitmap("clipboard.ico")
        except:
            pass
        root.mainloop()
    except Exception as e:
        messagebox.showerror("Application Error", f"An error occurred: {str(e)}")
        print(f"Error: {str(e)}")
        traceback.print_exc()

if __name__ == "__main__":
    main()