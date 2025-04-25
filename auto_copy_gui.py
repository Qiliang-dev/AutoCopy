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
        self.format_var = tk.StringVar(value=r'^20\d{2}_\d{2}_\d{2}_\d{6}_DA\d{5}_A')
        format_entry = ttk.Entry(format_frame, textvariable=self.format_var, width=40)
        format_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 匹配状态显示
        ttk.Label(format_frame, text="Match Status:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.match_status_label = ttk.Label(format_frame, text="Not checked")
        self.match_status_label.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
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
    
    def paste_to_excel(self):
        """手动将剪贴板内容粘贴到Excel"""
        if not self.excel_app:
            messagebox.showwarning("Warning", "Not connected to Excel. Please connect first.")
            return
            
        try:
            # 刷新当前单元格
            self.refresh_current_cell()
            
            # 获取剪贴板内容
            content = pyperclip.paste()
            
            # 直接设置Excel单元格的值
            self.excel_app.ActiveCell.Value = content
            
            self.log(f"Content pasted to cell {self.current_cell}")
            
            # 更新显示
            self.update_clipboard_display()
        except Exception as e:
            error_msg = f"Failed to paste: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("Paste Error", error_msg)
    
    def update_clipboard_display(self):
        """更新剪贴板内容显示"""
        try:
            content = pyperclip.paste()
            
            # 只有在内容变化时更新
            if content != self.clipboard_content:
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
                    
                    # 如果正在监控并且匹配成功，显示确认对话框
                    if self.running:
                        self.show_paste_confirmation(content)
                else:
                    self.match_status_label.config(text="Does Not Match", foreground="red")
            
            # 设置剪贴板检查定时器 - 每秒更新一次
            self.root.after(1000, self.update_clipboard_display)
            
        except Exception as e:
            self.log(f"Error updating clipboard display: {str(e)}")
            # 继续检查，即使有错误
            self.root.after(1000, self.update_clipboard_display)
    
    def schedule_cell_check(self):
        """定期检查Excel单元格"""
        if self.excel_app:
            self.refresh_current_cell()
        
        # 每100毫秒检查一次单元格
        self.excel_check_timer = self.root.after(100, self.schedule_cell_check)
    
    def show_paste_confirmation(self, content):
        """显示粘贴确认对话框，按回车键确认粘贴"""
        # 如果已经有弹窗，先关闭
        if self.confirmation_dialog is not None and self.confirmation_dialog.winfo_exists():
            self.confirmation_dialog.destroy()
            
        # 创建新的确认对话框
        self.confirmation_dialog = tk.Toplevel(self.root)
        self.confirmation_dialog.title("内容检测到 - 按回车键粘贴")
        self.confirmation_dialog.geometry("500x200")
        self.confirmation_dialog.resizable(False, False)
        
        # 确保对话框总是在前台
        self.confirmation_dialog.attributes('-topmost', True)
        
        # 设置窗口为模态，阻止其他窗口交互
        self.confirmation_dialog.grab_set()
        
        # 制作闪烁效果的背景
        frame = ttk.Frame(self.confirmation_dialog, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # 内容预览
        preview_label = ttk.Label(frame, text="内容匹配成功!", font=("Arial", 12, "bold"))
        preview_label.pack(pady=(0, 10))
        
        # 内容显示
        content_text = scrolledtext.ScrolledText(frame, height=4, width=50, wrap=tk.WORD)
        content_text.pack(fill=tk.BOTH, expand=True, pady=5)
        content_text.insert(tk.END, content)
        content_text.config(state=tk.DISABLED)
        
        # 单元格信息
        cell_info = ttk.Label(frame, text=f"目标单元格: {self.current_cell}")
        cell_info.pack(pady=5)
        
        # 确认按钮
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        paste_button = ttk.Button(button_frame, text="粘贴 (回车)", command=lambda: self._confirm_paste(content))
        paste_button.pack(side=tk.LEFT, padx=5)
        
        cancel_button = ttk.Button(button_frame, text="取消 (Esc)", command=lambda: self.confirmation_dialog.destroy())
        cancel_button.pack(side=tk.RIGHT, padx=5)
        
        # 焦点设置到粘贴按钮并绑定回车键
        paste_button.focus_set()
        
        # 绑定按键
        self.confirmation_dialog.bind("<Return>", lambda event: self._confirm_paste(content))
        self.confirmation_dialog.bind("<Escape>", lambda event: self.confirmation_dialog.destroy())
        
        # 闪烁效果
        self._blink_background(frame, 5)
    
    def _blink_background(self, widget, times):
        """创建闪烁效果以引起注意"""
        if times <= 0:
            return
            
        # 交替颜色
        current_bg = widget.cget("background")
        highlight_bg = "#90EE90"  # 浅绿色
        
        widget.configure(background=highlight_bg)
        self.root.after(300, lambda: widget.configure(background=current_bg))
        self.root.after(600, lambda: self._blink_background(widget, times-1))
    
    def _confirm_paste(self, content):
        """确认粘贴操作"""
        if self.confirmation_dialog and self.confirmation_dialog.winfo_exists():
            self.confirmation_dialog.destroy()
            self.paste_to_excel()
    
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
            self.log("Monitoring mode: Press Enter when prompted to paste content")
            messagebox.showinfo("Monitoring Started", 
                               "当检测到匹配内容时，确认对话框会自动弹出。\n"
                               "只需按下回车键即可将内容粘贴到Excel。")
            
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
            
            # 关闭任何打开的确认对话框
            if self.confirmation_dialog and self.confirmation_dialog.winfo_exists():
                self.confirmation_dialog.destroy()
                
            self.log("Monitoring stopped")
        except Exception as e:
            self.log(f"Stop monitoring error: {str(e)}")
    
    def on_closing(self):
        """关闭窗口处理"""
        try:
            if self.running:
                if messagebox.askokcancel("Exit", "Monitoring is still running. Are you sure you want to exit?"):
                    self.stop_monitoring()
                    self.running = False
                    
                    # 取消定时器
                    if self.excel_check_timer:
                        self.root.after_cancel(self.excel_check_timer)
                    
                    # 释放Excel资源
                    if self.excel_app:
                        self.excel_app = None
                        
                    self.root.destroy()
                    
            else:
                # 取消定时器
                if self.excel_check_timer:
                    self.root.after_cancel(self.excel_check_timer)
                
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