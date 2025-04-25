import pyperclip
import pyautogui
import re
import time
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import win32gui
import win32process
import psutil
import sys
import traceback

class AutoCopyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AutoCopy工具")
        self.root.geometry("500x300")
        self.root.resizable(False, False)
        
        # 设置图标
        try:
            self.root.iconbitmap("clipboard.ico")  # 如果有图标文件，可以使用
        except:
            pass  # 忽略图标加载错误
            
        # 添加关闭窗口处理
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.setup_ui()
        
        self.running = False
        self.monitor_thread = None
        self.previous_content = ""
        
        # Excel相关
        self.excel_file_name = "未检测到Excel"
        self.check_excel_timer = None
        
    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 状态显示区域
        status_frame = ttk.LabelFrame(main_frame, text="状态", padding="10")
        status_frame.pack(fill=tk.X, pady=5)
        
        # Excel文件信息
        ttk.Label(status_frame, text="当前Excel文件:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.excel_file_label = ttk.Label(status_frame, text=self.excel_file_name)
        self.excel_file_label.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        # 运行状态
        ttk.Label(status_frame, text="监控状态:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.status_label = ttk.Label(status_frame, text="未启动")
        self.status_label.grid(row=1, column=1, sticky=tk.W, pady=5)
        
        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="日志", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.log_text = tk.Text(log_frame, height=10, width=50, wrap=tk.WORD)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        # 控制按钮
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=10)
        
        self.start_button = ttk.Button(control_frame, text="开始监控", command=self.start_monitoring)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_button = ttk.Button(control_frame, text="停止监控", command=self.stop_monitoring, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        # 清除日志按钮
        self.clear_log_button = ttk.Button(control_frame, text="清除日志", command=self.clear_log)
        self.clear_log_button.pack(side=tk.LEFT, padx=5)
        
        # 退出按钮
        self.exit_button = ttk.Button(control_frame, text="退出", command=self.on_closing)
        self.exit_button.pack(side=tk.RIGHT, padx=5)
        
        # 重置匹配格式控制
        format_frame = ttk.LabelFrame(main_frame, text="匹配格式", padding="5")
        format_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(format_frame, text="格式模式:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.format_var = tk.StringVar(value=r'^20\d{2}_\d{2}_\d{2}_\d{6}_DA\d{5}_A')
        format_entry = ttk.Entry(format_frame, textvariable=self.format_var, width=40)
        format_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
    
    def log(self, message):
        """添加消息到日志区域"""
        try:
            self.log_text.configure(state=tk.NORMAL)
            self.log_text.insert(tk.END, f"{time.strftime('%H:%M:%S')} - {message}\n")
            self.log_text.see(tk.END)
            self.log_text.configure(state=tk.DISABLED)
        except Exception as e:
            print(f"日志错误: {e}")
    
    def clear_log(self):
        """清除日志内容"""
        try:
            self.log_text.configure(state=tk.NORMAL)
            self.log_text.delete(1.0, tk.END)
            self.log_text.configure(state=tk.DISABLED)
            self.log("日志已清除")
        except Exception as e:
            print(f"清除日志错误: {e}")
    
    def is_valid_format(self, text):
        """检查文本是否符合指定格式"""
        try:
            pattern = self.format_var.get()
            return bool(re.match(pattern, text))
        except re.error:
            self.log("错误：正则表达式格式不正确")
            return False
        except Exception as e:
            self.log(f"格式检查错误: {str(e)}")
            return False
    
    def get_active_excel_info(self):
        """获取当前活动的Excel文件信息"""
        try:
            # 获取前台窗口句柄
            hwnd = win32gui.GetForegroundWindow()
            
            # 获取进程ID
            _, process_id = win32process.GetWindowThreadProcessId(hwnd)
            
            # 获取进程名称
            process = psutil.Process(process_id)
            process_name = process.name().lower()
            
            # 如果是Excel进程
            if "excel" in process_name:
                # 获取窗口标题（通常包含文件名）
                window_title = win32gui.GetWindowText(hwnd)
                
                # Excel窗口标题通常是"文件名 - Excel"格式
                if " - Excel" in window_title:
                    return window_title.replace(" - Excel", "")
                return window_title
            return "未检测到Excel"
        except Exception as e:
            return "未检测到Excel"
    
    def update_excel_info(self):
        """定期更新Excel信息"""
        if not self.running:
            return
            
        try:
            excel_name = self.get_active_excel_info()
            if excel_name != self.excel_file_name:
                self.excel_file_name = excel_name
                self.excel_file_label.config(text=excel_name)
                if "Excel" in excel_name:
                    self.log(f"检测到Excel: {excel_name}")
        except Exception as e:
            self.log(f"更新Excel信息错误: {str(e)}")
            
        # 每秒检查一次
        self.check_excel_timer = self.root.after(1000, self.update_excel_info)
    
    def monitor_clipboard(self):
        """监控剪贴板内容"""
        try:
            self.previous_content = pyperclip.paste()
            self.log("开始监控剪贴板...")
            
            while self.running:
                try:
                    current_content = pyperclip.paste()
                    
                    # 检查剪贴板内容是否有变化
                    if current_content != self.previous_content:
                        content_preview = current_content[:30] + "..." if len(current_content) > 30 else current_content
                        self.log(f"检测到新内容: {content_preview}")
                        
                        # 检查新内容是否符合指定格式
                        if self.is_valid_format(current_content):
                            self.log("内容符合格式，执行粘贴操作...")
                            # 执行粘贴操作
                            pyautogui.hotkey('ctrl', 'v')
                            self.log("已粘贴内容到Excel")
                        else:
                            self.log("内容不符合指定格式，忽略")
                        
                        self.previous_content = current_content
                except Exception as e:
                    self.log(f"监控循环错误: {str(e)}")
                
                # 短暂暂停以减少CPU使用
                time.sleep(0.5)
        except Exception as e:
            self.log(f"监控线程错误: {str(e)}")
            traceback.print_exc()
    
    def start_monitoring(self):
        """开始监控"""
        try:
            self.running = True
            self.status_label.config(text="监控中")
            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
            
            self.log("正在启动监控...")
            
            # 启动监控线程
            self.monitor_thread = threading.Thread(target=self.monitor_clipboard)
            self.monitor_thread.daemon = True
            self.monitor_thread.start()
            
            # 启动Excel检测更新
            self.update_excel_info()
        except Exception as e:
            self.log(f"启动监控错误: {str(e)}")
            messagebox.showerror("错误", f"启动监控失败: {str(e)}")
            self.running = False
    
    def stop_monitoring(self):
        """停止监控"""
        try:
            self.running = False
            self.status_label.config(text="已停止")
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            
            if self.check_excel_timer:
                self.root.after_cancel(self.check_excel_timer)
            
            self.log("监控已停止")
        except Exception as e:
            self.log(f"停止监控错误: {str(e)}")
    
    def on_closing(self):
        """关闭窗口处理"""
        if self.running:
            if messagebox.askokcancel("退出", "监控仍在运行中，确定要退出吗？"):
                self.stop_monitoring()
                self.root.destroy()
        else:
            self.root.destroy()

def main():
    try:
        root = tk.Tk()
        app = AutoCopyApp(root)
        root.mainloop()
    except Exception as e:
        messagebox.showerror("程序错误", f"发生错误: {str(e)}")
        traceback.print_exc()

if __name__ == "__main__":
    main() 