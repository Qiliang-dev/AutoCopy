import pyperclip
import pyautogui
import re
import time

def is_valid_format(text):
    """检查文本是否符合指定格式：类似于2025_02_25_144352_DA00097_A开头"""
    pattern = r'^20\d{2}_\d{2}_\d{2}_\d{6}_DA\d{5}_A'
    return bool(re.match(pattern, text))

def monitor_clipboard():
    """监控剪贴板内容并根据条件执行粘贴操作"""
    print("Starting clipboard monitoring...")
    previous_content = pyperclip.paste()
    
    try:
        while True:
            current_content = pyperclip.paste()
            
            # 检查剪贴板内容是否有变化
            if current_content != previous_content:
                print(f"New clipboard content detected: {current_content[:50]}...")
                
                # 检查新内容是否符合指定格式
                if is_valid_format(current_content):
                    print("Content matches pattern, executing paste operation...")
                    # 给用户一点时间切换到Excel并选中目标单元格
                    time.sleep(0.5)
                    # 执行粘贴操作
                    pyautogui.hotkey('ctrl', 'v')
                    print("Content pasted to selected cell")
                else:
                    print("Content doesn't match pattern, ignoring")
                
                previous_content = current_content
            
            # 短暂暂停以减少CPU使用
            time.sleep(0.5)
    
    except KeyboardInterrupt:
        print("\nMonitoring stopped")

if __name__ == "__main__":
    print("AutoCopy Tool Started")
    print("Function: Automatically paste content that matches a pattern (like 2025_02_25_144352_DA00097_A)")
    print("Press Ctrl+C to stop")
    print("-" * 50)
    
    # 给用户时间切换到Excel
    print("Please switch to Excel and select target cell within 5 seconds...")
    for i in range(5, 0, -1):
        print(f"{i}...")
        time.sleep(1)
    
    monitor_clipboard() 