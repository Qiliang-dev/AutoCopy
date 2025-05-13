import sys
from cx_Freeze import setup, Executable

# 依赖包
build_exe_options = {
    "packages": ["tkinter", "pyperclip", "pyautogui", "win32com", "pythoncom"],
    "excludes": [],
    "include_files": [],
    "include_msvcr": True,
}

# 创建可执行文件
base = None
if sys.platform == "win32":
    base = "Win32GUI"  # 使用Windows GUI，不显示控制台窗口

setup(
    name="AutoCopy",
    version="1.0.0",
    description="AutoCopy Tool for Excel",
    options={"build_exe": build_exe_options},
    executables=[
        Executable(
            "auto_copy_gui.py", 
            base=base,
            target_name="AutoCopy.exe"
        )
    ],
) 