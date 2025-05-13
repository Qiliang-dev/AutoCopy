@echo off 
echo 正在启动AutoCopy工具... 
 
:: 检查Python是否安装 
where python 
if %ERRORLEVEL% NEQ 0 ( 
    echo 错误: 未安装Python! 
    echo 请先下载并安装Python 3.8或更高版本 
    echo 下载地址: https://www.python.org/downloads/ 
    pause 
    exit /b 
) 
 
:: 检查所需模块是否安装 
echo 检查所需模块... 
 
python -c "import pyperclip" 
if %ERRORLEVEL% NEQ 0 ( 
    echo 正在安装pyperclip模块... 
    pip install pyperclip 
) 
 
python -c "import pyautogui" 
if %ERRORLEVEL% NEQ 0 ( 
    echo 正在安装pyautogui模块... 
    pip install pyautogui 
) 
 
python -c "import win32com" 
if %ERRORLEVEL% NEQ 0 ( 
    echo 正在安装pywin32模块... 
    pip install pywin32 
) 
 
:: 启动程序 
echo 启动AutoCopy... 
python auto_copy_gui.py 
