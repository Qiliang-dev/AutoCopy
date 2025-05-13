@echo off
echo ============================================
echo AutoCopy 超简易打包工具 - 纯绿色版
echo ============================================

:: 创建目录
if not exist "dist_green" mkdir dist_green
if not exist "dist_green\AutoCopy" mkdir dist_green\AutoCopy

:: 复制主程序文件和配置
echo 复制程序文件...
copy auto_copy_gui.py dist_green\AutoCopy\

:: 创建配置目录
if not exist "dist_green\AutoCopy\config" mkdir dist_green\AutoCopy\config
if not exist "dist_green\AutoCopy\logs" mkdir dist_green\AutoCopy\logs

:: 创建配置文件
echo 创建配置文件...
echo {> dist_green\AutoCopy\config\settings.json
echo   "auto_paste": true,>> dist_green\AutoCopy\config\settings.json
echo   "clipboard_patterns": ["^[A-Z][0-9]{5}$", "^[0-9]{8}$"],>> dist_green\AutoCopy\config\settings.json
echo   "duplicate_time": 3,>> dist_green\AutoCopy\config\settings.json
echo   "excel_format": "General",>> dist_green\AutoCopy\config\settings.json
echo   "notification_duration": 2,>> dist_green\AutoCopy\config\settings.json
echo   "suppress_warnings": false>> dist_green\AutoCopy\config\settings.json
echo }>> dist_green\AutoCopy\config\settings.json

:: 创建一键启动脚本
echo 创建一键启动脚本...
echo @echo off > dist_green\AutoCopy\启动AutoCopy.bat
echo echo 正在启动AutoCopy工具... >> dist_green\AutoCopy\启动AutoCopy.bat
echo. >> dist_green\AutoCopy\启动AutoCopy.bat
echo :: 检查Python是否安装 >> dist_green\AutoCopy\启动AutoCopy.bat
echo where python >nul 2>nul >> dist_green\AutoCopy\启动AutoCopy.bat
echo if %%ERRORLEVEL%% NEQ 0 ( >> dist_green\AutoCopy\启动AutoCopy.bat
echo     echo 错误: 未安装Python! >> dist_green\AutoCopy\启动AutoCopy.bat
echo     echo 请先下载并安装Python 3.8或更高版本 >> dist_green\AutoCopy\启动AutoCopy.bat
echo     echo 下载地址: https://www.python.org/downloads/ >> dist_green\AutoCopy\启动AutoCopy.bat
echo     pause >> dist_green\AutoCopy\启动AutoCopy.bat
echo     exit /b >> dist_green\AutoCopy\启动AutoCopy.bat
echo ) >> dist_green\AutoCopy\启动AutoCopy.bat
echo. >> dist_green\AutoCopy\启动AutoCopy.bat
echo :: 检查所需模块是否安装 >> dist_green\AutoCopy\启动AutoCopy.bat
echo echo 检查所需模块... >> dist_green\AutoCopy\启动AutoCopy.bat
echo. >> dist_green\AutoCopy\启动AutoCopy.bat
echo python -c "import pyperclip" >nul 2>nul >> dist_green\AutoCopy\启动AutoCopy.bat
echo if %%ERRORLEVEL%% NEQ 0 ( >> dist_green\AutoCopy\启动AutoCopy.bat
echo     echo 正在安装pyperclip模块... >> dist_green\AutoCopy\启动AutoCopy.bat
echo     pip install pyperclip >> dist_green\AutoCopy\启动AutoCopy.bat
echo ) >> dist_green\AutoCopy\启动AutoCopy.bat
echo. >> dist_green\AutoCopy\启动AutoCopy.bat
echo python -c "import pyautogui" >nul 2>nul >> dist_green\AutoCopy\启动AutoCopy.bat
echo if %%ERRORLEVEL%% NEQ 0 ( >> dist_green\AutoCopy\启动AutoCopy.bat
echo     echo 正在安装pyautogui模块... >> dist_green\AutoCopy\启动AutoCopy.bat
echo     pip install pyautogui >> dist_green\AutoCopy\启动AutoCopy.bat
echo ) >> dist_green\AutoCopy\启动AutoCopy.bat
echo. >> dist_green\AutoCopy\启动AutoCopy.bat
echo python -c "import win32com" >nul 2>nul >> dist_green\AutoCopy\启动AutoCopy.bat
echo if %%ERRORLEVEL%% NEQ 0 ( >> dist_green\AutoCopy\启动AutoCopy.bat
echo     echo 正在安装pywin32模块... >> dist_green\AutoCopy\启动AutoCopy.bat
echo     pip install pywin32 >> dist_green\AutoCopy\启动AutoCopy.bat
echo ) >> dist_green\AutoCopy\启动AutoCopy.bat
echo. >> dist_green\AutoCopy\启动AutoCopy.bat
echo :: 启动程序 >> dist_green\AutoCopy\启动AutoCopy.bat
echo echo 启动AutoCopy... >> dist_green\AutoCopy\启动AutoCopy.bat
echo python auto_copy_gui.py >> dist_green\AutoCopy\启动AutoCopy.bat

:: 创建使用说明
echo 创建使用说明...
echo ===== AutoCopy 使用说明 ===== > dist_green\AutoCopy\使用说明.txt
echo. >> dist_green\AutoCopy\使用说明.txt
echo 一、安装说明: >> dist_green\AutoCopy\使用说明.txt
echo   1. 需要先安装Python 3.8或更高版本 >> dist_green\AutoCopy\使用说明.txt
echo      下载地址: https://www.python.org/downloads/ >> dist_green\AutoCopy\使用说明.txt
echo   2. 双击"启动AutoCopy.bat"，程序会自动检查并安装所需依赖 >> dist_green\AutoCopy\使用说明.txt
echo. >> dist_green\AutoCopy\使用说明.txt
echo 二、使用说明: >> dist_green\AutoCopy\使用说明.txt
echo   1. 启动程序后，点击"Connect to Excel"连接Excel >> dist_green\AutoCopy\使用说明.txt
echo   2. 点击"Start Monitoring"开始监控剪贴板 >> dist_green\AutoCopy\使用说明.txt
echo   3. 当复制符合格式的内容时，将自动粘贴到Excel中 >> dist_green\AutoCopy\使用说明.txt
echo   4. 点击"Stop Monitoring"停止监控 >> dist_green\AutoCopy\使用说明.txt
echo. >> dist_green\AutoCopy\使用说明.txt
echo 三、常见问题: >> dist_green\AutoCopy\使用说明.txt
echo   1. 如果无法连接Excel，请确保Excel已打开 >> dist_green\AutoCopy\使用说明.txt
echo   2. 如果自动粘贴不工作，请尝试手动点击"Paste Now"按钮 >> dist_green\AutoCopy\使用说明.txt
echo   3. 日志文件保存在logs目录中，可帮助排查问题 >> dist_green\AutoCopy\使用说明.txt

:: 创建ZIP压缩包
echo 创建ZIP压缩包...
if exist AutoCopy_Green.zip del AutoCopy_Green.zip
powershell -Command "Compress-Archive -Path 'dist_green\AutoCopy' -DestinationPath 'AutoCopy_Green.zip' -Force"

echo.
echo ============================================
echo 绿色版打包完成！
echo.
echo 文件位置: AutoCopy_Green.zip
echo.
echo 用户只需解压并运行"启动AutoCopy.bat"即可使用
echo ============================================

pause 