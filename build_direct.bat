@echo off
echo ============================================
echo AutoCopy 简易构建工具 - 直接生成可执行文件
echo ============================================

setlocal EnableDelayedExpansion

:: 检查Python环境
where python >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo 错误: 未安装Python或Python不在PATH中
    echo 请先安装Python并确保其在PATH中
    goto :end
)

:: 创建输出目录
echo 创建独立版本的目录结构...
if not exist "dist_standalone" mkdir dist_standalone
if not exist "dist_standalone\AutoCopy" mkdir dist_standalone\AutoCopy

:: 复制主程序文件
echo 复制主程序文件...
copy auto_copy_gui.py dist_standalone\AutoCopy\

:: 复制README
echo 复制README文件...
copy README.md dist_standalone\AutoCopy\README.txt

:: 创建启动脚本
echo 创建启动脚本...
echo @echo off > dist_standalone\AutoCopy\AutoCopy.bat
echo python auto_copy_gui.py >> dist_standalone\AutoCopy\AutoCopy.bat

:: 创建配置目录
if not exist "dist_standalone\AutoCopy\config" mkdir dist_standalone\AutoCopy\config
if not exist "dist_standalone\AutoCopy\logs" mkdir dist_standalone\AutoCopy\logs

:: 创建配置文件
echo 创建配置文件...
echo {> dist_standalone\AutoCopy\config\settings.json
echo   "auto_paste": true,>> dist_standalone\AutoCopy\config\settings.json
echo   "clipboard_patterns": ["^[A-Z][0-9]{5}$", "^[0-9]{8}$"],>> dist_standalone\AutoCopy\config\settings.json
echo   "duplicate_time": 3,>> dist_standalone\AutoCopy\config\settings.json
echo   "excel_format": "General",>> dist_standalone\AutoCopy\config\settings.json
echo   "notification_duration": 2,>> dist_standalone\AutoCopy\config\settings.json
echo   "suppress_warnings": false>> dist_standalone\AutoCopy\config\settings.json
echo }>> dist_standalone\AutoCopy\config\settings.json

:: 创建依赖项安装文件
echo 创建依赖项安装脚本...
echo @echo off > dist_standalone\AutoCopy\install_dependencies.bat
echo echo 正在安装AutoCopy所需依赖... >> dist_standalone\AutoCopy\install_dependencies.bat
echo pip install pyperclip pyautogui pywin32 >> dist_standalone\AutoCopy\install_dependencies.bat
echo echo 依赖安装完成，现在可以运行AutoCopy.bat了 >> dist_standalone\AutoCopy\install_dependencies.bat
echo pause >> dist_standalone\AutoCopy\install_dependencies.bat

:: 创建使用说明
echo 创建使用说明...
echo === AutoCopy 使用说明 === > dist_standalone\AutoCopy\使用说明.txt
echo. >> dist_standalone\AutoCopy\使用说明.txt
echo 1. 首次使用：运行"install_dependencies.bat"安装必要依赖 >> dist_standalone\AutoCopy\使用说明.txt
echo 2. 启动程序：双击"AutoCopy.bat" >> dist_standalone\AutoCopy\使用说明.txt
echo 3. 连接Excel：在程序中点击"Connect to Excel"按钮 >> dist_standalone\AutoCopy\使用说明.txt
echo 4. 开始监控：点击"Start Monitoring"开始自动复制功能 >> dist_standalone\AutoCopy\使用说明.txt
echo. >> dist_standalone\AutoCopy\使用说明.txt
echo 详细说明请参阅README.txt文件 >> dist_standalone\AutoCopy\使用说明.txt

:: 创建requirements.txt
echo 创建requirements.txt...
echo pyperclip > dist_standalone\AutoCopy\requirements.txt
echo pyautogui >> dist_standalone\AutoCopy\requirements.txt
echo pywin32 >> dist_standalone\AutoCopy\requirements.txt

:: 创建ZIP压缩包
echo 创建ZIP压缩包...
if exist AutoCopy_Standalone.zip del AutoCopy_Standalone.zip
powershell -Command "Compress-Archive -Path 'dist_standalone\AutoCopy' -DestinationPath 'AutoCopy_Standalone.zip' -Force"

echo.
echo ============================================
echo 构建完成！
echo.
echo 生成的文件在: AutoCopy_Standalone.zip
echo.
echo 注意: 用户需要安装Python并运行install_dependencies.bat
echo       才能使用此版本的AutoCopy。
echo ============================================

:end
pause 