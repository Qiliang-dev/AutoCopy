@echo off
echo ============================================
echo 准备AutoCopy资源文件
echo ============================================

:: 创建资源目录
if not exist resources mkdir resources

:: 下载图标文件(如果不存在)
if not exist clipboard.ico (
    echo 下载剪贴板图标文件...
    powershell -Command "& {Invoke-WebRequest -Uri 'https://icons-for-free.com/iconfiles/png/512/clipboard+copy+document+list+office+paste+icon-1320161389075402003.png' -OutFile 'resources\clipboard_temp.png'}"
    
    :: 使用ImageMagick转换PNG到ICO (如果已安装)
    where convert >nul 2>nul
    if %ERRORLEVEL% EQU 0 (
        echo 转换PNG到ICO格式...
        convert resources\clipboard_temp.png -resize 256x256 clipboard.ico
        del resources\clipboard_temp.png
    ) else (
        echo 注意: 未找到ImageMagick, 无法自动转换图标。
        echo 请手动将PNG文件转换为ICO格式并命名为clipboard.ico
        echo 您可以使用在线工具如: https://www.icoconverter.com/
    )
)

:: 创建空的日志目录
if not exist logs mkdir logs
echo 已创建日志目录: logs

:: 创建配置文件目录
if not exist config mkdir config

:: 创建默认配置文件
echo 创建默认配置文件...
echo {> config\settings.json
echo   "auto_paste": true,>> config\settings.json
echo   "clipboard_patterns": ["^[A-Z][0-9]{5}$", "^[0-9]{8}$"],>> config\settings.json
echo   "duplicate_time": 3,>> config\settings.json
echo   "excel_format": "General",>> config\settings.json
echo   "notification_duration": 2,>> config\settings.json
echo   "suppress_warnings": false>> config\settings.json
echo }>> config\settings.json

echo ============================================
echo 资源文件准备完成!
echo ============================================
pause 