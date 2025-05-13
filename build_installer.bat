@echo off
echo ============================================
echo Building AutoCopy Tool Installer Package
echo ============================================

set VERSION=1.0.0
set PACKAGE_NAME=AutoCopy_%VERSION%

:: 清理之前的构建
echo Cleaning previous builds...
if exist dist rmdir /S /Q dist
if exist build rmdir /S /Q build
if exist __pycache__ rmdir /S /Q __pycache__
if exist %PACKAGE_NAME%.zip del %PACKAGE_NAME%.zip

:: 安装必要的依赖包
echo Installing required packages...
pip install -q pyinstaller pyperclip pyautogui pywin32

:: 使用PyInstaller构建可执行文件
echo Building executable with PyInstaller...
pyinstaller --clean --noconfirm --onefile --windowed --name="AutoCopy" auto_copy_gui.py

if not exist dist\AutoCopy.exe (
    echo Build failed! Could not create executable.
    goto :error
)

:: 创建分发目录结构
echo Creating distribution package...
mkdir dist\AutoCopy

:: 移动可执行文件
move dist\AutoCopy.exe dist\AutoCopy\

:: 复制README文件
copy README.md dist\AutoCopy\README.txt

:: 创建版本信息文件
echo AutoCopy Tool v%VERSION% > dist\AutoCopy\version.txt
echo Build date: %date% %time% >> dist\AutoCopy\version.txt

:: 创建快捷方式脚本
echo @echo off > dist\AutoCopy\Start_AutoCopy.bat
echo start AutoCopy.exe >> dist\AutoCopy\Start_AutoCopy.bat

:: 创建压缩包
echo Creating ZIP archive...
powershell -command "Compress-Archive -Path 'dist\AutoCopy' -DestinationPath '%PACKAGE_NAME%.zip' -Force"

:: 清理临时文件
echo Cleaning temporary files...
rmdir /S /Q build
if exist *.spec del *.spec

echo ============================================
echo Build completed successfully!
echo Release package created: %PACKAGE_NAME%.zip
echo ============================================
goto :end

:error
echo ============================================
echo Build process failed!
echo ============================================

:end
pause 