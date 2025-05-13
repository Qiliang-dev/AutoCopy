@echo off
echo Building AutoCopy release version...

:: 安装必要的依赖包（如果尚未安装）
echo Installing required packages...
pip install pyinstaller pyperclip pyautogui pywin32

:: 使用PyInstaller构建可执行文件
echo Building executable...
pyinstaller --clean --noconfirm --onefile --windowed --icon=clipboard.ico --name="AutoCopy" auto_copy_gui.py

:: 如果build成功，复制其他必要文件到dist文件夹
if exist dist\AutoCopy.exe (
    echo Build successful!
    
    :: 创建README文件
    echo Creating README file...
    echo AutoCopy Tool > dist\README.txt
    echo ============= >> dist\README.txt
    echo. >> dist\README.txt
    echo This tool automatically copies content from clipboard to Excel when it matches the specified pattern. >> dist\README.txt
    echo. >> dist\README.txt
    echo Instructions: >> dist\README.txt
    echo 1. Open Excel and select the target cell >> dist\README.txt
    echo 2. Click "Connect to Excel" in the application >> dist\README.txt
    echo 3. Click "Start Monitoring" >> dist\README.txt
    echo 4. When matching content is detected, it will be automatically pasted to Excel >> dist\README.txt
    
    :: 创建快捷方式
    echo @echo off > dist\AutoCopy.bat
    echo start AutoCopy.exe >> dist\AutoCopy.bat
    
    echo Release package created in the "dist" folder.
    echo You can distribute the entire "dist" folder.
) else (
    echo Build failed!
)

pause 