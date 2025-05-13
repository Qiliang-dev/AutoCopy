@echo off
echo ============================================
echo AutoCopy发布版本创建向导
echo ============================================

:: 检查环境
where python >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo 错误: 未安装Python或Python不在PATH中
    echo 请先安装Python并确保其在PATH中
    goto :end
)

:: 设置版本号
set /p VERSION="请输入版本号 (默认为1.0.0): " || set VERSION=1.0.0
set PACKAGE_NAME=AutoCopy_%VERSION%

echo.
echo 步骤1: 准备资源文件
echo ----------------------------------------
call build_resources.bat
if %ERRORLEVEL% NEQ 0 (
    echo 错误: 资源文件准备失败
    goto :end
)

echo.
echo 步骤2: 创建发布包
echo ----------------------------------------
call build_installer.bat
if %ERRORLEVEL% NEQ 0 (
    echo 错误: 发布包创建失败
    goto :end
)

:: 检查发布包是否已创建
if not exist %PACKAGE_NAME%.zip (
    echo 错误: 未找到发布包 %PACKAGE_NAME%.zip
    goto :end
)

:: 打开资源管理器显示发布包
echo.
echo 步骤3: 打开发布包位置
echo ----------------------------------------
echo 发布包已创建: %PACKAGE_NAME%.zip
echo 正在打开文件位置...
explorer /select,%cd%\%PACKAGE_NAME%.zip

echo.
echo ============================================
echo 发布版本创建完成!
echo 您可以在资源管理器中找到发布包: %PACKAGE_NAME%.zip
echo ============================================

:end
pause 