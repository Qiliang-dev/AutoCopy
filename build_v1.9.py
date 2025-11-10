"""
AutoCopy v1.9 打包脚本
使用PyInstaller打包成独立exe文件
"""

import subprocess
import sys
from pathlib import Path

def build_exe():
    """打包AutoCopy为exe文件"""
    
    print("=" * 70)
    print("  AutoCopy v1.9 打包脚本")
    print("=" * 70)
    print()
    
    # 确保在正确的目录
    script_dir = Path(__file__).parent
    print(f"工作目录: {script_dir}")
    print()
    
    # 检查必要文件
    main_file = script_dir / "autocopy_main.py"
    icon_file = script_dir / "resources" / "icons" / "autocopy.ico"
    
    if not main_file.exists():
        print(f"❌ 错误: 找不到主文件 {main_file}")
        return False
    
    print(f"✅ 主文件: {main_file}")
    
    if icon_file.exists():
        print(f"✅ 图标文件: {icon_file}")
    else:
        print(f"⚠️  警告: 找不到图标文件 {icon_file}")
        icon_file = None
    
    print()
    print("-" * 70)
    print("开始打包...")
    print("-" * 70)
    print()
    
    # PyInstaller参数
    cmd = [
        "pyinstaller",
        "--name=Autocopy_V1.9",           # 输出文件名
        "--onefile",                       # 打包成单个文件
        "--noconsole",                     # ✅ 不显示命令行窗口
        "--clean",                         # 清理临时文件
        "--noconfirm",                     # 不询问确认
    ]
    
    # 添加图标
    if icon_file:
        cmd.append(f"--icon={icon_file}")
    
    # 添加数据文件
    cmd.extend([
        f"--add-data={script_dir / 'autocopy'};autocopy",
        f"--add-data={script_dir / 'resources'};resources",
    ])
    
    # 隐藏导入（确保所有依赖都被包含）
    hidden_imports = [
        "win32com.client",
        "pythoncom",
        "pyperclip",
        "pyautogui",
        "tkinter",
        "queue",
        "json",
        "logging.handlers",
    ]
    
    for module in hidden_imports:
        cmd.append(f"--hidden-import={module}")
    
    # 主文件
    cmd.append(str(main_file))
    
    print("执行命令:")
    print(" ".join(cmd))
    print()
    
    try:
        # 运行PyInstaller
        result = subprocess.run(cmd, check=True, cwd=str(script_dir))
        
        print()
        print("=" * 70)
        print("✅ 打包完成！")
        print("=" * 70)
        print()
        
        exe_file = script_dir / "dist" / "Autocopy_V1.9.exe"
        if exe_file.exists():
            size_mb = exe_file.stat().st_size / (1024 * 1024)
            print(f"📦 输出文件: {exe_file}")
            print(f"📊 文件大小: {size_mb:.2f} MB")
            print()
            print("🎉 现在可以运行 Autocopy_V1.9.exe 了！")
            print()
            print("特性:")
            print("  ✅ 无命令行窗口")
            print("  ✅ 自动保存和加载设置")
            print("  ✅ 第一次运行时自动创建log文件")
            print("  ✅ 极致稳定（消息队列架构）")
        else:
            print(f"⚠️  警告: 找不到输出文件 {exe_file}")
        
        print()
        return True
        
    except subprocess.CalledProcessError as e:
        print()
        print("=" * 70)
        print("❌ 打包失败！")
        print("=" * 70)
        print(f"错误: {e}")
        print()
        print("可能的原因:")
        print("1. 未安装PyInstaller: pip install pyinstaller")
        print("2. 缺少依赖包: pip install pywin32 pyperclip pyautogui")
        print("3. 文件路径问题")
        print()
        return False
    
    except Exception as e:
        print()
        print("=" * 70)
        print("❌ 发生未知错误！")
        print("=" * 70)
        print(f"错误: {e}")
        import traceback
        traceback.print_exc()
        print()
        return False

if __name__ == "__main__":
    try:
        success = build_exe()
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n\n用户中断打包")
        sys.exit(1)

