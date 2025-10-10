# AutoCopy 工具

AutoCopy是一个自动监控剪贴板并支持快速粘贴到Excel的工具。它可以识别特定模式的文本，并在匹配时自动粘贴到Excel中的活动单元格。

## 主要功能

- 自动监控剪贴板内容变化
- 根据配置的正则表达式模式匹配特定文本格式
- 一键快速粘贴到Excel活动单元格
- 可自定义设置和配置选项
- 简洁的用户界面，易于使用

## 安装步骤

1. 下载最新的AutoCopy发布包 (`AutoCopy_x.x.x.zip`)
2. 解压缩到您想要的位置
3. 运行`AutoCopy.exe`或使用提供的快捷方式启动应用程序

## 使用说明

1. 启动AutoCopy应用程序
2. 点击"连接Excel"按钮连接到已打开的Excel实例
3. 复制符合预设模式的文本内容到剪贴板
4. 根据设置，内容将自动粘贴到Excel中的活动单元格，或者您可以点击"立即粘贴"按钮手动粘贴

## 配置选项

AutoCopy的配置选项保存在`config/settings.json`文件中，您可以根据需要修改以下设置：

- `auto_paste`: 启用/禁用自动粘贴功能
- `clipboard_patterns`: 要匹配的文本模式列表（正则表达式）
- `duplicate_time`: 重复内容的最小时间间隔（秒）
- `excel_format`: Excel单元格格式
- `notification_duration`: 通知显示持续时间（秒）
- `suppress_warnings`: 是否抑制警告消息

## 注意事项

- 使用前确保Excel已打开
- 确保Excel中的活动单元格是您想要粘贴内容的位置
- 应用程序需要使用Windows系统

## 问题排查

如果遇到问题：

1. 查看`logs`文件夹中的日志文件
2. 确保Excel已正确打开并可访问
3. 检查您的剪贴板内容是否符合配置的模式
4. 重新启动应用程序并尝试重新连接Excel

## 开发环境设置

### 安装依赖

```bash
pip install -r requirements.txt
```

### 运行开发版本

```bash
python auto_copy_gui.py
```

### 打包为可执行文件

```bash
pyinstaller --onefile --noconsole --name="AutoCopy" --icon="resources/icons/autocopy.ico" auto_copy_gui.py
```

打包后的可执行文件位于 `dist/AutoCopy.exe`

**注意**：打包时会生成 `build/`、`dist/` 文件夹和 `.spec` 文件，这些是临时文件，可以在打包完成后删除。

## 许可证

本软件遵循MIT许可证。