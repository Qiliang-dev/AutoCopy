from pathlib import Path
import sys
import ctypes
import tkinter as tk


def resolve_resource_path(relative_path: str) -> Path:
    """Return absolute path to resource, compatible with PyInstaller."""
    try:
        base_path = Path(getattr(sys, "_MEIPASS"))
    except AttributeError:
        base_path = Path(__file__).resolve().parent.parent.parent
    return (base_path / relative_path).resolve()


def set_app_window_icon(root: tk.Tk, relative_icon_path: str = "resources/icons/autocopy.ico") -> None:
    """Apply the application icon to the main window and taskbar."""
    try:
        icon_path = resolve_resource_path(relative_icon_path)
        if icon_path.exists():
            root.iconbitmap(default=str(icon_path))
            try:
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("AutoCopy.Autocopy_V1.6")
            except Exception:
                pass
    except Exception:
        pass
