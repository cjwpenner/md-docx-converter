"""
Run once to create the desktop shortcut for MD-DOCX Converter.

Prerequisites:
    pip install pywin32

This script is a one-off utility; pywin32 is not a runtime dependency
of the converter itself.
"""
import sys
from pathlib import Path

try:
    import win32com.client
except ImportError:
    print("pywin32 not installed. Run: pip install pywin32")
    sys.exit(1)

try:
    import winreg
    key = winreg.OpenKey(
        winreg.HKEY_CURRENT_USER,
        r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    )
    desktop = Path(winreg.QueryValueEx(key, "Desktop")[0])
except Exception:
    desktop = Path.home() / "Desktop"
launch_pyw = (Path(__file__).parent / "md_docx_converter" / "launch.pyw").resolve()
python_exe = Path(sys.executable).with_name("pythonw.exe")

shell = win32com.client.Dispatch("WScript.Shell")
shortcut = shell.CreateShortCut(str(desktop / "MD-DOCX Converter.lnk"))
shortcut.TargetPath = str(python_exe)
shortcut.Arguments = f'"{launch_pyw}"'
shortcut.WorkingDirectory = str(launch_pyw.parent)
shortcut.Description = "Convert between Markdown and Word documents"
shortcut.save()

print(f"Shortcut created on Desktop: MD-DOCX Converter.lnk")
print(f"  Target: {python_exe}")
print(f"  Script: {launch_pyw}")
