"""
Desktop launcher for MD-DOCX Converter.
Opens a cmd console window and runs converter.py inside it.
converter.py handles the 'Press Enter to close' prompt itself.
"""
import subprocess
import sys
from pathlib import Path

converter = Path(__file__).parent / "converter.py"
subprocess.run(
    ["cmd", "/c", f'python "{converter}"'],
    creationflags=subprocess.CREATE_NEW_CONSOLE,
)
