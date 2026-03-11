"""
Desktop launcher for MD-DOCX Converter.
Opens a cmd console window and runs converter.py inside it.
converter.py handles the 'Press Enter to close' prompt itself.
"""
import subprocess
from pathlib import Path

converter = Path(__file__).parent / "converter.py"
python = Path(r"C:\Users\Chris\AppData\Local\Programs\Python\Python311\python.exe")

subprocess.run(
    ["cmd", "/k", str(python), str(converter)],
    creationflags=subprocess.CREATE_NEW_CONSOLE,
)
