import re
from pathlib import Path
from docx import Document as DocxDocument


def md_heading_offset(md_text: str) -> int:
    """Return 1 if there is exactly one # heading (use Title style), else 0."""
    h1_lines = re.findall(r'^#(?!#)', md_text, re.MULTILINE)
    return 1 if len(h1_lines) == 1 else 0


def docx_heading_offset(docx_path: Path) -> int:
    """Return 1 if the DOCX contains a Title-style paragraph, else 0."""
    doc = DocxDocument(str(docx_path))
    for para in doc.paragraphs:
        if para.style.name == "Title":
            return 1
    return 0
