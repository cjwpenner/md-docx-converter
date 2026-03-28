import tempfile
from pathlib import Path

from mcp.server.fastmcp import FastMCP

from md_docx_converter.md_to_docx import convert_md_to_docx
from md_docx_converter.docx_to_md import convert_docx_to_md

mcp = FastMCP("md-docx-converter")


@mcp.tool()
def read_docx(path: str) -> str:
    """Read a Word (.docx) document and return its full content as Markdown text.
    Use this when the user asks you to read, summarise, edit, or work with a Word document.
    """
    docx_path = Path(path)
    if not docx_path.exists():
        return f"Error: file not found: {docx_path}"
    if docx_path.suffix.lower() != ".docx":
        return f"Error: expected a .docx file, got: {docx_path.suffix}"
    try:
        with tempfile.NamedTemporaryFile(suffix=".md", delete=False) as tmp:
            tmp_path = Path(tmp.name)
        convert_docx_to_md(docx_path, tmp_path)
        md_text = tmp_path.read_text(encoding="utf-8")
        tmp_path.unlink(missing_ok=True)
        # Also clean up any extracted images folder
        images_dir = tmp_path.parent / (tmp_path.stem + "_images")
        if images_dir.exists():
            import shutil
            shutil.rmtree(images_dir, ignore_errors=True)
        return md_text
    except Exception as e:
        return f"Error reading {docx_path.name}: {e}"


@mcp.tool()
def write_docx(markdown: str, output_path: str) -> str:
    """Convert Markdown text to a Word (.docx) document and save it to disk.
    Use this when the user asks you to create or save a Word document from text or Markdown content.
    The output_path should be an absolute path ending in .docx.
    """
    out_path = Path(output_path)
    if out_path.suffix.lower() != ".docx":
        out_path = out_path.with_suffix(".docx")
    try:
        with tempfile.NamedTemporaryFile(
            suffix=".md", mode="w", encoding="utf-8", delete=False
        ) as tmp:
            tmp.write(markdown)
            tmp_path = Path(tmp.name)
        convert_md_to_docx(tmp_path, out_path)
        tmp_path.unlink(missing_ok=True)
        return f"Saved: {out_path}"
    except Exception as e:
        return f"Error writing {out_path.name}: {e}"


@mcp.tool()
def convert_md_file_to_docx(path: str) -> str:
    """Convert a Markdown (.md) file to a Word (.docx) file.
    The output is saved alongside the input file with the same name and .docx extension.
    """
    md_path = Path(path)
    if not md_path.exists():
        return f"Error: file not found: {md_path}"
    if md_path.suffix.lower() != ".md":
        return f"Error: expected a .md file, got: {md_path.suffix}"
    try:
        out_path = md_path.with_suffix(".docx")
        convert_md_to_docx(md_path, out_path)
        return f"Saved: {out_path}"
    except Exception as e:
        return f"Error converting {md_path.name}: {e}"


@mcp.tool()
def convert_docx_file_to_md(path: str) -> str:
    """Convert a Word (.docx) file to a Markdown (.md) file.
    The output is saved alongside the input file. Returns the Markdown content.
    """
    docx_path = Path(path)
    if not docx_path.exists():
        return f"Error: file not found: {docx_path}"
    if docx_path.suffix.lower() != ".docx":
        return f"Error: expected a .docx file, got: {docx_path.suffix}"
    try:
        out_path = docx_path.with_suffix(".md")
        convert_docx_to_md(docx_path, out_path)
        md_text = out_path.read_text(encoding="utf-8")
        return f"Saved: {out_path}\n\n{md_text}"
    except Exception as e:
        return f"Error converting {docx_path.name}: {e}"


def run():
    mcp.run()
