from pathlib import Path
import tempfile
from md_docx_converter.docx_to_md import convert_docx_to_md

FIXTURES = Path(__file__).parent / "fixtures"


def _convert(docx_file):
    with tempfile.TemporaryDirectory() as tmp:
        out_md = Path(tmp) / (docx_file.stem + ".md")
        convert_docx_to_md(docx_file, out_md)
        return out_md.read_text(encoding="utf-8")


def test_title_becomes_h1():
    md = _convert(FIXTURES / "rich.docx")
    assert md.startswith("# Document Title")


def test_heading1_becomes_h2_when_title_present():
    md = _convert(FIXTURES / "rich.docx")
    assert "## Section One" in md


def test_bold_run_wrapped():
    md = _convert(FIXTURES / "rich.docx")
    assert "**bold word**" in md


def test_blockquote():
    md = _convert(FIXTURES / "rich.docx")
    assert "> Quote text" in md


def test_table_pipe_syntax():
    md = _convert(FIXTURES / "rich.docx")
    assert "| H1 |" in md
    assert "| R1 |" in md


def test_no_title_h1_stays_h1():
    md = _convert(FIXTURES / "no_title.docx")
    assert "# Section One" in md


def test_code_style_becomes_fenced_block():
    """A paragraph with Code style becomes a fenced code block."""
    from docx import Document
    import pytest
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        doc = Document()
        try:
            doc.add_paragraph("print('hello')", style="Code")
        except KeyError:
            import pytest
            pytest.skip("No 'Code' style in default template")
        docx_path = tmp_path / "code.docx"
        doc.save(str(docx_path))
        out_md = tmp_path / "code.md"
        convert_docx_to_md(docx_path, out_md)
        result = out_md.read_text(encoding="utf-8")
        assert "```" in result
        assert "print('hello')" in result


def test_image_extracted_and_referenced():
    """Converting a DOCX with an embedded image produces a ![...] reference."""
    md = _convert(FIXTURES / "with_image.docx")
    assert "![" in md
