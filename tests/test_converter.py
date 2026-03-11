from pathlib import Path
from md_docx_converter.converter import determine_output_path, validate_extension


def test_md_input_gives_docx_output():
    p = Path("/some/dir/report.md")
    assert determine_output_path(p) == Path("/some/dir/report.docx")


def test_docx_input_gives_md_output():
    p = Path("/some/dir/report.docx")
    assert determine_output_path(p) == Path("/some/dir/report.md")


def test_valid_md_extension():
    assert validate_extension(Path("file.md")) is True


def test_valid_docx_extension():
    assert validate_extension(Path("file.docx")) is True


def test_invalid_extension_txt():
    assert validate_extension(Path("file.txt")) is False


def test_invalid_extension_pdf():
    assert validate_extension(Path("file.pdf")) is False


import tempfile
import pytest
from md_docx_converter.md_to_docx import convert_md_to_docx
from md_docx_converter.docx_to_md import convert_docx_to_md

FIXTURES = Path(__file__).parent / "fixtures"
TEMPLATE = Path(r"C:\Users\Chris\AppData\Roaming\Microsoft\Templates\Normal.dotm")


def test_round_trip_headings_survive():
    if not TEMPLATE.exists():
        pytest.skip("Word template not found")
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        src_md = FIXTURES / "simple.md"
        mid_docx = tmp / "simple.docx"
        out_md = tmp / "simple_rt.md"

        convert_md_to_docx(src_md, mid_docx, TEMPLATE)
        convert_docx_to_md(mid_docx, out_md)

        result = out_md.read_text(encoding="utf-8")
        assert "# My Title" in result
        assert "## Section One" in result


def test_round_trip_table_survives():
    if not TEMPLATE.exists():
        pytest.skip("Word template not found")
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        src_md = FIXTURES / "simple.md"
        mid_docx = tmp / "simple.docx"
        out_md = tmp / "simple_rt.md"

        convert_md_to_docx(src_md, mid_docx, TEMPLATE)
        convert_docx_to_md(mid_docx, out_md)

        result = out_md.read_text(encoding="utf-8")
        assert "Col A" in result
        assert "One" in result


def test_round_trip_bold_survives():
    if not TEMPLATE.exists():
        pytest.skip("Word template not found")
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        src_md = FIXTURES / "simple.md"
        mid_docx = tmp / "simple.docx"
        out_md = tmp / "simple_rt.md"

        convert_md_to_docx(src_md, mid_docx, TEMPLATE)
        convert_docx_to_md(mid_docx, out_md)

        result = out_md.read_text(encoding="utf-8")
        assert "**bold**" in result
