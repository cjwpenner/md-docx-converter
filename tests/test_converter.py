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
