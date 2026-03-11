from pathlib import Path
import pytest
from md_docx_converter.image_handler import resolve_image_path

FIXTURES = Path(__file__).parent / "fixtures"

def test_relative_path_found():
    result = resolve_image_path("sample.png", FIXTURES)
    assert result == (FIXTURES / "sample.png").resolve()

def test_relative_path_not_found():
    result = resolve_image_path("missing.png", FIXTURES)
    assert result is None

def test_absolute_path_returns_none():
    result = resolve_image_path("/absolute/path/image.png", FIXTURES)
    assert result is None

def test_url_returns_none():
    result = resolve_image_path("https://example.com/img.png", FIXTURES)
    assert result is None


import tempfile
from md_docx_converter.image_handler import extract_docx_images

def test_extract_images_creates_folder():
    with tempfile.TemporaryDirectory() as tmp:
        out_md = Path(tmp) / "output.md"
        extract_docx_images(FIXTURES / "with_image.docx", out_md)
        images_dir = Path(tmp) / "output_images"
        assert images_dir.exists()

def test_extract_images_returns_map():
    with tempfile.TemporaryDirectory() as tmp:
        out_md = Path(tmp) / "output.md"
        image_map = extract_docx_images(FIXTURES / "with_image.docx", out_md)
        assert len(image_map) >= 1
        for extracted_path in image_map.values():
            assert Path(extracted_path).exists()
