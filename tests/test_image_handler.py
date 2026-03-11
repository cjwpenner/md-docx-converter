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
