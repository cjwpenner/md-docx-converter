from md_docx_converter.heading_mapper import md_heading_offset, docx_heading_offset

def test_single_h1_gives_offset_1():
    md = "# My Title\n\n## Section\n\nSome text."
    assert md_heading_offset(md) == 1

def test_multiple_h1_gives_offset_0():
    md = "# First\n\n# Second\n\n## Sub"
    assert md_heading_offset(md) == 0

def test_no_h1_gives_offset_0():
    md = "## Section\n\nSome text."
    assert md_heading_offset(md) == 0

def test_h1_only_no_other_headings_gives_offset_1():
    md = "# Just a title\n\nSome body text."
    assert md_heading_offset(md) == 1
