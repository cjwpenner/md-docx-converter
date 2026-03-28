# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A Python CLI tool that converts between `.md` (Markdown) and `.docx` (Microsoft Word) formats bidirectionally. The tool is invokable via a desktop shortcut and uses a simple CLI interface.

## Project Spec (PRD.md)

Key requirements from `PRD.md`:

- **CLI interface**: Prompts for a full file path, then converts it to the other format in the same directory
- **Output naming**: `filename.docx` → `filename.md`; `filename.md` → `filename.docx`
- **Overwrite protection**: Ask before replacing an existing output file
- **Validation**: Warn if the input file is not `.md` or `.docx`
- **Desktop shortcut**: The tool must be launchable from the Windows desktop

## Heading Hierarchy Rules (Critical)

The heading level mapping between MD and DOCX is **context-dependent** and requires a full-document pre-scan:

### MD → DOCX
- If there is **exactly one `#`** at the top of the file → map it to Word **Title** style
- If there are **multiple `#`** lines → map all `#` lines to Word **Heading 1** (no Title is used)
- `##` → Heading 2, `###` → Heading 3, etc. (shifted down one level when a Title is present)

### DOCX → MD
- If the Word doc has a **Title** style → Title becomes `#`, Heading 1 becomes `##`, etc.
- If there is **no Title** style → Heading 1 becomes `#`, Heading 2 becomes `##`, etc.

## Markdown Dialect

**GitHub Flavored Markdown (GFM)** is the target standard. See `MarkdownSyntax.md` for the full element mapping — it is the authoritative reference for what this tool supports, approximates, and drops.

## Key Conversion Rules

- **Lossy elements**: best-effort, silent drop — no warnings, no comment markers
- **Tables**: always attempt conversion; merged cells unmerged (content kept), no-header Word tables get a blank header row inserted
- **Images**: extracted to `{name}_images/` subfolder (DOCX→MD); re-embedded from relative path (MD→DOCX); missing images → placeholder text
- **Unsupported Word formatting**: underline, highlight, small caps, font size changes → bold; font colour → stripped (text kept)

## Architecture

```
md_docx_converter/
├── converter.py       # CLI entry point — input, validation, orchestration
├── md_to_docx.py      # MD → DOCX: walks markdown-it-py AST, builds python-docx document
├── docx_to_md.py      # DOCX → MD: reads python-docx document, emits GFM text
├── heading_mapper.py  # Pre-scan logic for Title/Heading 1 context rule (both directions)
├── image_handler.py   # Image extraction (DOCX→MD) and re-embedding (MD→DOCX)
└── launch.pyw         # Windowless Python launcher for the desktop shortcut
```

**Libraries**:
- `markdown-it-py` — GFM-compliant Markdown parser (produces AST for `md_to_docx.py`)
- `python-docx` — read/write `.docx` files and apply Word styles

## Running the Tool

```bash
python converter.py
```

Desktop shortcut points to `launch.pyw` (opens a console window, stays open after completion).

## Design Spec

Full design decisions, conversion mapping tables, and rationale: `docs/superpowers/specs/2026-03-11-md-docx-converter-design.md`
