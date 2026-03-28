# MD-DOCX Converter

A Python tool for bidirectional conversion between **Markdown** (`.md`) and **Microsoft Word** (`.docx`). Designed to make it easy to move content between Word documents and AI tools like Claude, ChatGPT, and GitHub Copilot.

## What it does

- Converts `.md` → `.docx` with correct heading hierarchy (Title, Heading 1–9)
- Converts `.docx` → `.md` as clean GitHub Flavored Markdown (GFM)
- Runs from a simple desktop shortcut — no command line knowledge needed
- Handles headings, bold/italic/strikethrough, lists, task lists, tables, blockquotes, code blocks, images, and hyperlinks

See [MarkdownSyntax.md](MarkdownSyntax.md) for the full element mapping and notes on what is preserved, approximated, or dropped.

## Requirements

- Windows 10/11
- Python 3.11+
- The following Python packages (installed via pip):

```
pip install markdown-it-py python-docx
```

## Setup

### 1. Clone the repository

```bash
git clone https://github.com/cjwpenner/md-docx-converter.git
cd md-docx-converter
```

### 2. Install dependencies

```bash
pip install markdown-it-py python-docx
```

### 3. Create the desktop shortcut

```bash
pip install pywin32
python create_shortcut.py
```

This creates an **MD-DOCX Converter** shortcut on your Windows desktop. `pywin32` is only needed to create the shortcut — it is not required to run the converter itself.

### 4. Run the converter

Double-click **MD-DOCX Converter** on your desktop. A console window opens and prompts:

```
MD ↔ DOCX Converter
--------------------
Enter file path:
```

Paste or type the full path to your `.md` or `.docx` file and press Enter. The converted file is saved in the same directory with the extension swapped.

You can also run directly from the command line:

```bash
python md_docx_converter/converter.py
```

## Conversion notes

### Heading hierarchy

The heading level mapping is context-dependent:

- **MD → DOCX**: If there is exactly one `#` in the document, it becomes a Word **Title**. All other headings shift down by one level. If there are multiple `#` headings, they all become **Heading 1** with no Title.
- **DOCX → MD**: If the document has a **Title** style, it becomes `#`. All headings shift up accordingly. If there is no Title, **Heading 1** becomes `#`.

### Lossy elements

Word formatting that has no Markdown equivalent is approximated as **bold**:

| Word formatting | Markdown output |
|---|---|
| Underline | `**bold**` |
| Highlight | `**bold**` |
| Small caps | `**bold**` |
| Font colour | Stripped (text kept) |

### Images

- **DOCX → MD**: Embedded images are extracted to a `{filename}_images/` folder next to the output `.md` file.
- **MD → DOCX**: Images referenced by relative path are re-embedded. Missing images become `[image not found: path]`.

## MCP server (Claude / AI integration)

This tool is also available as an **MCP server**, letting Claude and other AI assistants read and write Word documents directly.

### Install

```bash
pip install mcp-md-docx
```

### Configure Claude Desktop

Add to `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "md-docx": {
      "command": "python",
      "args": ["-m", "mcp_md_docx"]
    }
  }
}
```

### Configure Claude Code

```bash
claude mcp add md-docx -- python -m mcp_md_docx
```

### Tools exposed

| Tool | What it does |
|---|---|
| `read_docx` | Read a `.docx` file — returns full Markdown text to the AI |
| `write_docx` | Create a `.docx` from Markdown text the AI has written |
| `convert_md_file_to_docx` | Convert a `.md` file on disk to `.docx` |
| `convert_docx_file_to_md` | Convert a `.docx` file on disk to `.md` |

Once configured, you can say things like:
- *"Read `report.docx` and summarise it"*
- *"Turn this into a Word document and save it to my Desktop"*
- *"Convert all the bullet points in `notes.docx` into a table"*

## Project structure

```
md_docx_converter/
├── converter.py       # CLI entry point
├── md_to_docx.py      # Markdown → Word conversion
├── docx_to_md.py      # Word → Markdown conversion
├── heading_mapper.py  # Heading hierarchy pre-scan logic
├── image_handler.py   # Image extraction and embedding
└── launch.pyw         # Desktop shortcut launcher
mcp_md_docx/
├── server.py          # MCP server (four tools)
└── __main__.py        # Entry point for python -m mcp_md_docx
create_shortcut.py     # One-time shortcut setup script
pyproject.toml         # PyPI packaging config
```

## License

This project is licensed under the **GNU General Public License v3.0** (GPLv3). You are free to use, modify, and distribute this software, provided that any derivative works are also distributed under the same licence.

See [LICENSE](LICENSE) for the full licence text.
