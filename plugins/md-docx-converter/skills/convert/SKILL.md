---
name: convert
description: Convert between Markdown (.md) and Word (.docx) files. Use when the user asks to convert a file to Word, save something as a Word document, convert a Word doc to Markdown, or read a .docx file.
argument-hint: <file-path>
---

Convert the file at `$ARGUMENTS` between Markdown and Word format using the md-docx-converter MCP tools.

## Steps

1. If no argument was provided, ask the user for the file path.

2. Check the file extension:
   - `.md` file → call `convert_md_file_to_docx` with the full absolute path
   - `.docx` file → call `convert_docx_file_to_md` with the full absolute path
   - Anything else → tell the user only `.md` and `.docx` files are supported

3. Report the result: on success, tell the user where the output file was saved.
   On error, show the error message clearly.

## Additional tools available

- `read_docx(path)` — read a Word document and return its content as Markdown (useful when you want to work with the content without saving a .md file)
- `write_docx(markdown, output_path)` — convert a Markdown string directly to a .docx file (useful when generating Word docs from scratch)
