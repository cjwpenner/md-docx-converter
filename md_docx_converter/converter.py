import sys
from pathlib import Path

# Ensure the repo root is on sys.path so md_docx_converter is importable
# regardless of the working directory when launched via desktop shortcut.
_repo_root = Path(__file__).resolve().parent.parent
if str(_repo_root) not in sys.path:
    sys.path.insert(0, str(_repo_root))


def validate_extension(path: Path) -> bool:
    return path.suffix.lower() in (".md", ".docx")


def determine_output_path(input_path: Path) -> Path:
    if input_path.suffix.lower() == ".md":
        return input_path.with_suffix(".docx")
    return input_path.with_suffix(".md")


def run():
    print("MD ↔ DOCX Converter")
    print("--------------------")

    while True:
        raw = input("\nEnter file path (or press Enter to quit): ").strip().strip('"').strip("'").strip()
        if not raw:
            break

        input_path = Path(raw)

        if not input_path.exists():
            print(f"  File not found: {input_path}")
            continue

        if not validate_extension(input_path):
            print(f"  Not a .md or .docx file: {input_path.name}")
            continue

        out_path = determine_output_path(input_path)

        if out_path.exists():
            answer = input(f"  {out_path.name} already exists. Overwrite? [y/N] ").strip().lower()
            if answer != "y":
                print("  Cancelled.")
                continue

        try:
            if input_path.suffix.lower() == ".md":
                from md_docx_converter.md_to_docx import convert_md_to_docx
                convert_md_to_docx(input_path, out_path)
            else:
                from md_docx_converter.docx_to_md import convert_docx_to_md
                convert_docx_to_md(input_path, out_path)

            print(f"  ✓ Saved: {out_path}")
        except Exception as e:
            print(f"  Error: {e}")


if __name__ == "__main__":
    run()
