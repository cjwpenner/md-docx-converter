import zipfile
from pathlib import Path
from docx.shared import Inches


def resolve_image_path(src: str, md_dir: Path) -> Path | None:
    """
    Resolve a Markdown image src to an absolute Path, or None if unresolvable.
    URLs and absolute paths are never fetched — return None.
    """
    if src.startswith(("http://", "https://", "ftp://")):
        return None
    candidate = Path(src)
    if candidate.is_absolute():
        return None
    resolved = (md_dir / candidate).resolve()
    return resolved if resolved.exists() else None


def embed_image(para, image_path: Path) -> None:
    """Add an inline image to an existing paragraph as a picture run."""
    run = para.add_run()
    run.add_picture(str(image_path), width=Inches(4))


def extract_docx_images(docx_path: Path, out_md_path: Path) -> dict[str, Path]:
    """
    Extract all images from a DOCX file into {stem}_images/ next to out_md_path.
    Returns a dict mapping ZIP internal name (e.g. 'word/media/image1.png')
    to the extracted file Path.
    """
    stem = out_md_path.stem
    images_dir = out_md_path.parent / f"{stem}_images"
    images_dir.mkdir(exist_ok=True)

    image_map: dict[str, Path] = {}
    counter = 1

    with zipfile.ZipFile(str(docx_path), "r") as zf:
        for name in zf.namelist():
            if name.startswith("word/media/"):
                suffix = Path(name).suffix or ".png"
                out_name = f"image{counter}{suffix}"
                out_path = images_dir / out_name
                out_path.write_bytes(zf.read(name))
                image_map[name] = out_path
                counter += 1

    return image_map
