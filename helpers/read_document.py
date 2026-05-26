#!/usr/bin/env python3
"""
read_document.py — Document reading helper for the docx-cowork agent skill.

Converts DOC, DOCX, Pages, ODT, RTF, HTML, and other formats to Markdown,
and optionally extracts embedded images as base64 for vision-capable models.

Conversion strategy per format:
  .docx           → pandoc directly (native support)
  .odt .rtf .html → pandoc directly (native support)
  .doc .pages     → LibreOffice --headless → .docx, then pandoc
  anything else   → pandoc first, LibreOffice fallback, error if both fail

Usage (CLI):
    python helpers/read_document.py path/to/file.docx
    python helpers/read_document.py path/to/file.pages --no-images
    python helpers/read_document.py path/to/file.doc --output-dir ./out

Usage (Python API):
    from helpers.read_document import read_document, cleanup

    result = read_document("report.docx", extract_images=True)
    print(result["markdown"])
    for img in result["images"]:          # empty list if no images
        # img["mime_type"], img["data"] (base64 str), img["filename"]
        pass
    cleanup(result["tmp_dir"])            # always call when done
"""

from __future__ import annotations

import argparse
import base64
import mimetypes
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

# ---------------------------------------------------------------------------
# Format classification
# ---------------------------------------------------------------------------

# pandoc can read these directly (pandoc 3.x input formats relevant here)
PANDOC_NATIVE: frozenset[str] = frozenset({
    ".docx", ".odt", ".rtf",
    ".html", ".htm",
    ".epub",
    ".rst",
    ".tex", ".latex",
    ".md", ".markdown",
    ".txt",
})

# pandoc cannot read these; LibreOffice is required as the first step
LIBREOFFICE_REQUIRED: frozenset[str] = frozenset({
    ".doc",    # legacy Word binary — pandoc 3.x dropped support
    ".pages",  # Apple Pages — pandoc has never supported this
    ".numbers",
    ".key",
    ".ppt",
    ".pps",
})

# Image extensions we care about passing to the model
IMAGE_EXTENSIONS: frozenset[str] = frozenset({
    ".png", ".jpg", ".jpeg", ".gif",
    ".webp", ".bmp", ".tiff", ".tif", ".svg",
})


# ---------------------------------------------------------------------------
# System tool discovery
# ---------------------------------------------------------------------------

def _require_pandoc() -> str:
    """Return pandoc path or raise with install hint."""
    path = shutil.which("pandoc")
    if not path:
        raise EnvironmentError(
            "pandoc is not installed or not on PATH.\n"
            "  macOS :  brew install pandoc\n"
            "  Ubuntu:  sudo apt install pandoc\n"
            "  Windows: choco install pandoc  (or https://pandoc.org/installing.html)"
        )
    return path


def _find_libreoffice() -> str | None:
    """Return LibreOffice/soffice path, or None if not available."""
    candidates = [
        "libreoffice",
        "soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/libreoffice",
        "/usr/bin/soffice",
        "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
    ]
    for cmd in candidates:
        found = shutil.which(cmd) or (os.path.isfile(cmd) and cmd)
        if found:
            return found
    return None


# ---------------------------------------------------------------------------
# Conversion helpers
# ---------------------------------------------------------------------------

def _pandoc_to_docx(input_path: str, output_path: str) -> list[str]:
    """
    Convert a pandoc-supported format to DOCX.
    Returns a list of warning strings (empty on clean success).
    Raises RuntimeError on failure.
    """
    pandoc = _require_pandoc()
    result = subprocess.run(
        [pandoc, input_path, "-o", output_path],
        capture_output=True, text=True, timeout=120,
    )
    warnings: list[str] = []
    if result.stderr.strip():
        warnings.append(result.stderr.strip())
    if result.returncode != 0:
        raise RuntimeError(
            f"pandoc conversion failed (exit {result.returncode}):\n{result.stderr.strip()}"
        )
    return warnings


def _libreoffice_to_docx(input_path: str, out_dir: str) -> str:
    """
    Use LibreOffice to convert input_path → DOCX inside out_dir.
    Returns the path of the produced .docx file.
    Raises EnvironmentError if LibreOffice is not found, RuntimeError on failure.
    """
    lo = _find_libreoffice()
    if not lo:
        raise EnvironmentError(
            "LibreOffice is not installed. It is required to convert "
            f"'{Path(input_path).suffix}' files.\n"
            "  macOS :  brew install --cask libreoffice\n"
            "  Ubuntu:  sudo apt install libreoffice\n"
            "  Windows: https://www.libreoffice.org/download/"
        )
    result = subprocess.run(
        [lo, "--headless", "--convert-to", "docx", "--outdir", out_dir, input_path],
        capture_output=True, text=True, timeout=180,
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"LibreOffice conversion failed (exit {result.returncode}):\n{result.stderr.strip()}"
        )
    # LibreOffice writes <stem>.docx in out_dir
    stem = Path(input_path).stem
    produced = os.path.join(out_dir, stem + ".docx")
    if not os.path.isfile(produced):
        # Sometimes LibreOffice writes with a slightly different stem; find it
        candidates = list(Path(out_dir).glob("*.docx"))
        if not candidates:
            raise RuntimeError(
                "LibreOffice reported success but no .docx was found in output dir."
            )
        produced = str(candidates[0])
    return produced


def convert_to_docx(file_path: str, tmp_dir: str) -> tuple[str, list[str]]:
    """
    Convert any supported document to DOCX.

    Returns:
        (docx_path, warnings)  — docx_path is inside tmp_dir.
    Raises:
        EnvironmentError  — missing required system tool
        RuntimeError      — conversion failed
    """
    ext = Path(file_path).suffix.lower()
    warnings: list[str] = []
    docx_out = os.path.join(tmp_dir, "document.docx")

    # Already DOCX — just copy so we have a clean working copy
    if ext == ".docx":
        shutil.copy2(file_path, docx_out)
        return docx_out, warnings

    # Formats that need LibreOffice first
    if ext in LIBREOFFICE_REQUIRED:
        lo_docx = _libreoffice_to_docx(file_path, tmp_dir)
        shutil.move(lo_docx, docx_out)
        return docx_out, warnings

    # Formats pandoc handles natively
    if ext in PANDOC_NATIVE:
        warnings.extend(_pandoc_to_docx(file_path, docx_out))
        return docx_out, warnings

    # Unknown extension — try pandoc first, then LibreOffice
    warnings.append(
        f"Unknown extension '{ext}'. Trying pandoc first, then LibreOffice."
    )
    try:
        warnings.extend(_pandoc_to_docx(file_path, docx_out))
        return docx_out, warnings
    except (EnvironmentError, RuntimeError) as pandoc_err:
        warnings.append(f"pandoc failed: {pandoc_err}")

    try:
        lo_docx = _libreoffice_to_docx(file_path, tmp_dir)
        shutil.move(lo_docx, docx_out)
        return docx_out, warnings
    except (EnvironmentError, RuntimeError) as lo_err:
        warnings.append(f"LibreOffice failed: {lo_err}")

    raise RuntimeError(
        f"Cannot convert '{ext}' to DOCX.\n"
        "Neither pandoc nor LibreOffice could handle this format.\n"
        "Supported formats: " + ", ".join(sorted(PANDOC_NATIVE | LIBREOFFICE_REQUIRED | {".docx"}))
    )


# ---------------------------------------------------------------------------
# DOCX → Markdown + image extraction
# ---------------------------------------------------------------------------

def docx_to_markdown(docx_path: str, tmp_dir: str, extract_images: bool = True) -> tuple[str, list[str]]:
    """
    Convert a DOCX file to Markdown.

    If extract_images=True, images are written to tmp_dir/media/ and their
    paths are returned in the second element.

    Returns:
        (markdown_text, [absolute_image_paths])
    """
    pandoc = _require_pandoc()
    md_path = os.path.join(tmp_dir, "output.md")
    media_dir = os.path.join(tmp_dir, "media")

    cmd = [pandoc, docx_path, "-o", md_path, "--wrap=none"]
    if extract_images:
        os.makedirs(media_dir, exist_ok=True)
        cmd.append(f"--extract-media={media_dir}")

    result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    if result.returncode != 0:
        raise RuntimeError(
            f"pandoc DOCX→Markdown failed (exit {result.returncode}):\n{result.stderr.strip()}"
        )

    with open(md_path, "r", encoding="utf-8") as fh:
        markdown = fh.read()

    image_paths: list[str] = []
    if extract_images and os.path.isdir(media_dir):
        for root, _, files in os.walk(media_dir):
            for fname in sorted(files):
                if Path(fname).suffix.lower() in IMAGE_EXTENSIONS:
                    image_paths.append(os.path.join(root, fname))

    return markdown, image_paths


def _encode_images(image_paths: list[str]) -> list[dict]:
    """
    Base64-encode a list of image file paths.

    Returns a list of dicts:
        {
            "filename": str,       # original filename
            "mime_type": str,      # e.g. "image/png"
            "data": str,           # base64-encoded content
        }
    """
    encoded = []
    for path in image_paths:
        mime, _ = mimetypes.guess_type(path)
        if not mime:
            mime = "image/png"
        with open(path, "rb") as fh:
            b64 = base64.b64encode(fh.read()).decode("utf-8")
        encoded.append({
            "filename": Path(path).name,
            "mime_type": mime,
            "data": b64,
        })
    return encoded


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def read_document(file_path: str, extract_images: bool = True) -> dict:
    """
    Read a document and return its content as Markdown plus any embedded images.

    Parameters:
        file_path      — absolute or relative path to the document
        extract_images — whether to extract and base64-encode embedded images

    Returns a dict:
        {
            "markdown": str,          # full document as Markdown
            "images":   list[dict],   # [{filename, mime_type, data}] or []
            "warnings": list[str],    # non-fatal messages from conversion
            "tmp_dir":  str,          # MUST be passed to cleanup() when done
        }

    Raises:
        FileNotFoundError  — file_path does not exist
        EnvironmentError   — required tool (pandoc / LibreOffice) not installed
        RuntimeError       — conversion or extraction failed
    """
    file_path = os.path.abspath(file_path)
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"Document not found: {file_path}")

    tmp_dir = tempfile.mkdtemp(prefix="docx_cowork_")
    try:
        docx_path, warnings = convert_to_docx(file_path, tmp_dir)
        markdown, image_paths = docx_to_markdown(docx_path, tmp_dir, extract_images)
        images = _encode_images(image_paths) if extract_images else []

        return {
            "markdown": markdown,
            "images": images,
            "warnings": warnings,
            "tmp_dir": tmp_dir,
        }
    except Exception:
        # Clean up on error so callers don't have to
        shutil.rmtree(tmp_dir, ignore_errors=True)
        raise


def cleanup(tmp_dir: str) -> None:
    """
    Delete the temporary directory created by read_document().
    Always call this when you are done with the result, even if an error occurred.
    """
    if tmp_dir:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Read a document and print its Markdown content.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    p.add_argument("file", help="Path to the document to read")
    p.add_argument(
        "--no-images",
        action="store_true",
        help="Skip image extraction (faster, smaller output)",
    )
    p.add_argument(
        "--output-dir",
        metavar="DIR",
        help="Save output.md and extracted images here instead of printing to stdout",
    )
    p.add_argument(
        "--json",
        action="store_true",
        help="Print the full result dict as JSON (includes base64 image data)",
    )
    return p


def main() -> None:
    args = _build_parser().parse_args()
    extract = not args.no_images

    try:
        result = read_document(args.file, extract_images=extract)
    except (FileNotFoundError, EnvironmentError, RuntimeError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)

    try:
        if args.json:
            import json
            # Don't include tmp_dir in JSON output
            out = {k: v for k, v in result.items() if k != "tmp_dir"}
            print(json.dumps(out, indent=2))

        elif args.output_dir:
            out_dir = os.path.abspath(args.output_dir)
            os.makedirs(out_dir, exist_ok=True)

            md_out = os.path.join(out_dir, "output.md")
            with open(md_out, "w", encoding="utf-8") as fh:
                fh.write(result["markdown"])
            print(f"Markdown → {md_out}")

            for img in result["images"]:
                img_out = os.path.join(out_dir, img["filename"])
                with open(img_out, "wb") as fh:
                    fh.write(base64.b64decode(img["data"]))
                print(f"Image    → {img_out}  ({img['mime_type']})")

        else:
            print(result["markdown"])
            if result["images"]:
                print(f"\n[{len(result['images'])} image(s) extracted — "
                      "use --output-dir to save them]", file=sys.stderr)

        if result["warnings"]:
            print("\nWarnings:", file=sys.stderr)
            for w in result["warnings"]:
                print(f"  • {w}", file=sys.stderr)

    finally:
        cleanup(result.get("tmp_dir", ""))


if __name__ == "__main__":
    main()
