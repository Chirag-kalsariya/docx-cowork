#!/usr/bin/env python3
"""
read_document.py — Document reading helper for the docx-cowork agent skill.

Uses only Python standard library + pandoc (external) + LibreOffice (external,
only for .doc and .pages). No pip install required.

Supported word-processing formats:
  .docx .odt .rtf .epub  → pandoc (native, no extra tool needed)
  .doc  .pages           → LibreOffice --headless → .docx → pandoc

The helper extracts the document to a temp directory and returns file paths
and stats. It never returns inline content — the agent reads only the sections
it needs, keeping the context window small.

Python API:
    from helpers.read_document import extract, read_lines, cleanup

    info = extract("report.docx", extract_images=True)
    # info["markdown_path"]  → path to output.md  (read it yourself, in parts)
    # info["image_paths"]    → ["/tmp/docx_cowork_.../media/img1.png", ...]
    # info["stats"]          → {char_count, line_count, image_count, size_bytes}
    # info["warnings"]       → non-fatal messages
    # info["tmp_dir"]        → pass to cleanup() when done

    # Read a specific range of lines (0-based, end is exclusive):
    text = read_lines(info["markdown_path"], start=0, end=100)

    cleanup(info["tmp_dir"])   # ALWAYS call when done

CLI:
    python3 helpers/read_document.py file.docx           # print stats
    python3 helpers/read_document.py file.docx --lines 1-50   # print lines
    python3 helpers/read_document.py file.docx --all          # print full markdown
    python3 helpers/read_document.py file.docx --output-dir ./out  # save files
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

# ---------------------------------------------------------------------------
# Format classification
# ---------------------------------------------------------------------------

SUPPORTED_FORMATS: frozenset[str] = frozenset({
    ".docx",   # Microsoft Word (modern Open XML)
    ".doc",    # Microsoft Word (legacy binary)
    ".odt",    # OpenDocument Text
    ".rtf",    # Rich Text Format
    ".pages",  # Apple Pages
    ".epub",   # E-book
})

PANDOC_NATIVE: frozenset[str] = frozenset({
    ".docx", ".odt", ".rtf", ".epub",
})

LIBREOFFICE_REQUIRED: frozenset[str] = frozenset({
    ".doc",    # legacy Word binary — pandoc 3.x dropped this format
    ".pages",  # Apple Pages — never supported by pandoc
})

IMAGE_EXTENSIONS: frozenset[str] = frozenset({
    ".png", ".jpg", ".jpeg", ".gif",
    ".webp", ".bmp", ".tiff", ".tif", ".svg",
})

# Documents smaller than this are considered "small" — safe to read in full.
# Larger documents should be read in sections to protect the context window.
SMALL_DOC_CHARS = 20_000


# ---------------------------------------------------------------------------
# System tool discovery
# ---------------------------------------------------------------------------

def _require_pandoc() -> str:
    path = shutil.which("pandoc")
    if not path:
        raise EnvironmentError(
            "pandoc is not installed or not on PATH.\n"
            "  macOS:   brew install pandoc\n"
            "  Ubuntu:  sudo apt install pandoc\n"
            "  Windows: winget install JohnMacFarlane.Pandoc\n"
            "           or https://pandoc.org/installing.html"
        )
    return path


def _find_libreoffice() -> str | None:
    candidates = [
        "libreoffice",
        "soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",    # macOS
        "/usr/bin/libreoffice",                                     # Linux
        "/usr/bin/soffice",                                         # Linux alt
        r"C:\Program Files\LibreOffice\program\soffice.exe",        # Windows
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",  # Windows 32-bit
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
    lo = _find_libreoffice()
    if not lo:
        raise EnvironmentError(
            "LibreOffice is not installed. It is required to convert "
            f"'{Path(input_path).suffix}' files.\n"
            "  macOS:   brew install --cask libreoffice\n"
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
    stem = Path(input_path).stem
    produced = os.path.join(out_dir, stem + ".docx")
    if not os.path.isfile(produced):
        candidates = list(Path(out_dir).glob("*.docx"))
        if not candidates:
            raise RuntimeError(
                "LibreOffice reported success but no .docx was found in output dir."
            )
        produced = str(candidates[0])
    return produced


def _convert_to_docx(file_path: str, tmp_dir: str) -> tuple[str, list[str]]:
    ext = Path(file_path).suffix.lower()
    warnings: list[str] = []
    docx_out = os.path.join(tmp_dir, "document.docx")

    if ext not in SUPPORTED_FORMATS:
        raise ValueError(
            f"Unsupported file format: '{ext}'.\n"
            "docx-cowork supports word-processing documents only.\n"
            "Supported: " + ", ".join(sorted(SUPPORTED_FORMATS))
        )
    if ext == ".docx":
        shutil.copy2(file_path, docx_out)
        return docx_out, warnings
    if ext in LIBREOFFICE_REQUIRED:
        lo_docx = _libreoffice_to_docx(file_path, tmp_dir)
        shutil.move(lo_docx, docx_out)
        return docx_out, warnings
    # PANDOC_NATIVE
    warnings.extend(_pandoc_to_docx(file_path, docx_out))
    return docx_out, warnings


def _docx_to_markdown(docx_path: str, tmp_dir: str, extract_images: bool) -> tuple[str, list[str]]:
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

    image_paths: list[str] = []
    if extract_images and os.path.isdir(media_dir):
        for root, _, files in os.walk(media_dir):
            for fname in sorted(files):
                if Path(fname).suffix.lower() in IMAGE_EXTENSIONS:
                    image_paths.append(os.path.join(root, fname))

    return md_path, image_paths


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def extract(file_path: str, extract_images: bool = True) -> dict:
    """
    Extract a document to a temp directory and return paths + stats.

    Returns:
        {
            "markdown_path": str,     # absolute path to output.md
            "image_paths":  list[str],# absolute paths to extracted images
            "stats": {
                "char_count":  int,   # characters in the markdown
                "line_count":  int,   # lines in the markdown
                "image_count": int,   # number of extracted images
                "size_bytes":  int,   # byte size of the markdown file
                "is_large":    bool,  # True if > SMALL_DOC_CHARS chars
            },
            "warnings": list[str],
            "tmp_dir":  str,          # pass to cleanup() when done
        }

    Raises:
        FileNotFoundError  — file does not exist
        ValueError         — unsupported format
        EnvironmentError   — pandoc or LibreOffice not installed
        RuntimeError       — conversion failed
    """
    file_path = os.path.abspath(file_path)
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"Document not found: {file_path}")

    tmp_dir = tempfile.mkdtemp(prefix="docx_cowork_")
    try:
        docx_path, warnings = _convert_to_docx(file_path, tmp_dir)
        md_path, image_paths = _docx_to_markdown(docx_path, tmp_dir, extract_images)

        size_bytes = os.path.getsize(md_path)
        with open(md_path, "r", encoding="utf-8") as fh:
            content = fh.read()
        char_count = len(content)
        line_count = content.count("\n") + 1

        return {
            "markdown_path": md_path,
            "image_paths": image_paths,
            "stats": {
                "char_count":  char_count,
                "line_count":  line_count,
                "image_count": len(image_paths),
                "size_bytes":  size_bytes,
                "is_large":    char_count > SMALL_DOC_CHARS,
            },
            "warnings": warnings,
            "tmp_dir": tmp_dir,
        }
    except Exception:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        raise


def read_lines(markdown_path: str, start: int = 0, end: int | None = None) -> str:
    """
    Read a range of lines from the extracted markdown file.

    Parameters:
        markdown_path — path from extract()["markdown_path"]
        start         — first line to return (0-based, inclusive)
        end           — last line to return (0-based, exclusive); None = end of file

    Returns the selected lines as a single string.
    """
    with open(markdown_path, "r", encoding="utf-8") as fh:
        lines = fh.readlines()
    selected = lines[start:end]
    return "".join(selected)


def cleanup(tmp_dir: str) -> None:
    """
    Delete the temporary directory created by extract().
    Always call this when done, even if an error occurred.
    """
    if tmp_dir:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Extract a document and inspect its content.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    p.add_argument("file", help="Path to the document")
    p.add_argument("--no-images", action="store_true", help="Skip image extraction")
    p.add_argument("--all", action="store_true", help="Print the full markdown to stdout")
    p.add_argument(
        "--lines", metavar="START-END",
        help="Print a line range, e.g. --lines 1-50 (1-based, inclusive)"
    )
    p.add_argument(
        "--output-dir", metavar="DIR",
        help="Save output.md and images to this directory"
    )
    return p


def main() -> None:
    args = _build_parser().parse_args()

    try:
        info = extract(args.file, extract_images=not args.no_images)
    except (FileNotFoundError, ValueError, EnvironmentError, RuntimeError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)

    try:
        s = info["stats"]
        print(f"Extracted: {info['markdown_path']}")
        print(f"Stats    : {s['char_count']:,} chars | {s['line_count']:,} lines | "
              f"{s['image_count']} image(s) | {s['size_bytes']:,} bytes")
        if s["is_large"]:
            print(f"Note     : large document (>{SMALL_DOC_CHARS:,} chars) — read in sections")
        if info["image_paths"]:
            print("Images   :")
            for p in info["image_paths"]:
                print(f"  {p}")

        if args.output_dir:
            out = os.path.abspath(args.output_dir)
            os.makedirs(out, exist_ok=True)
            shutil.copy2(info["markdown_path"], os.path.join(out, "output.md"))
            for p in info["image_paths"]:
                shutil.copy2(p, os.path.join(out, Path(p).name))
            print(f"Saved to : {out}")

        elif args.all:
            print("\n" + read_lines(info["markdown_path"]))

        elif args.lines:
            parts = args.lines.split("-")
            start = max(0, int(parts[0]) - 1)
            end   = int(parts[1]) if len(parts) > 1 else None
            print("\n" + read_lines(info["markdown_path"], start, end))

        if info["warnings"]:
            print("\nWarnings:", file=sys.stderr)
            for w in info["warnings"]:
                print(f"  • {w}", file=sys.stderr)

    finally:
        cleanup(info.get("tmp_dir", ""))


if __name__ == "__main__":
    main()
