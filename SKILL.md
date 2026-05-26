---
name: docx-cowork-read
description: >
  Read and extract content from DOCX, DOC, Apple Pages, ODT, RTF, HTML, and
  other document formats. Converts the document to Markdown and optionally
  extracts embedded images for vision-capable models. Use this skill whenever
  the user asks you to read, summarize, analyze, or extract content from a
  document file.
tools: [bash, read_file, write_file]
applyTo: "**"
---

# docx-cowork · Read Skill

Use this skill to read any document the user provides. You will convert it to
Markdown (and optionally extract images), analyze the content, then clean up
all temporary files.

---

## Supported Formats

| Format | Extension(s) | Conversion path |
|---|---|---|
| Microsoft Word (modern) | `.docx` | pandoc (native) |
| OpenDocument Text | `.odt` | pandoc (native) |
| Rich Text Format | `.rtf` | pandoc (native) |
| HTML / Web page | `.html`, `.htm` | pandoc (native) |
| EPUB | `.epub` | pandoc (native) |
| reStructuredText | `.rst` | pandoc (native) |
| LaTeX | `.tex`, `.latex` | pandoc (native) |
| Microsoft Word (legacy) | `.doc` | LibreOffice → DOCX → pandoc |
| Apple Pages | `.pages` | LibreOffice → DOCX → pandoc |
| Apple Numbers | `.numbers` | LibreOffice → DOCX → pandoc |
| PowerPoint (legacy) | `.ppt` | LibreOffice → DOCX → pandoc |

> **Google Docs**: these live in the cloud. Ask the user to export the file
> first (File → Download → Microsoft Word `.docx`) then use this skill.
>
> **Unsupported format**: if both pandoc and LibreOffice fail, report the
> error and tell the user which formats are supported.

---

## Workflow

Follow these steps **in order** every time you read a document.

### 1 · Check dependencies

Before anything else, confirm pandoc is available:

```bash
pandoc --version
```

If pandoc is missing, tell the user and stop:

```
pandoc is not installed.
  macOS:   brew install pandoc
  Ubuntu:  sudo apt install pandoc
  Windows: https://pandoc.org/installing.html
```

### 2 · Identify the file

Confirm the file path exists and note the extension.

```bash
ls -lh "<file_path>"
```

- If the file does not exist, ask the user for the correct path.
- For `.doc` or `.pages`, check whether LibreOffice is available:

```bash
which libreoffice || which soffice || \
  ls /Applications/LibreOffice.app/Contents/MacOS/soffice 2>/dev/null \
  && echo "found" || echo "not found"
```

If LibreOffice is missing and the file is `.doc` or `.pages`, tell the user:

```
LibreOffice is required to convert .<ext> files.
  macOS:   brew install --cask libreoffice
  Ubuntu:  sudo apt install libreoffice
  Windows: https://www.libreoffice.org/download/
```

### 3 · Run the helper

Call the Python helper. It handles all conversion and image extraction
automatically:

```bash
python helpers/read_document.py "<file_path>"
```

**With image extraction** (default — recommended when the model supports vision):

```bash
python helpers/read_document.py "<file_path>" --output-dir /tmp/docx_out_$$
```

**Without images** (faster, text-only models):

```bash
python helpers/read_document.py "<file_path>" --no-images
```

**Or call the Python API directly** (preferred when you are already in a Python
context):

```python
from helpers.read_document import read_document, cleanup

result = read_document("<file_path>", extract_images=True)
# result["markdown"]  → full document as Markdown string
# result["images"]    → list of {filename, mime_type, data(base64)} dicts
# result["warnings"]  → list of non-fatal conversion warnings
# result["tmp_dir"]   → temp directory path — pass to cleanup() when done
```

### 4 · Read the Markdown

`result["markdown"]` (or `output.md`) contains the full document content.
Read it and answer the user's question.

### 5 · Handle images

Check whether the model you are running in accepts image input.

**If the model supports vision** (multimodal):
- Iterate over `result["images"]`
- Each image has `mime_type` and base64 `data`
- Pass images to the model alongside the Markdown

**If the model does NOT support vision**:
- Skip images entirely
- Mention to the user that `N image(s) were found but cannot be displayed in
  this context

### 6 · Report warnings

If `result["warnings"]` is non-empty, show them to the user as informational
notes (e.g. "pandoc produced a warning about unsupported style X").

### 7 · Clean up — ALWAYS DO THIS

Delete every temporary file after you are done, regardless of success or error.

**Python API:**
```python
cleanup(result["tmp_dir"])
```

**CLI / bash:**
```bash
rm -rf /tmp/docx_out_*
rm -rf /tmp/docx_cowork_*
```

The helper always names temp dirs with the prefix `docx_cowork_` so the glob
above is safe to use as a fallback sweep.

---

## Error Handling Reference

| Situation | Action |
|---|---|
| `pandoc` not found | Tell user to install pandoc, stop |
| `LibreOffice` not found (needed for .doc/.pages) | Tell user to install LibreOffice, stop |
| File not found | Ask user for correct path |
| Conversion failed | Show the error message, list supported formats |
| Unknown extension | Try pandoc first, then LibreOffice; report failure if both fail |
| Images found but model is text-only | Process Markdown only, mention skipped images |

---

## Quick Reference — Helper API

```python
from helpers.read_document import read_document, cleanup

# Read and extract
result = read_document(
    file_path="report.docx",
    extract_images=True,   # False = skip image extraction
)

# Access content
markdown: str       = result["markdown"]
images:   list      = result["images"]    # [{filename, mime_type, data}]
warnings: list[str] = result["warnings"]

# Cleanup — REQUIRED
cleanup(result["tmp_dir"])
```

```bash
# CLI usage
python helpers/read_document.py report.docx                     # print Markdown
python helpers/read_document.py report.docx --output-dir ./out  # save files
python helpers/read_document.py report.docx --no-images         # skip images
python helpers/read_document.py report.docx --json              # full JSON output
```
