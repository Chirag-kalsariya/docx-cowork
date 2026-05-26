# docx-cowork · Read

Read, summarize, analyze, or extract content from a word-processing document.

**Helper:** `helpers/read_document.py`  
**Supported formats:** `.docx` `.doc` `.odt` `.rtf` `.pages`

---

## Workflow

### Step 1 — Check pandoc

```bash
pandoc --version
```

Missing? Stop and tell the user:
```
pandoc is required.  brew install pandoc  /  sudo apt install pandoc
```

### Step 2 — Identify the file and format

```bash
ls -lh "<file_path>"
```

- File missing → ask the user for the correct path.
- Extension not in `.docx .doc .odt .rtf .pages` → tell the user this format
  is not supported and list what is. Do not proceed.

For `.doc` or `.pages`, also check LibreOffice:

```bash
which libreoffice || which soffice || \
  ls /Applications/LibreOffice.app/Contents/MacOS/soffice 2>/dev/null \
  && echo "found" || echo "not found"
```

If not found:
```
LibreOffice is required for .<ext> files.
  macOS:   brew install --cask libreoffice
  Ubuntu:  sudo apt install libreoffice
```

### Step 3 — Run the helper

**Python API (preferred):**

```python
from helpers.read_document import read_document, cleanup

result = read_document("<file_path>", extract_images=True)
# result["markdown"]  → full document as Markdown
# result["images"]    → [{filename, mime_type, data(base64)}]  or []
# result["warnings"]  → non-fatal conversion messages
# result["tmp_dir"]   → pass to cleanup() when done
```

**CLI alternative:**

```bash
# Print Markdown to stdout
python3 helpers/read_document.py "<file_path>"

# Save Markdown + images to a directory
python3 helpers/read_document.py "<file_path>" --output-dir /tmp/docx_out_$$

# Text-only (skip image extraction)
python3 helpers/read_document.py "<file_path>" --no-images
```

### Step 4 — Read the Markdown

`result["markdown"]` contains the full document text. Use it to answer the
user's question.

### Step 5 — Handle images

| Model capability | Action |
|---|---|
| Supports vision (multimodal) | Pass `{mime_type, data}` from each `result["images"]` entry to the model alongside the Markdown |
| Text-only | Skip images; tell user "N image(s) found but cannot be displayed in this context" |

### Step 6 — Report warnings

If `result["warnings"]` is non-empty, show them as informational notes.

### Step 7 — Clean up (REQUIRED, always)

```python
cleanup(result["tmp_dir"])
```

Or bash fallback:
```bash
rm -rf /tmp/docx_cowork_*
```

---

## Conversion Reference

| Extension | Path |
|---|---|
| `.docx` | pandoc (native) |
| `.odt` | pandoc (native) |
| `.rtf` | pandoc (native) |
| `.doc` | LibreOffice → `.docx` → pandoc |
| `.pages` | LibreOffice → `.docx` → pandoc |

---

## Error Reference

| Error | Response |
|---|---|
| `pandoc` not found | Install instructions, stop |
| `LibreOffice` not found (`.doc`/`.pages`) | Install instructions, stop |
| File not found | Ask user for path |
| `ValueError` — unsupported format | List supported formats, stop |
| Conversion `RuntimeError` | Show error, suggest checking the file |
