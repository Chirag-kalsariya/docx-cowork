# docx-cowork · Read

Read, summarize, analyze, or extract content from a word-processing document.

**Helper:** `helpers/read_document.py`
**Supported formats:** `.docx` `.doc` `.odt` `.rtf` `.pages`

"If the file is not one of those formats, this skill cannot help. Tell the user and ask the agent to check other available skills."

---

## Workflow

### Step 1 — Check pandoc

```bash
pandoc --version
```

If pandoc is missing, stop and tell the user:

"pandoc is not installed. Please install it and try again."

| Platform | Command |
|---|---|
| macOS | `brew install pandoc` |
| Ubuntu / Debian | `sudo apt install pandoc` |
| Windows | `winget install JohnMacFarlane.Pandoc` or https://pandoc.org/installing.html |

---

### Step 2 — Identify the file and format

```bash
ls -lh "<file_path>"
```

- File not found → ask the user for the correct path.
- Extension not in `.docx .doc .odt .rtf .pages` → tell the user this format is not supported, list what is, and ask the agent to check other skills.

For `.doc` or `.pages`, also check whether LibreOffice is available:

| Platform | Check command |
|---|---|
| macOS | `ls /Applications/LibreOffice.app/Contents/MacOS/soffice 2>/dev/null && echo found` |
| Linux | `which libreoffice || which soffice` |
| Windows | `where soffice.exe` |

If LibreOffice is not found, stop and tell the user:

"LibreOffice is required to convert .<ext> files. Please install it and try again."

| Platform | Install |
|---|---|
| macOS | `brew install --cask libreoffice` |
| Ubuntu / Debian | `sudo apt install libreoffice` |
| Windows | https://www.libreoffice.org/download/ |

---

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

---

### Step 4 — Read the Markdown

`result["markdown"]` contains the full document text. Use it to answer the user's question.

---

### Step 5 — Handle images

| Model capability | Action |
|---|---|
| Supports vision (multimodal) | Pass `{mime_type, data}` from each `result["images"]` entry to the model alongside the Markdown |
| Text-only | Skip images and tell the user "N image(s) were found but cannot be displayed in this context" |

---

### Step 6 — Report warnings

If `result["warnings"]` is non-empty, show them as informational notes.

---

### Step 7 — Clean up (REQUIRED, always)

```python
cleanup(result["tmp_dir"])
```

Bash fallback (all platforms):

```bash
# macOS / Linux
rm -rf /tmp/docx_cowork_*

# Windows (PowerShell)
Remove-Item "$env:TEMP\docx_cowork_*" -Recurse -Force
```

---

## Conversion Reference

| Extension | Conversion path |
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
| `pandoc` not found | Show install instructions for all platforms, stop |
| `LibreOffice` not found (`.doc` / `.pages`) | Show install instructions for all platforms, stop |
| File not found | Ask user for the correct path |
| `ValueError` — unsupported format | List supported formats, ask agent to check other skills |
| Conversion `RuntimeError` | Show the error message, suggest checking the file is not corrupted |
