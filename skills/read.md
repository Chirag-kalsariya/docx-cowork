# docx-cowork · Read

Read, summarize, analyze, or extract content from a word-processing document.

**Helper:** `helpers/read_document.py`
**Supported formats:** `.docx` `.doc` `.odt` `.rtf` `.pages` `.epub`
**Dependencies:** pandoc (always), LibreOffice (only for `.doc` and `.pages`)
**No pip install needed** — the helper uses only the Python standard library.

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

Confirm the file exists and note its extension.

- Extension not in `.docx .doc .odt .rtf .pages .epub` → tell the user this format is not supported, list what is, and ask the agent to check other skills.

For `.doc` or `.pages`, also check whether LibreOffice is available:

| Platform | Check command |
|---|---|
| macOS | `ls /Applications/LibreOffice.app/Contents/MacOS/soffice 2>/dev/null && echo found` |
| Linux | `which libreoffice \|\| which soffice` |
| Windows | `where soffice.exe` |

If LibreOffice is not found, stop and tell the user:

"LibreOffice is required to convert .<ext> files. Please install it and try again."

| Platform | Install |
|---|---|
| macOS | `brew install --cask libreoffice` |
| Ubuntu / Debian | `sudo apt install libreoffice` |
| Windows | https://www.libreoffice.org/download/ |

---

### Step 3 — Extract the document

Run the helper to convert the document and extract its contents to a temp directory. The helper never returns the full text inline — it gives you file paths and stats so you control how much you load.

```python
from helpers.read_document import extract, read_lines, cleanup

info = extract("<file_path>", extract_images=True)
# info["markdown_path"]  → path to output.md on disk
# info["image_paths"]    → list of image file paths on disk
# info["stats"]          → {char_count, line_count, image_count, size_bytes, is_large}
# info["warnings"]       → non-fatal messages
# info["tmp_dir"]        → pass to cleanup() when done
```

CLI alternative:

```bash
python3 helpers/read_document.py "<file_path>"
```

---

### Step 4 — Check stats before reading

Always check `info["stats"]` first to decide how much to load.

```python
stats = info["stats"]
# stats["char_count"]  — total characters in the markdown
# stats["line_count"]  — total lines
# stats["image_count"] — number of embedded images
# stats["is_large"]    — True if char_count > 20,000
```

| Document size | What to do |
|---|---|
| `is_large` is False (small, < 20K chars) | Read the full markdown in one call |
| `is_large` is True (large) | Read only the sections relevant to the user's question |

---

### Step 5 — Read the markdown (on demand)

**Small document — read everything:**

```python
text = read_lines(info["markdown_path"])
```

**Large document — read only what you need:**

```python
# Read the first 100 lines to understand structure / table of contents
preview = read_lines(info["markdown_path"], start=0, end=100)

# Read a specific section by line range (0-based, end is exclusive)
section = read_lines(info["markdown_path"], start=200, end=350)
```

Use the preview to locate the relevant section, then read only those lines. Do not load the entire file into context if it is large.

---

### Step 6 — Handle images (only if needed)

Only load images if the user's task actually requires looking at visual content (charts, diagrams, screenshots, etc.).

| Model capability | Action |
|---|---|
| Supports vision AND images are needed | Read each file in `info["image_paths"]` and pass to the model |
| Text-only model | Skip images; tell user "N image(s) found but cannot be displayed in this context" |
| Images not needed for the task | Skip images entirely |

---

### Step 7 — Report warnings

If `info["warnings"]` is non-empty, show them as informational notes to the user.

---

### Step 8 — Clean up (REQUIRED, always)

Delete the temp directory when you are done — whether the task succeeded or failed.

```python
cleanup(info["tmp_dir"])
```

Bash fallback:

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
| `.epub` | pandoc (native) |
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
| Conversion `RuntimeError` | Show the error message, suggest the file may be corrupted |
