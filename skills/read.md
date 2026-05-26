# docx-cowork · Read

Read, summarize, analyze, or extract content from a word-processing document.

**Helper:** `helpers/read_document.py`
**Supported formats:** `.docx` `.doc` `.odt` `.rtf` `.pages` `.epub`
**Dependencies:** pandoc (always), LibreOffice (only for `.doc` and `.pages`)
**No pip install needed** — the helper uses only the Python standard library.

**Temp directory location (created automatically per platform):**

| Platform | Location | Example |
|---|---|---|
| macOS | `/var/folders/.../T/` | `/var/folders/xy/.../T/docx_cowork_abc123/` |
| Linux | `/tmp/` | `/tmp/docx_cowork_abc123/` |
| Windows | `%TEMP%` | `C:\Users\you\AppData\Local\Temp\docx_cowork_abc123\` |

"The OS does not reliably auto-clean these: Linux only clears /tmp on reboot, macOS clears /var/folders after several days, Windows never auto-cleans %TEMP%. Always call cleanup() with the exact path returned by extract() — do not use glob-based rm -rf commands as agents typically block wildcard deletes for safety."

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

### Step 4 — Decide what to read before reading anything

After extracting, look at `info["stats"]` and the user's task together. Make all decisions here before loading any content.

```python
stats = info["stats"]
# stats["char_count"]  — total characters in the markdown
# stats["line_count"]  — total lines
# stats["image_count"] — number of embedded images
# stats["is_large"]    — True if char_count > 20,000
```

**Decision 1 — How much text to read:**

| Situation | Decision |
|---|---|
| `is_large` is False | Read the full file in one call |
| `is_large` is True AND user wants a full summary or full extraction | Read the full file in chunks (100–200 lines at a time), do not skip anything |
| `is_large` is True AND user has a specific question or wants a specific section | Read the first 100 lines to get structure, then read only the relevant section(s) |
| `is_large` is True AND user wants just metadata (title, author, page count, etc.) | Read the first 50 lines only |

**Decision 2 — Whether to read images:**

| Situation | Decision |
|---|---|
| `image_count` is 0 | Skip — no images to load |
| User's task is text-only (summarize, extract text, answer a text question) | Skip images |
| User's task involves visuals (describe a chart, read a diagram, check a screenshot) | Load images |
| Model does not support vision | Skip images; tell user "N image(s) were found but cannot be displayed in this context" |
| Unsure whether images are relevant | Skip for now; mention to the user that N image(s) exist and offer to inspect them |

---

### Step 5 — Read the markdown (on demand)

**Small document — read everything:**

```python
text = read_lines(info["markdown_path"])
```

**Large document — full summary or full extraction (read in chunks):**

```python
chunk_size = 150
total_lines = info["stats"]["line_count"]
for start in range(0, total_lines, chunk_size):
    chunk = read_lines(info["markdown_path"], start=start, end=start + chunk_size)
    # process chunk
```

**Large document — specific question or section:**

```python
# Step 1: read first 100 lines to understand structure
preview = read_lines(info["markdown_path"], start=0, end=100)

# Step 2: identify the relevant line range from the preview
# Step 3: read only that range
section = read_lines(info["markdown_path"], start=200, end=350)
```

---

### Step 6 — Read images (on demand, only if decided in Step 4)

```python
# Read each image file and pass to the model
for image_path in info["image_paths"]:
    # read image_path as bytes and pass to vision model
    pass
```

---

### Step 7 — Report warnings

If `info["warnings"]` is non-empty, show them as informational notes to the user.

---

### Step 8 — Clean up (REQUIRED, always)

The temp directory is NOT automatically cleaned by the OS on all platforms. Always delete it when done using the exact path returned by `extract()`:

```python
cleanup(info["tmp_dir"])
```

`cleanup()` uses Python's `shutil.rmtree` with the exact path — it works on macOS, Linux, and Windows without needing shell permissions. Do not use glob-based bash commands like `rm -rf /tmp/docx_cowork_*` — agents typically block wildcard deletes.

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
