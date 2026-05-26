---
name: docx-cowork
description: "
  Create, read, and update DOCX, DOC, ODT, RTF, and Apple Pages documents.
  Use this skill for any task involving word-processing files. Load the
  relevant sub-skill based on what the user wants to do."
license: "MIT"
---

# docx-cowork

Agent skill for working with word-processing documents.

## Supported Formats

| Format | Extension |
|---|---|
| Microsoft Word (modern) | `.docx` |
| Microsoft Word (legacy) | `.doc` |
| OpenDocument Text | `.odt` |
| Rich Text Format | `.rtf` |
| Apple Pages | `.pages` |
| E-book | `.epub` |

"Not supported: presentations (.ppt, .pptx, .key), spreadsheets (.xlsx, .numbers), web pages (.html), or markup files (.md, .rst). This skill cannot handle those. Ask the agent to check other available skills — one of them may be able to help."

---

## Routing — Load the Right Sub-Skill

Read the user's intent and load **only** the sub-skill you need.

| User wants to… | Load |
|---|---|
| Read, summarize, extract, or analyze an existing document | [`skills/read.md`](skills/read.md) |
| Create a new document from scratch | [`skills/create.md`](skills/create.md) |
| Edit, update, or modify an existing document | [`skills/update.md`](skills/update.md) |

Load the sub-skill file, follow its workflow, then return to this index if
you need to chain operations (e.g., read a document and then update it).

---

## Shared Rules (apply to all sub-skills)

1. **Always clean up.** Delete the temp directory using `cleanup(info["tmp_dir"])` with the exact path when done — whether the task succeeded or failed. The OS does not reliably auto-clean temp files on all platforms. Use Python's `cleanup()`, not glob-based shell commands.
2. **Check dependencies first.** Confirm `pandoc` (and `LibreOffice` when needed) are available before starting any conversion. No pip install is required — the helpers use only the Python standard library.
3. **Unsupported format?** Tell the user clearly, list what IS supported, and ask the agent to check other skills.
4. **Google Docs?** Ask the user to export to `.docx` first (File → Download → Microsoft Word `.docx`), then proceed.
