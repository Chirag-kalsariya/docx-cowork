# docx-cowork · Create

"Status: Coming soon. This sub-skill is not yet implemented. If you need to create a document right now, check other available skills — one of them may be able to help."

Create a new word-processing document (`.docx`) from structured content,
a template, or instructions provided by the user.

**Planned helper:** `helpers/create_document.py`
**Output format:** `.docx` (Microsoft Word Open XML)

---

## Planned Workflow (draft)

1. Gather content from the user (title, sections, body text, tables, etc.)
2. Build the document in memory using `python-docx`
3. Apply any styles or formatting requested
4. Save to the user-specified output path
5. Report success with the file path and size
6. Clean up any temporary files

Works on macOS, Linux, and Windows.

---

"When this sub-skill is implemented, the full workflow will be added here."
