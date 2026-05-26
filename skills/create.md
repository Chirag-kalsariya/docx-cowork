# docx-cowork · Create

> **Status: Coming soon.**  
> This sub-skill is not yet implemented.

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

---

*When this skill is implemented, full workflow details will be added here.*
