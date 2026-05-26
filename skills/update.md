# docx-cowork · Update

> **Status: Coming soon.**  
> This sub-skill is not yet implemented.

Edit or update an existing word-processing document — change text, add or
remove sections, update tables, replace content, and more.

**Planned helper:** `helpers/update_document.py`  
**Supported input formats:** `.docx` `.doc` `.odt` `.rtf` `.pages`  
**Output format:** `.docx`

---

## Planned Workflow (draft)

1. Read the existing document using the **read** sub-skill
2. Identify the sections or elements the user wants to change
3. Apply changes using `python-docx`
4. Save the updated document (overwrite or new file as user prefers)
5. Clean up any temporary files

---

*When this skill is implemented, full workflow details will be added here.*
