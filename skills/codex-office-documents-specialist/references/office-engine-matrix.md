# Office Engine Matrix

## Use The Lightest Reliable Engine

| Task | Word | Excel | PowerPoint |
|---|---|---|---|
| Inspect unknown file safely | `inspect_docx.py` | `inspect_xlsx.py` | `inspect_pptx.py` |
| Generate new editable file from notes | `markdown_to_docx.py` | `markdown_to_xlsx.py` | `structured_pptx.py` |
| Fill a reusable template | `fill_docx_template.py` | prefer sheet copy or native Excel template workflow | prefer `.pptx` template + `structured_pptx.py` |
| Preserve exact native layout | Word COM | Excel COM | PowerPoint COM |
| Update fields or TOC | Word COM | n/a | n/a |
| Run macro | Word COM | Excel COM | PowerPoint COM |
| Manage charts or pivots | limited in bundled scripts | Excel COM | PowerPoint COM or python-pptx for simple charts |
| Export exact PDF | Word COM | Excel COM | PowerPoint COM |
| Cross-platform enterprise workflow | Aspose-style reference | Aspose-style reference | Aspose-style reference |

## Fast Rules

- Use native COM when Office itself is the rendering authority.
- Use Python generation when the file is mostly structural and should stay editable.
- Use template-first workflows when branding, themes, or placeholder layout already exist.
- Use external slide APIs only when the user explicitly allows a third-party service.
- Treat Aspose as the broadest architecture reference, not as an assumed local runtime.

## Repo-Derived Takeaways

- OfficeMCP and cs-office-mcp-server are strongest for native Office automation and macros on Windows.
- mcp-ms-office-documents is strongest for content generation, template fill, and storage handoff.
- Office-Word-MCP-Server is strongest for fine-grained python-docx feature coverage in Word.
- 2slides is strongest for theme-driven or narrated deck generation through an external API.
- Aspose MCP is strongest for unified cross-format architecture, structured outputs, session design, and security boundaries.
