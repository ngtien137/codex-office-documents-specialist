---
name: codex-office-documents-specialist
description: Comprehensive Microsoft Office document analysis, creation, editing, generation, and layout polishing across Word, Excel, PowerPoint, Outlook-style drafts, and PDF handoffs. Use when Codex needs to inspect or modify .docx, .xlsx, .pptx, .docm, .xlsm, .pptm, .eml, or Office-adjacent files; preserve layout on Windows; generate deliverables from Markdown or structured JSON; fill templates; analyze Office packages; or choose between COM automation, OOXML inspection, python-docx, openpyxl, python-pptx, remote slide APIs, or Aspose-style cross-platform workflows.
---

# Codex Office Documents Specialist

## Overview

Treat Office work as an engine-selection problem first and an editing problem second. Pick the narrowest engine that preserves the right fidelity:

- native Office COM on Windows when pagination, macros, charts, hidden app behavior, or exact layout must match what a human would produce inside Office
- OOXML inspection when the first job is understanding structure without editing
- python-docx, openpyxl, or python-pptx when the user wants deterministic generation from Markdown, JSON, or templates
- remote slide APIs only when the user explicitly allows third-party generation and visual theme quality matters more than offline editability
- Aspose-style workflows as an architectural reference for cross-platform or session-based operations, not as an assumed local dependency

Use bundled scripts first:

- `scripts/inspect_docx.py` for safe Word package inspection
- `scripts/profile_docx.py` for Word document type, topic, and tailored advice
- `scripts/inspect_xlsx.py` for safe Excel workbook inspection
- `scripts/inspect_pptx.py` for safe PowerPoint deck inspection
- `scripts/markdown_to_docx.py` for Word generation from Markdown
- `scripts/fill_docx_template.py` for DOCX placeholder fill
- `scripts/markdown_to_xlsx.py` for table-first Excel generation from Markdown
- `scripts/structured_pptx.py` for structured PowerPoint generation from JSON
- `scripts/word_com_actions.ps1` when Word itself must apply the final edit or export PDF
- `scripts/project_memory.py` to create project backups and append `.backup/memories.md` entries deterministically

## Standard Workflow

1. Inspect the document and identify the real constraint:
   - preserve native Office layout
   - batch edit content
   - create a new report, workbook, slide deck, or draft
   - fill a template
   - generate from Markdown or JSON
   - polish sections, spacing, tables, formulas, or slides
   - export a fixed-layout review artifact such as PDF
2. If the file is a Word document and the user wants analysis or editing, profile it.
   - use `scripts/profile_docx.py` to infer likely document kind and topic
   - report both formatting state and topic-specific advice
   - if it looks like a thesis, dissertation, graduation report, or academic report, explicitly ask whether the user wants academic-formatting defaults applied
3. Choose the engine with the rules above.
4. Before making any edit, present a numbered change plan to the user.
   - list each proposed change as a separate item
   - state what will change and why
   - let the user approve individual items or approve all
   - do not edit until the user confirms, unless the user explicitly asked for direct execution without confirmation
5. Before any in-place edit or overwrite, create a project backup.
   - use `<project-root>/.backup/`
   - if the target file already exists, save a timestamped copy of the prior version there
   - if `.backup/` does not exist, create it
6. Work on a copy unless the user explicitly wants in-place edits.
7. Apply only the approved items and keep the scope tight.
8. Update fields, tables of contents, and page numbers after structural changes.
9. Append a short memory entry to `<project-root>/.backup/memories.md`.
   - create the file if it does not exist
   - record what changed, which file was edited, and where the backup was stored
10. Verify output:
   - re-run inspection for structural changes
   - open in native Office or export PDF when layout matters
   - spot-check headings, sheets, formulas, tables, slide order, notes, and section breaks after edits

## Editing Rules

- For edit requests, always send a pre-edit confirmation list first.
- If `profile_docx.py` suggests the file is academic, ask whether the user wants academic-formatting defaults before proposing detailed edit items.
- Use a numbered list when changes are distinct and need approval item by item.
- Make it easy for the user to answer with:
  - approve one item
  - approve multiple item numbers
  - approve all
  - reject specific items
- Before writing to an existing file, always create a timestamped backup inside `.backup/`.
- After each completed edit or overwrite, always append a concise entry to `.backup/memories.md`.
- Memory entries should include:
  - timestamp
  - target file path
  - backup file path if one was created
  - approved change items
  - brief result or warning
- Keep original files untouched until the user confirms destructive edits.
- Prefer style-based edits over run-by-run formatting hacks.
- Use heading styles instead of fake bold paragraphs when a document needs a navigable structure.
- Update tables of contents and fields before final delivery.
- Keep numbering, page setup, and section boundaries consistent after inserts or deletions.
- Keep sheet names, formulas, hidden sheets, and named ranges consistent after Excel edits.
- Keep slide order, layout placeholders, notes, and title consistency intact after PowerPoint edits.
- Use tracked revisions or comments when the user wants reviewable edits instead of silent replacement.
- Export PDF for final review when the user cares about exact page layout.

## Analysis Rules

- When the user sends a Word document for analysis, do not stop at structure counts.
- Check:
  - formatting state
  - likely document kind
  - likely topic
  - topic-specific presentation advice
- Use `scripts/profile_docx.py` for the first pass.
- If confidence is low, say so explicitly instead of overstating the classification.

## Agent Collaboration

- This skill internalizes selected patterns from the `agency-agents` repo.
- Default role announcement:
  - `Codex Office Documents Specialist`
- Do not imply that non-Office local skills are part of the core Office engine.
- If companion skills are being used, announce them explicitly before proceeding and frame them as optional supporting lenses.
- Read `references/agent-collaboration.md` when:
  - the user asks for multiple agents
  - the task needs a second lens such as document structure, visual polish, or skeptical final QA
- By default, stay in `Codex Office Documents Specialist` mode only.
- Only pull in companion skills when:
  - the user explicitly asks for multiple agents, or
  - the task clearly needs a non-Office lens such as content strategy, visual critique, or final QA

## Engine Selection

### Word

- Use `scripts/inspect_docx.py` first for unknown `.docx` or `.docm` files.
- Use `scripts/word_com_actions.ps1` for:
  - pagination-sensitive edits
  - headers, footers, page numbers, sections, TOC, tracked revisions, comments, or PDF export
  - fixes that must match what Word itself would render
- Use `scripts/markdown_to_docx.py` for new reports, proposals, briefs, and essays created from Markdown.
- Use `scripts/fill_docx_template.py` for reusable templates with `{{placeholders}}`.
- Avoid direct OOXML rewrites when tracked changes, complex numbering, comments, or fragile layout matter.

### Excel

- Use `scripts/inspect_xlsx.py` first for unknown `.xlsx` or `.xlsm` files.
- Use `scripts/markdown_to_xlsx.py` when the workbook is mostly a report made from headings and tables.
- Escalate to native Excel COM when the job needs:
  - macros
  - pivot tables
  - chart formatting that must match Excel
  - recalculation behavior that openpyxl cannot guarantee
  - print setup or PDF export that depends on Excel rendering
- Preserve formulas and sheet-level structure unless the user explicitly approves a rebuild.

### PowerPoint

- Use `scripts/inspect_pptx.py` first for unknown `.pptx` or `.pptm` files.
- Use `scripts/structured_pptx.py` when the user has a clear slide model and wants an editable deck.
- Prefer a template `.pptx` when brand theme, master slides, or placeholders matter.
- Escalate to native PowerPoint COM when the job needs:
  - exact theme fidelity
  - speaker notes polishing in the real deck
  - animation, transition, or media timing edits
  - PDF export that must match PowerPoint rendering
- Use an external service such as 2slides only when the user explicitly allows API-based generation.

### Email And Outlook-Style Drafts

- Draft `.eml` or Office-style message content locally when the user wants a saved email artifact.
- Do not send mail, calendar invites, or Outlook actions without explicit confirmation.
- When a request crosses into Outlook automation, prefer a short native COM job over inventing HTML structure by hand.

### Cross-Platform And Large Surface Area Tasks

- Use the Aspose repo review as a design reference when the user wants:
  - one workflow across Word, Excel, PowerPoint, PDF, OCR, or email
  - session-based editing
  - structured tool contracts and output schemas
  - cross-platform handling instead of Windows-only COM
- Do not assume Aspose is available locally unless the environment proves it.

## Human-Like Layout Checklist

- Normalize heading hierarchy and avoid skipped heading levels unless the source already requires it.
- Check paragraph spacing before and after headings instead of stacking blank lines.
- Keep tables readable: consistent header row, sane column widths, and no accidental empty rows.
- Keep headers, footers, and page numbering aligned with section boundaries.
- Check for orphaned headings near page breaks after large edits.
- Re-run `update_fields` and `update_toc` before final export.
- For Excel, check frozen panes, filter ranges, column widths, number formats, and print area before delivery.
- For PowerPoint, check title consistency, slide order, placeholder overflow, image cropping, and notes presence before delivery.

## Academic Document Rules

- For essays, reports, theses, or academic papers, recommend starting major sections on a new page.
- If the user says the file is a thesis, dissertation, graduation report, or academic report, read `references/academic-formatting.md` and apply those defaults unless the user provides another institutional guide.
- Major sections usually include:
  - `Introduction`
  - `Main Content`
  - `Conclusion`
  - `References`
  - `Appendix`
- Use real page breaks, not stacked blank paragraphs, when moving major headings to a new page.
- Do not force page breaks before lower-level headings such as `1.`, `1.1`, or `1.1.1` unless the user explicitly asks for that structure.

## Resources

### `scripts/inspect_docx.py`
Run safe structure inspection on `.docx` or `.docm` files without needing extra Python packages. Use it to capture metadata, paragraph and table samples, styles used, headers, footers, comments, and media before editing.

Example:

```bash
python scripts/inspect_docx.py "C:\path\to\document.docx"
```

### `scripts/profile_docx.py`
Profile a Word document by likely document kind, topic, and formatting advice. Use this before editing or reviewing a user-provided `.docx` when you need to decide whether to offer academic formatting or topic-specific advice.

Example:

```bash
python scripts/profile_docx.py "C:\path\to\document.docx"
```

### `scripts/inspect_xlsx.py`
Inspect `.xlsx` or `.xlsm` files and emit workbook-level structure, sheet summaries, table names, chart counts, formula counts, comments, hyperlinks, merged ranges, and sample rows.

Example:

```bash
python scripts/inspect_xlsx.py "C:\path\to\workbook.xlsx"
```

### `scripts/inspect_pptx.py`
Inspect `.pptx` or `.pptm` files and emit slide summaries, title text, shape counts, tables, charts, images, notes, and media counts.

Example:

```bash
python scripts/inspect_pptx.py "C:\path\to\deck.pptx"
```

### `scripts/word_com_actions.ps1`
Run Microsoft Word automation on Windows against an existing document or a new blank document. Use it for replace operations, headings, tables, comments, margins, orientation, page numbers, table of contents maintenance, tracked revisions, and PDF export.

Example:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\word_com_actions.ps1 `
  -ActionsPath .\job.json
```

### `references/task-matrix.md`
Read this file when deciding how to map an Office request to inspection, generation, native COM, or external services, and when building a final QA pass for layout-sensitive deliverables.

### `scripts/markdown_to_docx.py`
Create a new `.docx` from Markdown using python-docx. It supports headings, ordered and unordered lists, tables, quotes, alignment tags, page breaks, horizontal rules, image URLs, metadata, headers and footers, and an update-on-open TOC field.

Example:

```bash
python scripts/markdown_to_docx.py report.md --output report.docx --toc --title "Quarterly Report"
```

### `scripts/fill_docx_template.py`
Fill an existing `.docx` template using a JSON object of placeholder values. It supports `{{placeholder}}` in body paragraphs, tables, headers, footers, first-page headers and footers, and even-page headers and footers.

Example:

```bash
python scripts/fill_docx_template.py template.docx --data values.json --output filled.docx
```

### `references/word-actions.md`
Read this file before creating a JSON job for `word_com_actions.ps1`. It defines the supported action schema and provides copyable examples.

### `references/docx-generation.md`
Read this file before creating Markdown-generation or template-fill jobs. It describes the supported Markdown syntax, placeholder behavior, and the safest handoff into Word COM for final polish.

### `references/academic-formatting.md`
Read this file for thesis, dissertation, graduation-report, or academic-report formatting defaults. It captures page setup, typography, numbering, heading levels, table and figure rules, and reference-list expectations from the local guide PDF in this project.

### `references/document-profiles.md`
Read this file when profiling a Word document and deciding which advice to give based on format and topic.

### `scripts/markdown_to_xlsx.py`
Create a workbook from Markdown headings and pipe tables. Use `## Sheet: Name` to start a new sheet. This is the fastest path for workbook-style reports that are mostly tables.

Example:

```bash
python scripts/markdown_to_xlsx.py report.md --output report.xlsx --summary
```

### `scripts/structured_pptx.py`
Create a PowerPoint deck from structured JSON. It supports title, section, content, table, image, quote, and two-column slides, and can optionally start from a `.pptx` template.

Example:

```bash
python scripts/structured_pptx.py slides.json --output review-deck.pptx --summary
```

### `scripts/project_memory.py`
Create timestamped backups in `.backup/` and append memories in a consistent format. Use it when a project requires deterministic backup and edit logging across Word, Excel, and PowerPoint workflows.

Example:

```bash
python scripts/project_memory.py backup --project-root /path/to/project --target /path/to/project/report.docx
python scripts/project_memory.py remember --project-root /path/to/project --target /path/to/project/report.docx --backup /path/to/project/.backup/20260319-104500-report.docx --approved "items 1, 2" --changes "updated headings and spacing" --result "saved successfully"
```

### `references/office-engine-matrix.md`
Read this file when choosing between Word COM, Excel COM, PowerPoint COM, Python generation, template fill, or external Office services.

### `references/excel-generation.md`
Read this file before generating or reshaping `.xlsx` files from Markdown and before deciding when Excel COM is required.

### `references/powerpoint-generation.md`
Read this file before generating or editing `.pptx` decks, especially when choosing between structured JSON, templates, COM, or external slide APIs.

### `references/repo-review.md`
Read this file when you need the distilled learnings from OfficeMCP, mcp-ms-office-documents, Office-Word-MCP-Server, 2slides, Aspose MCP, and cs-office-mcp-server.

### `references/agent-collaboration.md`
Read this file when the user asks for additional agents or when the Office workflow would benefit from structure, design, research, or QA companion skills inspired by the `agency-agents` repo.

## Output Expectations

- Before editing, provide a clean proposed-change list and wait for approval.
- After editing, report the backup path and confirm that `.backup/memories.md` was updated.
- Report the chosen engine and why.
- State whether edits were made on a copy or in place.
- Summarize every structural change: headings, tables, margins, sections, headers, footers, comments, tracked changes, or exports.
- Call out anything not verified, especially page layout, fonts, broken references, or localized style-name issues.
