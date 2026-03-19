# codex-office-documents-specialist

High-fidelity Office document skills for Codex across Word, Excel, and PowerPoint.

This repository packages a single production-ready Codex skill focused on:

- inspecting Office files safely before editing
- choosing the right engine for the job
- editing Word documents with native-fidelity workflows on Windows
- generating Office deliverables from Markdown, JSON, and templates
- applying a disciplined approval, backup, and memory-log workflow

## Why This Skill Stands Out

Most document automation stacks are strong at generation or strong at native editing. This skill is designed to handle both.

| Area | What this skill does well |
| --- | --- |
| Word fidelity | Uses Microsoft Word COM for layout-sensitive edits, TOC updates, page numbering, comments, revisions, and PDF export |
| Safe inspection | Inspects `.docx`, `.xlsx`, and `.pptx` structure before editing so the workflow starts from facts, not guesses |
| Generation speed | Creates Word, Excel, and PowerPoint files from Markdown or structured JSON when deterministic output is faster than manual editing |
| Template workflows | Fills DOCX templates with placeholders across body, tables, headers, and footers |
| Academic formatting | Detects likely thesis/report documents and offers formatting updates before any edit is made |
| Controlled edits | Requires a numbered pre-edit approval list, creates backups, and records changes in project memory |

## Included Skill

| Skill | Purpose |
| --- | --- |
| `skills/codex-office-documents-specialist` | Main Office-wide skill for Word, Excel, PowerPoint, formatting review, engine selection, backup flow, and academic document handling |

## Capability Matrix

| File Type | Inspect | Generate | Edit | Native Fidelity |
| --- | --- | --- | --- | --- |
| `.docx` / `.docm` | Yes | Yes | Yes | Strongest via Word COM on Windows |
| `.xlsx` / `.xlsm` | Yes | Yes | Partial, generation-first | Escalate to Excel COM when macros, pivots, or print layout matter |
| `.pptx` / `.pptm` | Yes | Yes | Structured generation-first | Escalate to PowerPoint COM when theme fidelity, notes, or export quality matter |
| `.eml` / Outlook-style drafts | Limited | Draft content | Limited | Outlook automation is intentionally explicit-only |

## Core Workflows

### 1. Analyze First

The skill inspects the Office package before editing:

- Word: styles, headings, sections, TOC, header/footer, comments, media
- Excel: sheets, tables, charts, formulas, merges, hyperlinks
- PowerPoint: slide titles, shape counts, tables, charts, pictures, notes

### 2. Choose the Right Engine

The skill treats Office work as an engine-selection problem:

- use OOXML inspection when the first task is understanding structure
- use `python-docx`, `openpyxl`, or `python-pptx` for deterministic generation
- use Word COM when layout must match what a human would get in Word
- escalate to native Excel or PowerPoint automation when rendering fidelity matters more than portability

### 3. Require Approval Before Editing

Before any in-place change, the skill:

1. proposes a numbered change list
2. waits for approval
3. creates a backup
4. applies only approved changes
5. updates fields and verification artifacts
6. writes an entry to project memory

### 4. Handle Academic Documents Properly

When a Word file looks like a thesis, dissertation, graduation report, or academic report, the skill:

- profiles the document kind and topic
- asks whether academic formatting defaults should be applied
- proposes the exact updates before editing
- avoids silent formatting changes

## Repository Layout

```text
skills/
  codex-office-documents-specialist/
    agents/
    references/
    scripts/
    SKILL.md
```

## Bundled Scripts

| Script | Purpose |
| --- | --- |
| `inspect_docx.py` | Inspect DOCX or DOCM structure safely |
| `profile_docx.py` | Infer likely document kind, topic, and advice for Word files |
| `inspect_xlsx.py` | Inspect workbook structure, sheets, tables, charts, and formulas |
| `inspect_pptx.py` | Inspect deck structure, slides, tables, charts, media, and notes |
| `markdown_to_docx.py` | Generate DOCX from Markdown |
| `fill_docx_template.py` | Fill DOCX templates with placeholders |
| `markdown_to_xlsx.py` | Generate XLSX workbooks from Markdown tables |
| `structured_pptx.py` | Generate PPTX decks from structured JSON |
| `word_com_actions.ps1` | Run native Microsoft Word automation jobs on Windows |
| `project_memory.py` | Create backups and append memory entries consistently |

## Installation

### Prerequisites

| Requirement | Notes |
| --- | --- |
| Codex skills directory | You need access to your local `$CODEX_HOME/skills/` folder |
| Python 3.9+ | Required for the bundled Python scripts |
| Windows + Microsoft Word | Needed for Word COM workflows and highest-fidelity Word edits |
| Optional Python packages | The scripts may rely on packages such as `python-docx`, `openpyxl`, `python-pptx`, `lxml`, and `Pillow` depending on the workflow |

### Install the Skill

Copy the skill folder into your local Codex skills directory:

```text
skills/codex-office-documents-specialist
```

After installation, the skill should exist at:

```text
$CODEX_HOME/skills/codex-office-documents-specialist
```

## How To Use

### Trigger the Skill

Ask Codex to use:

```text
$codex-office-documents-specialist
```

Typical requests:

- "Use $codex-office-documents-specialist to inspect this DOCX and tell me what formatting is being used."
- "Use $codex-office-documents-specialist to convert this Markdown report into a Word document."
- "Use $codex-office-documents-specialist to fill this DOCX template and then export a polished PDF."
- "Use $codex-office-documents-specialist to review whether this report matches thesis formatting conventions."

### Example Commands

Inspect a Word document:

```bash
python scripts/inspect_docx.py "/path/to/document.docx"
```

Profile an academic-looking Word document:

```bash
python scripts/profile_docx.py "/path/to/document.docx"
```

Generate a DOCX from Markdown:

```bash
python scripts/markdown_to_docx.py report.md --output report.docx --toc --title "Quarterly Report"
```

Fill a DOCX template:

```bash
python scripts/fill_docx_template.py template.docx --data values.json --output filled.docx
```

Generate an XLSX workbook:

```bash
python scripts/markdown_to_xlsx.py report.md --output report.xlsx --summary
```

Generate a PPTX deck:

```bash
python scripts/structured_pptx.py slides.json --output review-deck.pptx --summary
```

## Editing Policy

This skill intentionally favors controlled document work over silent automation.

| Rule | Behavior |
| --- | --- |
| Pre-edit approval | Always propose a numbered change list before editing |
| Backup first | Create a timestamped copy in `.backup/` before overwriting |
| Memory logging | Append what changed to `.backup/memories.md` |
| Scope control | Apply only approved changes |
| Verification | Re-inspect, update fields, and verify structure after edits |

## Best Fit

Use this skill when you need:

- Word edits that must preserve real Office layout
- a hybrid workflow of inspect -> propose -> edit -> verify
- document generation from structured source material
- thesis/report formatting review
- a more disciplined alternative to ad hoc Office scripting

## Current Limits

| Area | Current Limit |
| --- | --- |
| Excel native automation | Not yet bundled as a dedicated COM runner |
| PowerPoint native automation | Not yet bundled as a dedicated COM runner |
| Cross-platform native fidelity | Strongest results still depend on Windows Office for Word layout-sensitive work |
| Public repo cleanliness | Ignore rules are included, but local cache artifacts should still be removed before commit if present |

## Privacy Note

This repository is prepared to avoid publishing local absolute paths in the skill docs. Before pushing, still verify that transient local artifacts such as `__pycache__/` and `.pyc` files are not committed.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
