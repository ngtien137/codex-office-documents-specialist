# Backup And Memory

## Required Behavior

- Before editing an existing file, create a backup in `.backup/`.
- Before overwriting an existing output file, create a backup in `.backup/`.
- After each completed edit, append a short entry to `.backup/memories.md`.
- Use `scripts/project_memory.py` when you want deterministic backup or memory logging behavior across Word, Excel, and PowerPoint tasks.

## Location Rules

- Backup directory: `<project-root>/.backup/`
- Memory log: `<project-root>/.backup/memories.md`

Create `.backup/` and `memories.md` if they do not exist.

## Backup Naming

Use a timestamped filename:

```text
YYYYMMDD-HHMMSS-original-file-name.ext
```

Examples:

```text
.backup/20260319-103000-report.docx
.backup/20260319-154500-budget.xlsx
.backup/20260319-161500-review-deck.pptx
```

## Memory Entry Template

```markdown
## 2026-03-19 10:30
- target: path/to/file.docx
- backup: .backup/20260319-103000-file.docx
- approved: items 1, 2, 3
- changes: brief summary of what was actually changed
- result: saved successfully
- warnings: anything not verified
```

## Practical Notes

- Log what was actually executed, not every proposed idea.
- If no backup file was needed because the target did not exist yet, write `backup: none`.
- Keep entries short and chronological.
- When multiple files are changed in one task, either:
  - use one entry with multiple target lines, or
  - create one entry per file when that is clearer.
