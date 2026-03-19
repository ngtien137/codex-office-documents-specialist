# Excel Generation

## When To Use Bundled Excel Generation

Use `scripts/markdown_to_xlsx.py` when:

- the workbook is mainly a report with headings and tables
- the user wants an editable `.xlsx`
- formulas are simple cell formulas that can live directly in sheet cells
- the job does not depend on macros, pivots, slicers, or workbook events

Use `scripts/inspect_xlsx.py` first when the workbook already exists and you need to understand:

- sheet names and sheet state
- used range size
- table names
- chart counts
- comment and hyperlink counts
- formula density

## Markdown Shape

Supported structure:

- `## Sheet: Revenue` starts a new sheet
- `# Heading` or `### Heading` writes a styled heading row
- pipe tables become worksheet tables
- cell text beginning with `=` is written as an Excel formula

Example:

```markdown
## Sheet: Revenue

# Q1 Summary

| Region | Revenue | Growth |
|---|---:|---:|
| North | 120000 | =C2 |
| South | 95000 | =C3 |
```

Example command:

```bash
python scripts/markdown_to_xlsx.py report.md --output report.xlsx --summary
```

## Escalate To Native Excel

Do not force openpyxl to solve these jobs:

- macros or VBA
- pivot tables
- chart fidelity that must match Excel
- workbook events or external connections
- page setup and PDF layout that must match Excel rendering

For those, use native Excel COM patterns inspired by OfficeMCP or cs-office-mcp-server.

## Editing Rules

- preserve formulas unless the user explicitly approves recalculation changes
- preserve hidden sheets and named ranges unless the user explicitly approves cleanup
- if replacing values broadly, inspect first and keep the replacement scope narrow
