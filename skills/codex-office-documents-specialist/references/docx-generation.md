# DOCX Generation Workflows

## Choose The Right Path

- Use `scripts/markdown_to_docx.py` when the user gives notes, Markdown, or structured content and wants a new `.docx`.
- Use `scripts/fill_docx_template.py` when the user already has a branded `.docx` template and only needs variables filled in.
- Use `scripts/word_com_actions.ps1` after generation when exact final layout, PDF export, or Word-native updates matter.

## Approval Workflow

Before editing an existing document, provide a proposed-change list first.

- Put each proposed change on its own numbered line.
- Describe the concrete effect, not vague intent.
- Wait for the user to approve:
  - one item
  - multiple item numbers
  - all items
- Only then apply the approved changes.

## Backup And Memory Rules

Before any edit that changes an existing file, create a backup copy first.

- Store backups in `<project-root>/.backup/`.
- Use a timestamp in the backup filename so multiple revisions do not collide.
- If the output path already exists during generation, back up that prior file before overwriting it.
- If the output file is brand new, no file backup is required, but still update the memory log.

After the edit or overwrite finishes, append a short entry to `<project-root>/.backup/memories.md`.

Suggested entry shape:

```markdown
## 2026-03-19 10:30
- target: path/to/file.docx
- backup: .backup/20260319-103000-file.docx
- approved: items 1, 2, 4
- changes: normalized Heading 4, removed blank paragraphs, updated TOC
- notes: page layout not manually reviewed
```

Example:

```text
1. Convert tertiary headings such as 1.1.1 and 1.1.2 to Heading 4.
2. Remove extra blank paragraphs in the document body.
3. Start Introduction, Main Content, and Conclusion on new pages with real page breaks.
```

## Markdown To DOCX

Supported block syntax:

- Headings: `#` through `######`
- Ordered lists: `1. item`
- Bullet lists: `- item`, `* item`, `+ item`
- Tables:

```markdown
| Name | Value |
|------|-------|
| A    | B     |
```

- Page break: `---`
- Horizontal line: `***`
- Quote: `> text`
- Images: `![alt](https://...)`
- Alignment:
  - `<center>text</center>`
  - `<div align="right">text</div>`

Supported inline formatting:

- `**bold**`
- `*italic*`
- `***bold italic***`
- `~~strikethrough~~`
- `__underline__`
- `` `code` ``
- `[link](https://...)`

Example:

```bash
python scripts/markdown_to_docx.py notes.md --output report.docx --toc --header "Internal Draft" --footer "Page {page} of {pages}"
```

## Template Filling

Supported placeholder targets:

- body paragraphs
- tables in the body
- headers and footers
- first-page headers and footers
- even-page headers and footers

Placeholder values can contain inline Markdown. If a body placeholder value contains block Markdown such as lists, headings, or tables, the filler inserts new paragraphs after the placeholder paragraph.

Template style guidance:

- For the best output, keep built-in styles for headings and lists in the template.
- The filler looks for built-in style IDs such as `Heading1`, `Heading2`, `ListBullet`, and `ListNumber`.
- If a template does not contain the matching built-in style, generation still succeeds but those inserted blocks can fall back to normal paragraphs.

Example JSON:

```json
{
  "recipient_name": "Nguyen Van A",
  "subject": "Project update",
  "body": "## Summary\n\n- Item 1\n- Item 2\n\nPlease review **today**."
}
```

Example:

```bash
python scripts/fill_docx_template.py formal-letter.docx --data values.json --output filled-letter.docx
```

## Final Polish

After generation or template fill:

1. inspect the output with `scripts/inspect_docx.py`
2. if layout matters, run a short Word COM job to update fields or TOC
3. export PDF from Word COM for review or delivery

This hybrid flow is usually better than forcing python-docx to solve every layout problem.
