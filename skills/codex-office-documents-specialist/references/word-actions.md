# Word COM Actions

Use `scripts/word_com_actions.ps1` with a JSON file that contains top-level document options and an `actions` array.

## Top-Level Schema

```json
{
  "source": "optional-existing-document.docx",
  "saveAs": "required-for-new-document.docx",
  "exportPdf": "optional-output.pdf",
  "visible": false,
  "actions": []
}
```

Rules:

- Omit `source` to create a new blank document.
- Set `saveAs` when creating a new document or when you want a copy.
- Set `exportPdf` when the user needs a reviewable fixed-layout output.
- Use `visible: true` only when interactive Word UI is helpful during debugging.

## Supported Actions

### `replace_text`

```json
{
  "type": "replace_text",
  "find": "old text",
  "replace": "new text",
  "matchCase": false,
  "wholeWord": false
}
```

### `append_heading`

```json
{
  "type": "append_heading",
  "text": "1. Scope",
  "level": 1
}
```

Use heading levels `1` through `9`.

### `append_paragraph`

```json
{
  "type": "append_paragraph",
  "text": "This report summarizes the current document state.",
  "style": "Normal"
}
```

### `append_page_break`

```json
{
  "type": "append_page_break"
}
```

### `insert_table`

```json
{
  "type": "insert_table",
  "rows": [
    ["Field", "Value"],
    ["Project", "Word agent"],
    ["Status", "Ready"]
  ],
  "headerRow": true,
  "style": "Table Grid",
  "autoFit": "content"
}
```

`autoFit` supports `fixed`, `content`, or `window`.

### `add_comment`

```json
{
  "type": "add_comment",
  "find": "legacy wording",
  "comment": "Rewrite this sentence in a more formal tone.",
  "matchCase": false,
  "wholeWord": false
}
```

### `set_margins_cm`

```json
{
  "type": "set_margins_cm",
  "top": 2.5,
  "bottom": 2.5,
  "left": 3.0,
  "right": 2.0
}
```

### `set_orientation`

```json
{
  "type": "set_orientation",
  "orientation": "landscape"
}
```

Use `portrait` or `landscape`.

### `set_header_text`

```json
{
  "type": "set_header_text",
  "text": "Internal Draft"
}
```

### `set_footer_text`

```json
{
  "type": "set_footer_text",
  "text": "For review only"
}
```

### `add_page_numbers`

```json
{
  "type": "add_page_numbers",
  "alignment": "right",
  "firstPage": true
}
```

Alignment supports `left`, `center`, or `right`.

### `create_toc`

```json
{
  "type": "create_toc",
  "upperHeadingLevel": 1,
  "lowerHeadingLevel": 3
}
```

### `update_toc`

```json
{
  "type": "update_toc"
}
```

### `update_fields`

```json
{
  "type": "update_fields"
}
```

### `set_track_revisions`

```json
{
  "type": "set_track_revisions",
  "enabled": true
}
```

### `accept_all_revisions`

```json
{
  "type": "accept_all_revisions"
}
```

### `reject_all_revisions`

```json
{
  "type": "reject_all_revisions"
}
```

## Example Job

```json
{
  "saveAs": "output/sample-agent-doc.docx",
  "exportPdf": "output/sample-agent-doc.pdf",
  "actions": [
    { "type": "append_heading", "text": "Word Agent Demo", "level": 1 },
    {
      "type": "append_paragraph",
      "text": "This file was generated through Microsoft Word COM automation.",
      "style": "Normal"
    },
    {
      "type": "insert_table",
      "rows": [
        ["Capability", "Status"],
        ["Layout editing", "Ready"],
        ["PDF export", "Ready"]
      ],
      "headerRow": true,
      "style": "Table Grid",
      "autoFit": "content"
    },
    { "type": "add_page_numbers", "alignment": "right", "firstPage": true },
    { "type": "update_fields" }
  ]
}
```
