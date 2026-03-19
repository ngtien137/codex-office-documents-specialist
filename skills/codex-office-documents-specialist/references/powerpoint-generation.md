# PowerPoint Generation

## Choose The Right Path

Use `scripts/structured_pptx.py` when:

- the user can describe the deck in structured content
- the output should remain editable in PowerPoint
- the design can start from a normal template or a simple built-in layout

Use a template `.pptx` when:

- brand theme matters
- slide masters and placeholder positions already exist
- the user wants consistency with prior decks

Use native PowerPoint COM when:

- the real theme rendering matters
- notes, media timing, transitions, or animations must be edited
- the final PDF must match PowerPoint exactly

Use 2slides or another external slide API only when:

- the user explicitly allows a third-party service
- automated design polish is more important than reproducibility
- sync or async job polling is acceptable

## Supported Structured Slides

`structured_pptx.py` supports these slide types:

- `title`
- `section`
- `content`
- `table`
- `image`
- `quote`
- `two_column`

Example JSON:

```json
{
  "format": "16:9",
  "slides": [
    { "type": "title", "title": "Education Trends", "subtitle": "2026 review" },
    {
      "type": "content",
      "title": "Key Drivers",
      "bullets": [
        "Massification",
        { "text": "Digital delivery", "level": 1 },
        "Internationalization"
      ]
    }
  ]
}
```

Example command:

```bash
python scripts/structured_pptx.py slides.json --output review-deck.pptx --summary
```

## Inspection

Use `scripts/inspect_pptx.py` before editing an existing deck to understand:

- slide count
- title text
- pictures, tables, charts, and text boxes
- speaker notes presence
- media counts

## Common Risks

- python-pptx can create editable decks but does not fully emulate master themes, animation, or transition behavior
- image cropping can look different from user expectations if the template is unknown
- title placeholders differ across templates, so always inspect or test one slide before generating a large deck
