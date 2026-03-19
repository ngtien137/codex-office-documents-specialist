# Task Matrix

## Engine Choice

- Use `scripts/inspect_docx.py` first when the user gives an unknown `.docx` or `.docm` file and asks for edits, cleanup, or analysis.
- Use `scripts/profile_docx.py` after `inspect_docx.py` when the user gives a Word document for analysis and you need to identify document type, topic, and tailored advice.
- Use `scripts/inspect_xlsx.py` first when the user gives an unknown `.xlsx` or `.xlsm` file and asks for edits, cleanup, or analysis.
- Use `scripts/inspect_pptx.py` first when the user gives an unknown `.pptx` or `.pptm` file and asks for edits, cleanup, or analysis.
- Use `scripts/markdown_to_docx.py` when the user starts from Markdown, outline notes, meeting notes, or report content and wants a fresh `.docx`.
- Use `scripts/fill_docx_template.py` when the user has a reusable `.docx` template with placeholders and wants a branded output quickly.
- Use `scripts/markdown_to_xlsx.py` when the workbook is mostly made of headings and tables and the user wants an editable `.xlsx`.
- Use `scripts/structured_pptx.py` when the deck is well described in structured content and should remain editable after generation.
- Use `scripts/word_com_actions.ps1` when the user wants Word-like results, especially for layout, sections, headings, tables, comments, tracked changes, page numbering, headers, footers, or PDF output.
- Use native Excel COM patterns when the job needs macros, pivots, chart fidelity, or print layout.
- Use native PowerPoint COM patterns when the job needs exact themes, notes editing, animation, transitions, or export matching PowerPoint itself.
- Use remote slide services only when the user explicitly allows third-party APIs and design automation is more important than offline reproducibility.
- Stay out of raw OOXML edits for final-polish jobs unless Word is unavailable and the user accepts lower fidelity.

## Request Mapping

- "Chinh bo cuc", "dan trang", "can lai heading", "fix header footer":
  Use Word COM.
- "Tu markdown tao file Word", "convert note thanh .docx", "tao report tu outline":
  Use markdown-to-docx, then Word COM only if final polish is needed.
- "Dien vao template Word", "fill contract template", "thay placeholder trong mau":
  Use DOCX template filler, then Word COM for final pagination or field updates.
- "Tao workbook tu bang markdown", "xuat bang so lieu thanh Excel", "tao file xlsx de sua tiep":
  Use markdown-to-xlsx first.
- "Kiem tra workbook nay co bao nhieu sheet, table, cong thuc":
  Use inspect-xlsx first.
- "Can macro, pivot, chart, page setup Excel":
  Use native Excel COM.
- "Tao slide tu outline", "tao deck tu JSON", "tao file pptx editable":
  Use structured-pptx first.
- "Kiem tra slide, notes, picture, table, chart trong deck":
  Use inspect-pptx first.
- "Lam slide dep theo theme co san", "tao slide bang AI design service":
  Only use remote slide APIs if the user allows it.
- "Them muc luc", "update page number", "xuat PDF":
  Use Word COM.
- "Tim hieu file nay co gi", "kiem tra style dang dung", "dem bang va section":
  Use document inspection first.
- "Phan tich file Word nay", "xem dung dang cua no", "xem chu de cua no", "cho loi khuyen chinh sua":
  Use inspect-docx plus profile-docx, then give both formatting advice and topic-specific advice.
- "Lam luan van", "lam luan an", "lam bao cao tot nghiep", "trinh bay theo huong dan khoa":
  Read `references/academic-formatting.md` first, then inspect the file, then use Word COM for layout-sensitive cleanup.
- "Thay hang loat mot cum tu trong ca file":
  Inspect first, then use Word COM replace.
- "Them nhan xet nhu review", "giu track changes":
  Use Word COM with comments or tracked revisions enabled.

## High-Fidelity Editing Checklist

- Before editing, send a numbered proposed-change list and wait for approval.
- Let the user approve one item, several item numbers, or all items.
- Before editing an existing file or overwriting an output path, create a timestamped backup in `.backup/`.
- After the edit, append a note to `.backup/memories.md`.
- Preserve the source file by default; create a copy before risky edits.
- Rebuild structure with Word styles, not manual formatting.
- Preserve sheet order, formulas, named ranges, and hidden tabs unless the user explicitly approves a workbook rebuild.
- Preserve slide order, master usage, notes, and placeholders unless the user explicitly approves a deck rebuild.
- Update fields after changing headings, captions, or cross-references.
- Re-check section breaks after inserting tables or page breaks.
- Export PDF for sign-off when the user cares about exact layout.

## Academic Layout Default

- For academic documents, recommend page breaks before the largest structural sections.
- Default candidates:
  - introduction
  - main content/body
  - conclusion
  - references
  - appendix
- Do not insert page breaks before subsection levels unless the user explicitly requests a chapter-like layout.

## Common Failure Modes

- Localized Word installations may expose built-in style names differently.
  Prefer `append_heading` for built-in headings because it uses numeric built-in styles.
- Replace operations can affect headers, footers, or tables if the target text is broad.
  Inspect first and keep find strings specific.
- Section-level margins and orientation can differ inside the same document.
  Verify sections after changing page setup.
- Tables of contents only update correctly if real heading styles are applied.
- openpyxl does not execute macros or fully emulate Excel recalculation.
  If the workbook depends on macros, pivots, or workbook events, escalate to Excel COM.
- python-pptx can generate editable decks but does not cover PowerPoint animation or theme behavior completely.
  If visual fidelity is the hard requirement, escalate to PowerPoint COM or use a trusted template.
- Topic classification is heuristic.
  If profile confidence is low, present it as a best-effort guess and ask the user to confirm the document type before applying strong formatting assumptions.
