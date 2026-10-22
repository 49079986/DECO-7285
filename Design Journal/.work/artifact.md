# Design Journal Template Contract

- Reference: `C:\Users\yuyan\Desktop\S4\DECO 7285\Design Journal\Design_Journal_Format.docx`
- SHA-256: `7ec962d82172be98ec0d223367a6e69316611fdb36ed51d8fbf0860a8caaed4c`
- Reference pages: 1
- Sections: 1
- Evidence: `.qa_template/style.json`; section audit recorded in the task log. The packaged renderer could not run because its bundled runtime does not include LibreOffice, so Word 16 PDF export is used for visual review.

## Page system

- One A4 portrait section, 8.27 by 11.69 inches.
- Margins: left 0.59 in, right 0.79 in, top 0.96 in, bottom 0.69 in.
- No recurring header; footer is empty.

## Typography and components

- Normal paragraphs use the document Normal style, visually Calibri 11 pt, black, left aligned.
- The top metadata line contains Student Name, Student Number, and Team Name separated by tabs.
- A second metadata line contains the week number.
- The body uses three bold inline labels in order: Title, Description, Visual Evidence, and Statement on Learning.
- There are no tables, numbered lists, headers, footers, or content controls.
- The title is an inline label rather than Word's Title style; preserve this source-derived treatment because it is the defining template pattern.

## Editable slots

- Paragraph 0: replace metadata values while preserving the same one-line role.
- Paragraph 1: replace week value.
- Paragraph 2: append the journal title after the existing label.
- Paragraphs after Description: replace blank space with reflective description prose.
- Paragraphs after Visual Evidence: insert one relevant supplied image with a short caption.
- Paragraphs after Statement on Learning: insert first-person reflection and next-step prose.

## Package preservation and fidelity gates

- Start each output from a copy of the reference.
- Preserve the A4 page geometry, black text, label order, and restrained academic appearance.
- Keep each journal to one page where readable; do not reduce body text below 10 pt.
- Do not invent a student number. Use `Not provided` until the user supplies one.
- Use only supplied project visuals and facts found in the workspace.
- Verify every output through Word PDF export and inspect every rendered page for clipping, overlap, missing images, or unexpected pagination.

