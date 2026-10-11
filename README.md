# PDF / PPTX to Word

Local web app that turns PDF and PowerPoint files into editable Word documents.
Nothing leaves your machine: a small Python server converts the files and the browser page is only the front end.

## Run

Requires Python 3.12+ and [uv](https://docs.astral.sh/uv/).

```powershell
.\run.ps1
```

This installs dependencies, starts the server on http://127.0.0.1:8765 and opens the page.
Drop files on the page, wait for the "Ready" badge, click **Save .docx**.
Converted files are kept in the temp folder for one hour, then deleted.

## What it does

| Input | Engine | Kept |
|-------|--------|------|
| PDF | pdf2docx (PyMuPDF) | paragraphs, fonts, sizes, colours, bold/italic, tables, images, multi-column text, page size and margins |
| PPTX | python-pptx + python-docx | one slide per page at the slide's size; fonts, sizes, colours and alignment resolved through layout, master and theme; bullets and numbering; shapes placed side by side become columns; filled boxes keep their background; tables with merged cells; pictures with crop and rotation applied; optional speaker notes |

## Limits

- Scanned PDFs have no text layer; those pages are kept as pictures. No OCR.
- PDF hyperlinks, footnotes and form fields become plain text.
- PPTX charts, SmartArt and embedded objects become a placeholder line. EMF/WMF/SVG pictures are skipped.
- PPTX animations, transitions, slide backgrounds and layout artwork (logos, footers) are not carried over.
- Exact slide positions cannot be reproduced in a flowing Word document; vertical gaps and side-by-side columns approximate them.
- Password-protected PDFs are supported; the page asks for the password.

## Tests

```powershell
uv run python tests/make_fixtures.py
uv run pytest
```
