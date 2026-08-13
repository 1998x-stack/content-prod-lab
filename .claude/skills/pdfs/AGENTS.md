# AGENTS.md — PDF Processing Skill

## OVERVIEW
Core workflow discipline for every PDF operation: **render → verify → operate → re-verify**. Renderers are authoritative for visual QA; text extraction is not. After any edit, redaction, or OCR pass, re-render and confirm the result.

## STRUCTURE

- `SKILL.md` — entry point; routes to the specific `tasks/*.md` needed
- `scripts/` — 21 Python helpers (render, extract, edit, redact, OCR, convert, parity, preflight, smoke tests)
- `tasks/` — 16 per-task how-tos (read/review, extract, edit, convert, forms, redact, OCR, compare, parity, batch, preflight, coords, js tools, create)
- `examples/` — smoke_test.md, end-to-end usage walkthrough
- `js/` — pdfjs-dist + pdf-lib Node helpers: `extract_text_pdfjs.mjs`, `extract_form_fields.mjs`, `fill_form.mjs` (deps via `install_deps.sh`)
- `troubleshooting/` — environment and renderer issues
- Deps (system): LibreOffice (`lo_convert_to_pdf.py`), Poppler (`pdftoppm`/`pdftotext`/`pdfinfo`), tesseract + ocrmypdf (`ocr_pdf.py`), pandoc (`md_to_pdf.py`), latexmk (`latex_to_pdf.py`), Playwright (`html_to_pdf.py`)

## WHERE TO LOOK

- Read / inspect / extract → `scripts/pdf_extract.py` (subcommands: info, text, layout, coords, tables, images, forms) and `scripts/pdf_inspect.py`
- Edit / merge / split / rotate / watermark → `scripts/pdf_edit.py`
- Redact / sanitize → `scripts/pdf_redact.py`
- OCR → `scripts/ocr_pdf.py`
- Forms → `js/*.mjs` (extract text / fields, fill) + `scripts/place_text_by_boxes.py`, `forms_smoketest.py`
- Convert → `lo_convert_to_pdf.py` (Office→PDF), `md_to_pdf.py`, `latex_to_pdf.py`, `html_to_pdf.py`
- Verify → `scripts/render_pdf.py` (pdftoppm or pdfium engine), `compare_renders.py`, `renderer_parity.py`

## CONVENTIONS

- Author text-heavy or slide-like content in DOCX/PPTX (native tooling), then convert to PDF via LibreOffice
- Use ReportLab only for programmatic PDF generation
- Renderers are authoritative for visual QA; text extraction is NOT proof of layout
- At least 1 renderer for basic QA; use 2 engines when forms or tricky layout is involved
- Unicode em-dash → ASCII hyphen when stamping/placing text
- Read only the specific `tasks/*.md` reference needed, not all of them

## COMMANDS

- `python scripts/render_pdf.py in.pdf --out_dir out --dpi 150` → page PNGs for visual check
- `python scripts/pdf_extract.py in.pdf text` → extract text
- `python scripts/pdf_edit.py in.pdf --split 3 --out out.pdf` / `--merge a.pdf b.pdf` / `--rotate 90`
- `python scripts/pdf_redact.py in.pdf --text "regex" --out clean.pdf`
- `python scripts/ocr_pdf.py scanned.pdf` → text layer + copyable PDF
- `python scripts/lo_convert_to_pdf.py doc.docx out.pdf`
- `node js/fill_form.mjs in.pdf out.pdf '{"field":"value"}'`
- Flags above are approximate; exact flags live in `SKILL.md` / relevant task file

## ANTI-PATTERNS

- Delivering an edited PDF without a pixel-level render check
- Treating `pdftotext` output as proof of correct layout
- Scripting DOCX-style content generation with flat PDF libs when the DOCX path is cleaner
- Skipping the re-verify pass after redaction or OCR
- Editing a PDF in place with no backup of the original