# AGENTS.md — docx skill package

## OVERVIEW
The QA/rendering layer for every skill that ships a DOCX. Core loop: author → render → visually verify → deliver.

## STRUCTURE
- `SKILL.md` — entry point; `name` in frontmatter must match the directory name
- `manifest.txt` — machine-readable inventory; relative paths, one per line
- `render_docx.py` — the render gate: docx → page PNGs + optional PDF
- `scripts/` — 34 one-shot Python CLI helpers (positional `in.docx` + `--out`), incl. `redact_docx.py`, `accept_tracked_changes.py`, `comments_*.py`, `fields_materialize.py`, `insert_toc.py`, `internal_nav.py`, `heading_audit.py`, `style_lint.py`, `a11y_audit.py`, `render_and_diff.py`, `xlsx_to_docx_table.py`
- `tasks/` — 25 per-task how-tos (e.g. verify-render workflow)
- `ooxml/` — 4 internals docs: comments, hyperlinks_and_fields, rels_and_content_types, tracked_changes
- `troubleshooting/` — libreoffice_headless, run_splitting
- `examples/` — smoke test

## WHERE TO LOOK
- Create or edit a docx → `tasks/` (python-docx authoring path)
- Comments / tracked changes / fields → `ooxml/` for the format, matching script in `scripts/` for the operation
- Column/table conversions → `xlsx_to_docx_table.py`
- Layout defect in a rendered page → `troubleshooting/`, then `ooxml/`
- Render hangs or crashes → `troubleshooting/libreoffice_headless`

## CONVENTIONS
- Author with python-docx; patch OOXML (via scripts) only when python-docx can't express it: tracked changes, comments, fields.
- After ANY meaningful edit batch → re-render and review the PNGs. No exceptions.
- Every shipped docx must pass the visual gate first.
- PNGs/PDFs are internal QA only; return only the final .docx.
- Render failure ⇒ fix the LibreOffice profile before skipping the gate.
- Fields (REF/SEQ) must be materialized for deterministic rendering.
- Heading numbering must not skip levels.

## COMMANDS
```bash
cd .claude/skills/docx
python render_docx.py input.docx --output_dir out
python render_docx.py input.docx --output_dir out --emit_pdf
```
Deps: pdf2image, Poppler, LibreOffice (container-safe profile, writable HOME).

## ANTI-PATTERNS
- Shipping a docx verified only by text extraction (misses layout defects)
- Skipping re-render after an OOXML patch
- Comments with anchorless ranges (null text) — unverifiable in render
- Leaving tool-citation tokens in the delivered doc
- Jumping heading levels (H1→H3)
- Manual bold in place of semantic heading styles