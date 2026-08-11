# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this directory is

This is **not a software project** (no git repo, no build system, no tests). It is a working folder for two things:

1. **Claude Code skills** — reusable skill packages (each a self-contained directory with `SKILL.md`).
2. **Content production workspaces** — in-progress Chinese current-affairs articles, videos, and their final DOCX deliverables.

Nearly all produced content is in **Chinese** and is delivered as **`.docx`** files.

## Layout

**Skills** (all registered as project skills under `.claude/skills/<name>/`, each with its own `SKILL.md`):
- **`.claude/skills/docx/`** — a local standalone DOCX skill with a strict *render-and-verify* workflow (see below). This is the QA layer used by every other skill that ships a `.docx`.
- **`.claude/skills/article-to-short-video-script/`** — skill → finished WeChat/current-affairs article → production-ready short-video package (hook, rewritten voiceover, shot list, platform adaptation, DOCX). Full video re-engineering, not summarization.
- **`.claude/skills/wechat-high-energy-commentary/`** — skill → research → publish-ready WeChat Official Account article with images, delivered as DOCX.
- **`.claude/skills/geopolitical-deep-analysis-wechat/`** — skill → source-grounded geopolitical deep analysis on countries/conflicts/trade/energy/security, multi-school interpretation + scenario forecasting, delivered as Chinese WeChat article DOCX.
- **`.claude/skills/modern-qimin-jimeng-video/`** — skill → 现代齐民要术 life-science topics → 即梦 (Jimeng) AI animated short-video production scripts.
- **`.claude/skills/pdfs/`** — skill → PDF processing: render → verify → operate → re-verify (reading, OCR, forms, redaction, conversion, diff).
- **`.claude/skills/slides/`** — skill → PowerPoint/visual aid creation via pptxgenjs + bundled professional templates.
- **`.claude/skills/spreadsheets/`** — skill → spreadsheet handling (read, clean, convert, generate).
- **`dayouyuan/`** — working space for the WeChat-article → short-video pipeline. `prompts/` holds the writing prompts; finished deliverables live under `articles/YYYYMMDD/` (one dir per production date).
- **`qiminyaosu/`** — working space for 现代齐民要术: topic bibles (`现代齐民要术_六卷选题圣经/`, docx + xlsx), `prompts/`, `articles/`, `docs/`.
- **`geopolitical/`** — working space for geopolitical deep analysis (prompts/ + articles/YYYYMMDD/ deliverables).
- **`docs/gotchas.md`** — append-run gotchas journal; check it when hitting an environment or workflow problem.

## How the skills work

Each skill directory is a self-contained Claude Code project skill in `.claude/skills/<name>/`. Entry point is `SKILL.md`:
- YAML frontmatter `name:` + `description:` is what makes Claude detect and trigger the skill.
- The body routes to `references/*.md` — read only the specific reference needed for the task, not all of them.
- A `manifest.txt` lists the package files and is the machine-readable inventory.

The `docx` skill's internal structure: `render_docx.py` + `scripts/*.py` (one-off helpers, e.g. `comments_strip.py`, `redact_docx.py`, `a11y_audit.py`), `tasks/*.md` (per-task how-tos), `ooxml/*.md` (OOXML internals), `troubleshooting/*.md`.

## The non-negotiable render gate

Every skill that ships a `.docx` **must** verify it visually before delivery: run
```bash
cd .claude/skills/docx
python render_docx.py input.docx --output_dir out        # PNGs per page
python render_docx.py input.docx --output_dir out --emit_pdf   # + PDF, optional
```
then open the `page-<N>.png` images and confirm every page is clean (no clipping, overlap, missing glyphs, broken tables). Text extraction / reading XML misses layout defects. Rendered PNGs/PDFs are internal QA only — return only the final `.docx` unless the user asks for intermediates.

`render_docx.py` needs `pdf2image` + Poppler and LibreOffice (container-safe profile + writable HOME). If rendering fails, fix the LibreOffice profile first rather than skipping the gate.

## Content rules that recur across every writing skill

- **Never copy distinctive phrasing, famous lines, or memorable turns of a specific creator** — recreate the narrative mechanics from the facts of the current story.
- Keep **fact / inference / joke** strictly separated in the text.
- These are the consistent hard rules across the four skills and both prompt files.

## Local config

`.claude/settings.local.json` allow-lists `Read`/`Edit`/`Grep`/`WebFetch`/`WebSearch`/`Agent`/`Skill` — web search and Agent use are explicitly expected (research-heavy workflow), so don't treat them as surprises.

## Importing from another agent

There is a `~/.codex/config.toml` (Codex) — it defines only a model + reasoning-effort preference and per-path trust levels; there is nothing meaningful to import (no MCP servers, slash commands, subagents, skills, or instructions). To confirm, reply `/import` to scan the importable surface and apply with `/import --yes=<digest>`.