---
name: slides
description: Build, edit and export PowerPoint-style presentations with pptxgenjs or artifact tool library. Use when creating or modifying presentations or other visual aids like charts, posters etc.,
---


# Slides Skill

Use this skill as reference material when creating or editing presentation slide decks.

## Local environment (this machine)

- **Python interpreter:** use the managed venv: `/Users/x/.workbuddy/binaries/python/envs/default/bin/python` (has `python-pptx`, `numpy`, plus the office/pdf stack).
- **`container_tools/` scripts** are standalone and runnable with the venv python (they use python-pptx/numpy — no `artifact_tool` needed).
- **`artifact_tool/` and `pptxgenjs_helpers/`** require the claude.ai artifact runtime (`artifact_tool` / PptxGenJS), which is **not installed on this machine**. Prefer `container_tools/` + `python-pptx` for local work.
- **Scratch / output dirs:** use `/tmp` (or a workspace subfolder).

## Skill Folder Contents

Contents of the `slides/` skill folder:

- `container_tools/`: Standalone python scripts for slides and relevant asset manipulation.
- `artifact_tool/`: API documentation and coding examples for the artifact tool library.
- `pptxgenjs_helpers/`: JavaScript helpers for PptxGenJS.

## Implementation

You may choose whichever approach you think works best for this task. If it helps, feel free to use a template from the `slide_templates` folder (optional).
