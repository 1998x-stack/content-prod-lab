# SKILLS KNOWLEDGE BASE

**Parent:** [AGENTS.md](../../AGENTS.md) · **Read first:** this file, then the skill's SKILL.md

## OVERVIEW
All 8 production capabilities live here as self-contained skill packages. Two families:
- **Content skills** (write/transform Chinese content → .docx): `wechat-high-energy-commentary`, `geopolitical-deep-analysis-wechat`, `article-to-short-video-script`, `modern-qimin-jimeng-video`
- **Tooling skills** (manipulate Office/PDF/table files, include the QA gates): `docx`, `pdfs`, `slides`, `spreadsheets`

## PACKAGE ANATOMY
| Part | Purpose |
|------|---------|
| `SKILL.md` | entry point; YAML frontmatter `name:` **must equal the directory name** or Claude won't index it |
| `references/*.md` | deep how-to guides loaded on demand — read the smallest relevant one, not all |
| `manifest.txt` | plain-file inventory (present in 5/8: all 4 content skills + docx) |
| `templates/` `examples/` | output skeletons / worked examples (content skills) |
| `tasks/` `scripts/` `ooxml/` `troubleshooting/` | detailed ops docs + py/js helpers (tool skills) |

Skill changes require session restart or `/reload-skills` before they take effect.

## WHERE TO LOOK
| Need | Package → entry |
|------|-----------------|
| Write WeChat commentary | `wechat-high-energy-commentary/SKILL.md` |
| Deep geopolitical analysis | `geopolitical-deep-analysis-wechat/SKILL.md` |
| Article → 短视频脚本 | `article-to-short-video-script/SKILL.md` |
| 六卷选题 → 即梦视频稿 | `modern-qimin-jimeng-video/SKILL.md` |
| Create/edit any .docx | `docx/SKILL.md` → `tasks/` |
| PDF ops (render/verify/redact) | `pdfs/SKILL.md` |
| PPT/slides | `slides/SKILL.md` (pptxgenjs helpers in `pptxgenjs_helpers/`, QA in `container_tools/`) |
| Spreadsheets | `spreadsheets/SKILL.md` |

## CONVENTIONS (non-standard)
- **Every .docx deliverable passes the docx skill's render gate** — `.claude/skills/docx/render_docx.py`, PNGs inspected per page, fix + re-render until flawless. Applies to ALL content skills' output.
- Content skills share hard rules: research-before-write (no drafting from memory), fact ≠ inference ≠ joke separation, never copy a creator's memorable phrasing.
- `modern-qimin-jimeng-video` hard constraints: generation unit ≤10s, every shot needs START_FRAME/END_FRAME/BRIDGE_TO_NEXT (no END_FRAME = unfinished), AI frames must not contain readable text/branding, AI visuals must not pose as evidence.
- `name:` in SKILL.md frontmatter must exactly match the package directory; ZIP-installed skills (e.g. typo-fixed `geopolitical-deep-analysis-wechat`) must be unpacked to that path, zip deleted.
- Spreadsheets skill forbids openpyxl/pandas — use the documented workbook API.

## ANTI-PATTERNS (THIS DIRECTORY)
- Editing a skill and assuming it's live — restart or `/reload-skills`, verify with "Reloaded skills: N available".
- Skipping a skill's own QA gate (docx/pdfs/slides all render→verify) to "save time".
- Mixing two skills in one deliverable (e.g. hand-formatting .docx instead of delegating to the docx skill).
- Leaving AI-image placeholders unlabeled where the content skill requires them to be clearly AI.

## NOTES
- `spreadsheets` (4 files) and the content skills (6-10 files) are intentionally lean; detailed conventions live in their SKILL.md + references, not here.
- The three tooling skills have their own AGENTS.md: [`docx/AGENTS.md`](docx/AGENTS.md), [`pdfs/AGENTS.md`](pdfs/AGENTS.md), [`slides/AGENTS.md`](slides/AGENTS.md).