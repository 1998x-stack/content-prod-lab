# PROJECT KNOWLEDGE BASE

**Generated:** 2026-08-13
**Commit:** dd27ee0
**Branch:** main

## OVERVIEW
Chinese content-production lab: Claude Code **skills** (`.claude/skills/`) + **workspaces** (`dayouyuan/`, `qiminyaosu/`, `geopolitical/`) that turn researched current-affairs topics into WeChat articles, short-video production scripts, and life-science videos. All article deliverables are Chinese **.docx** files. NOT a software project — no build, no tests; the "codebase" is skill packages (SKILL.md + references + py/js helpers).

## STRUCTURE
```
content-prod-lab/
├── CLAUDE.md        # authoritative project guide — read FIRST, this file complements it
├── README.md        # user-facing overview
├── .claude/
│   ├── settings*.json    # note: TWO files, setting.local.json (authoritative) vs settings.local.json (likely stale)
│   └── skills/           # 8 skill packages — see .claude/skills/AGENTS.md
├── docs/gotchas.md  # append-run gotchas journal — check when hitting env/workflow problems
├── dayouyuan/       # 公众号文章 → 短视频 流水线 (prompts/ + articles/YYYYMMDD/ deliverables)
├── geopolitical/    # 地缘政治深度分析 workspace (prompts/ + articles/YYYYMMDD/)
├── qiminyaosu/      # 现代齐民要术 workspace (选题圣经 docx/xlsx, prompts/, articles/)
├── opencode/        # OpenCode/Claude Code deliverables: cookbook/, 教程/, 汇报/, 方案/, 深度研究/
├── series/          # 系列长文 (每系列一子目录: 钢斧之后/ — 命名 系列名_第X章_标题)
├── reports/         # 城市/机构年度报告深度研究 (命名 城市_年份政府工作报告_深度研究_日期)
├── proposals/       # 技术/产品方案 (命名 主题_方案名_v版本)
└── .omo/            # agent runtime state — never touch
```

## WHERE TO LOOK
| Task | Location |
|---|---|
| WeChat commentary article | `.claude/skills/wechat-high-energy-commentary/` |
| Geopolitical deep analysis | `.claude/skills/geopolitical-deep-analysis-wechat/` |
| Article → short-video script | `.claude/skills/article-to-short-video-script/` |
| 齐民要术 → 即梦 video | `.claude/skills/modern-qimin-jimeng-video/` |
| DOCX create/edit/quit the red-line gate | `.claude/skills/docx/` (QA layer for all .docx) |
| PDF read/edit/redact | `.claude/skills/pdfs/` |
| PPT / slides | `.claude/skills/slides/` |
| Spreadsheets | `.claude/skills/spreadsheets/` |
| OpenCode docs/decks/proposals | `opencode/` (cookbook/ 教程/ 汇报/ 方案/ 深度研究) |
| Long-form article series | `series/<系列名>/` (e.g. `series/钢斧之后/`) |
| City/gov annual-report deep dives | `reports/` |
| Tech/product proposals | `proposals/` |
| Gotchas / environment issues | `docs/gotchas.md` |

## CODE MAP (entry points)
| Script | Path | Role |
|--------|------|------|
| `render_docx.py` | `.claude/skills/docx/` | THE QA gate every .docx deliverable must pass (render → PNGs → visual check) |
| `render_pdf.py` `pdf_edit.py` | `.claude/skills/pdfs/scripts/` | PDF render/verify + structural edit |
| `render_slides.py` | `.claude/skills/slides/container_tools/` | PPTX → PNG QA gate for slides |
| `pptxgenjs_helpers/*.js` | `.claude/skills/slides/` | Node helpers used by generated slide code |
| SKILL.md per package | `.claude/skills/<name>/` | entry point; routes to `references/*.md` |

## CONVENTIONS (non-standard)
- **Skill dir = package**: frontmatter `name:` must equal directory name or the skill won't index; manifests `manifest.txt` list package files.
- **Render-and-verify gate is non-negotiable** for any .docx deliverable (see COMMANDS). Text extraction / reading XML misses layout defects.
- **Skill loading**: new/changed skills require session restart or `/reload-skills`.
- **Chinese-only content**, delivered as `.docx`. Workspaces archive by production date: `articles/YYYYMMDD/`.
- Research-first: current-affairs topics researched via web search before writing (settings allow WebSearch/Agent/Skill explicitly).

## ANTI-PATTERNS (THIS PROJECT)
- Never copy a creator's distinctive phrasing / famous lines — rebuild narrative mechanics from the facts.
- Keep fact / inference / joke strictly separated in the text (fact: 直接陈述；推断: 加限定词；玩梗: 明显是玩梗). Never upgrade "A happened before B" into "A caused B".
- Images must explain, not decorate | AI images must not masquerade as real news photos | don't embed unlicensed news images in publications.
- Never `meme-first` writing: 永远不要先想梗再往里面塞新闻；勿用无证据的确定性语言（毫无疑问/必然/已经证明/彻底失败）.
- Never ship a .docx without running the render gate — and never skip fixing LibreOffice to skip it.
- Don't commit junk: `.gitignore` guards `*.zip`/`.DS_Store`; `git add -A` is a footgun (137MB skill zip was committed before).
- Don't `sleep`-poll in macOS (no GNU `timeout`) — use background tasks.

## COMMANDS
```bash
# Render gate for any .docx deliverable (run from skill dir)
cd .claude/skills/docx
python render_docx.py input.docx --output_dir out            # PNGs per page
python render_docx.py input.docx --output_dir out --emit_pdf   # + PDF (optional)
```
Then open `page-*.png`; iterate render until every page is clean. PNGs/PDFs are internal QA only — ship only the final .docx.

## NOTES
- GitHub big pushes: use `ssh.github.com:443` (see `docs/gotchas.md`).
- Check `docs/gotchas.md` FIRST when hitting any env/workflow problem — it's append-only, add new lessons there.
- `write-geoplitical-article.md` typo was fixed → `geopolitical/prompts/write-geopolitical-article.md` (2026-08-13).