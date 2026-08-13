# SKILL KNOWLEDGE BASE: slides

**Generated:** 2026-08-13

## OVERVIEW
PPT/slides production for Chinese content labs: build decks with **PptxGenJS (node)** or the bundled **artifact tool (python)**, both backed by professional templates, then pass a render-and-verify QA gate (PPTX → page PNGs) before shipping. Entry point is `SKILL.md`.

## STRUCTURE
```
slides/
├── SKILL.md                  # entry point; routes to helpers/tools/templates
├── pptxgenjs_helpers/        # CommonJS helper module for generated slide code:
│   │                           layout.js, text.js, image.js, svg.js,
│   │                           latex.js, code.js, util.js (+ index.js, layout_builders.js)
├── container_tools/          # render_slides.py (PPTX→PNG QA gate),
│                               slides_test.py (overflow QA), create_montage.py,
│                               detect_font.py, ensure_raster_image.py
├── artifact_tool/            # python library (theme/fill/layout/charts/... specs) + examples/
└── slide_templates/          # 16 .pptx templates + Overview.png
```

## WHERE TO LOOK
| Need | Location |
|---|---|
| Build/deck pipeline | `SKILL.md` |
| PptxGenJS layout/text/image helpers | `pptxgenjs_helpers/` |
| Render-and-verify gate, overflow QA, font checks | `container_tools/` |
| Python deck building | `artifact_tool/` (+ `examples/` for quick starts) |
| Templates: Academic_*, Brand_Design_*, Consulting_Proposal_*, Market_*, Project_Kick-off_*, Pitch_Deck | `slide_templates/` |

## CONVENTIONS
- PptxGenJS layout helpers require the shared helpers module (`pptxgenjs_helpers/`) — never inline-duplicate its logic.
- Every deck must be rendered with `render_slides.py` and the PNGs visually verified before delivery; text extraction misses layout defects.
- Pick templates by audience/theme (Academic vs Consulting vs Pitch), not by creator preference.
- Font sizes in helpers must be numbers in **points (PT)**; run `detect_font.py` when glyphs render wrong / missing.
- Keep one build path per deck: don't mix node+PptxGenJS and artifact_tool in the same deck.

## COMMANDS
```bash
node <snippet>.js                                     # build the deck
python container_tools/render_slides.py deck.pptx --output_dir out   # QA gate → page-*.png
python container_tools/slides_test.py                 # overflow QA
```

## ANTI-PATTERNS
- Shipping a deck without rendering it at least once.
- Pixel/point unit confusion inside helper calls.
- Mixing artifact_tool and PPTX helpers in the same deck.
- Skipping the template list when starting a new deck.