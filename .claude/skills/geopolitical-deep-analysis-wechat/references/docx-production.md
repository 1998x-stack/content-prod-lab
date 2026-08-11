# DOCX production and visual QA

## Final deliverable

The formal output is a `.docx` file suitable for a WeChat public-account editing workflow.

Recommended filename:

`YYYY-MM-DD_选题_地缘政治深度分析.docx`

## Visual philosophy

Professional, restrained, modern, mobile-friendly after copy/paste into WeChat.

Avoid:

- decorative WordArt,
- heavy borders,
- excessive icons,
- saturated color blocks,
- report-style visual density.

Prefer:

- black/dark gray body text,
- one restrained accent color,
- generous white space,
- strong hierarchy.

## Page and typography guidance

Use stable, commonly available Chinese fonts. Prefer font stacks/fallbacks that minimize substitution issues.

Suggested hierarchy:

- Title: 20–24 pt, bold, centered
- Subtitle: 12–14 pt, subdued, centered
- Heading 1: 16–18 pt, bold, clear space before
- Heading 2: 13–15 pt, bold
- Body: 11.5–12 pt
- Caption/source: 9–10 pt

Use actual Word heading styles, not manual bolding alone.

Body guidance:

- 1.4–1.6 line spacing,
- 6–10 pt paragraph spacing after,
- left or justified alignment,
- avoid deep first-line indentation,
- keep visual breathing room.

## Core modules

Use restrained shading/borders for modules such as `核心判断` or `关键变量` when supported cleanly.

Do not turn every section into a colored card.

## Tables

- Keep narrow.
- Use repeatable header rows when relevant.
- Avoid text overflow.
- Prefer lists/cards when a table becomes too wide.

## Figures

- Maintain aspect ratio.
- Fit within page margins.
- Add a concise caption.
- Add source/data note.
- Use alt text when possible.

## Citations and links

Do not leave raw tool tokens or internal citation syntax in the final file.

Use normal readable citations and clickable hyperlinks.

A final `参考资料与延伸阅读` section should list the most important sources.

## Mandatory render-and-verify gate

A DOCX is not complete until it has been rendered to page images and visually inspected.

When the environment provides a DOCX renderer:

1. Generate the `.docx`.
2. Render all pages to PNG (and optional PDF for QA).
3. Inspect every page at normal/100% scale.
4. Check:
   - clipping,
   - overlap,
   - missing Chinese glyphs,
   - table overflow,
   - distorted images,
   - orphaned headings,
   - header/footer collisions,
   - inconsistent spacing,
   - broken hyperlinks/captions where detectable.
5. Fix defects.
6. Re-render.
7. Repeat until clean.
8. Deliver only the requested final `.docx` unless QA assets were explicitly requested.

## Final Word checklist

- semantic heading styles,
- consistent typography,
- mobile-readable paragraph rhythm,
- no oversized tables,
- no decorative clutter,
- captions/sources present,
- hyperlinks usable,
- all pages visually verified,
- document opens without repair warnings,
- no internal/tool artifacts exposed.
