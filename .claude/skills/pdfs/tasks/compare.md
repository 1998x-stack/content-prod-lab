# Visual regression: compare two PDFs

Use this when:
- you edited a PDF (crop/rotate/watermark/paginate/fill) and need confidence nothing broke
- you suspect a "looks fine in viewer" vs "broken render" issue

---

## Golden path

```bash
python scripts/compare_renders.py before.pdf after.pdf \
  --out_dir /tmp/_diff \
  --dpi 200 --engine pdfium
```

Outputs:
- `/tmp/_diff/summary.json`
- `/tmp/_diff/diff/page-<N>.png` for changed pages

Success criteria:
- if you expect a small change (e.g., watermark), only those pages should diff
- if you expect *no* visual change (e.g., metadata-only), there should be 0 changed pages

Tip: if you want a human skim, generate montages:

```bash
python scripts/create_montage.py /tmp/_diff/render_a --out /tmp/_diff/a_montage.png
python scripts/create_montage.py /tmp/_diff/render_b --out /tmp/_diff/b_montage.png
```