# Batch processing

If you need to run the same operation on many PDFs, prefer a batch pattern with a clean output root.

## Golden paths

### Render a corpus
```bash
python batch_pdf.py render \
  --in_glob "/tmp/in/**/*.pdf" \
  --out_root /tmp/_renders \
  --dpi 200 --engine pdftoppm
```

### Inspect a corpus (JSON per file)
```bash
python batch_pdf.py inspect \
  --in_glob "/tmp/in/**/*.pdf" \
  --out_root /tmp/_inspect
```

### Normalize/repair a corpus
```bash
python batch_pdf.py normalize \
  --in_glob "/tmp/in/**/*.pdf" \
  --out_root /tmp/_normalized
```

## Notes
- Keep outputs separate by operation; avoid overwriting the input corpus.
- After batch edits, spot-check a few files via render + montage:
  - `python render_pdf.py one.pdf --out_dir /tmp/_one --dpi 200`
  - `python create_montage.py /tmp/_one --out /tmp/_one_montage.png`
