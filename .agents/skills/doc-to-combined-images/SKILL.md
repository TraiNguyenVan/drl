---
name: doc-to-combined-images
description: Use when converting .doc/.docx/.pdf documents to images — rendering every page as PNG and merging all pages of each document into one combined image (side-by-side or vertical). Triggers: "render docs to images", "convert to png", "one image per document", "combine pages", "make page previews".
---

# Doc to Combined Images

Render Word/PDF documents to per-page PNGs, then merge each document's pages into a single combined image.

## When to Use
- User wants each document as one image (page previews, side-by-side review, mobile sharing)
- Batch-converting `.doc`/`.docx`/`.pdf` to images
- Per-page PNGs plus a combined variant

**Don't use** for: single page images only (use `pdftoppm` directly), or image manipulation unrelated to document rendering.

## Quick Reference
- Tool: `scripts/doc_to_images.py` (Python; needs `soffice`, `pdftoppm`, `PIL`)
- Default: 150 DPI, pages side-by-side, temp page dir auto-cleaned
- Combine layout: `--layout horizontal|vertical`, `--gap N`
- Trim pages: `--drop-first N`, `--drop-last N`
- Keep intermediate per-page PNGs: `--pages <dir>`

## Usage

```bash
python3 .agents/skills/doc-to-combined-images/scripts/doc_to_images.py \
  students/ --out combined/                    # all docs in a folder -> one image each
python3 .agents/skills/doc-to-combined-images/scripts/doc_to_images.py file1.docx file2.pdf --out out/ --dpi 200
python3 .agents/skills/doc-to-combined-images/scripts/doc_to_images.py students/ --out out/ --drop-last 1   # exclude last page
python3 .agents/skills/doc-to-combined-images/scripts/doc_to_images.py students/ --out out/ --layout vertical
python3 .agents/skills/doc-to-combined-images/scripts/doc_to_images.py students/ --out out/ --pages per_page/ # keep page PNGs too
```

Inputs: one or more files or directories. Every document produces one `<name>.png`.

## Pipeline
1. `.doc`/`.docx` → PDF via headless LibreOffice (`soffice --headless --convert-to pdf`)
2. PDF → PNG per page via `pdftoppm -png -r <dpi>`
3. Pages merged into one image per document (default side-by-side, white background)

## Common Mistakes
- `soffice` or `pdftoppm` missing → script prints a clear error; install `libreoffice` and `poppler-utils`
- Very wide merges (e.g. 8 pages) exceed Telegram photo limits (~10000 px wide) — resize down or send as document
- Vietnamese/spaced filenames work fine; always quote paths
