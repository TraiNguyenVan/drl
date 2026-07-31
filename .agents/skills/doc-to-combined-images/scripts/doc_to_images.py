#!/usr/bin/env python3
import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
from PIL import Image


def find_inputs(paths):
    files = []
    for p in paths:
        if os.path.isdir(p):
            for name in sorted(os.listdir(p)):
                if name.lower().endswith((".doc", ".docx")):
                    files.append(os.path.join(p, name))
        elif p.lower().endswith((".doc", ".docx")):
            files.append(p)
        elif p.lower().endswith(".pdf"):
            files.append(p)
        else:
            print(f"Skip (not doc/docx/pdf): {p}", file=sys.stderr)
    return files


def to_pdf(src, outdir, soffice):
    if src.lower().endswith(".pdf"):
        dest = os.path.join(outdir, os.path.basename(src))
        shutil.copy2(src, dest)
        return dest
    subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", "--outdir", outdir, src],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=True,
    )
    return os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + ".pdf")


def render_pages(pdf_path, outdir, dpi):
    stem = os.path.splitext(os.path.basename(pdf_path))[0]
    prefix = os.path.join(outdir, stem + "-")
    subprocess.run(
        ["pdftoppm", "-png", "-r", str(dpi), pdf_path, prefix],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=True,
    )
    pages = sorted(
        (p for p in os.listdir(outdir) if p.startswith(stem + "-") and p.endswith(".png")),
        key=lambda n: int(n[len(stem) + 1 : -4]),
    )
    return [os.path.join(outdir, p) for p in pages]


def combine(pages, dest, layout, gap):
    imgs = [Image.open(p).convert("RGB") for p in pages]
    if layout == "horizontal":
        w = sum(i.width for i in imgs) + gap * (len(imgs) - 1)
        h = max(i.height for i in imgs)
        canvas = Image.new("RGB", (w, h), "white")
        x = 0
        for i in imgs:
            canvas.paste(i, (x, 0))
            x += i.width + gap
    else:
        w = max(i.width for i in imgs)
        h = sum(i.height for i in imgs) + gap * (len(imgs) - 1)
        canvas = Image.new("RGB", (w, h), "white")
        y = 0
        for i in imgs:
            canvas.paste(i, (0, y))
            y += i.height + gap
    canvas.save(dest)
    return (w, h)


def main():
    ap = argparse.ArgumentParser(description="Convert doc/docx/pdf files to one combined image per file (pages side-by-side).")
    ap.add_argument("paths", nargs="+", help="Input files or directories")
    ap.add_argument("--out", required=True, help="Output directory for combined PNGs")
    ap.add_argument("--pages", required=False, help="Page dir for intermediate per-page PNGs (default: temp, auto-cleaned)")
    ap.add_argument("--dpi", type=int, default=150)
    ap.add_argument("--layout", choices=["horizontal", "vertical"], default="horizontal")
    ap.add_argument("--gap", type=int, default=0, help="White gap in px between pages")
    ap.add_argument("--drop-first", type=int, default=0, help="Drop first N pages per document")
    ap.add_argument("--drop-last", type=int, default=0, help="Drop last N pages per document")
    args = ap.parse_args()

    for tool in ("soffice", "pdftoppm"):
        if not shutil.which(tool):
            print(f"Error: '{tool}' not found on PATH.", file=sys.stderr)
            sys.exit(1)

    os.makedirs(args.out, exist_ok=True)
    workdir = tempfile.mkdtemp(prefix="doc_to_images_")
    pdf_dir = os.path.join(workdir, "pdf")
    page_dir = args.pages or os.path.join(workdir, "pages")
    for d in (pdf_dir, page_dir):
        os.makedirs(d, exist_ok=True)

    files = find_inputs(args.paths)
    if not files:
        print("No .doc/.docx/.pdf files found.", file=sys.stderr)
        sys.exit(1)

    ok, failed = 0, []
    for f in files:
        stem = os.path.splitext(os.path.basename(f))[0]
        try:
            pdf = to_pdf(f, pdf_dir, "soffice")
            pages = render_pages(pdf, page_dir, args.dpi)
            if args.drop_first:
                pages = pages[args.drop_first:]
            if args.drop_last:
                pages = pages[:-args.drop_last]
            if not pages:
                print(f"  {stem}: no pages after trimming", file=sys.stderr)
                failed.append(stem)
                continue
            dest = os.path.join(args.out, stem + ".png")
            size = combine(pages, dest, args.layout, args.gap)
            print(f"  {stem}: {len(pages)} pages -> {size[0]}x{size[1]}")
            ok += 1
        except Exception as e:
            print(f"  {stem}: FAILED ({e})", file=sys.stderr)
            failed.append(stem)

    shutil.rmtree(workdir, ignore_errors=True)
    print(f"\nDone: {ok} combined images written to {args.out}")
    if failed:
        print("Failed:", ", ".join(failed), file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
