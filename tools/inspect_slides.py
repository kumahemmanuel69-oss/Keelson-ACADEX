#!/usr/bin/env python3
"""Inspect an uploaded KATON.pdf before deciding how to enhance it.

Answers the questions that decide the approach:
  - Is the slide text real, extractable text, or is it a raster scan?
  - Is there a separable full-page background image, or is everything flattened?
  - What fonts, sizes and colours are actually in use?
  - How much free space is there on each slide to add content?

Usage: python3 tools/inspect_slides.py uploads/KATON.pdf
"""
import sys, os, collections
import pymupdf


def main(path):
    doc = pymupdf.open(path)
    print("FILE      : %s" % path)
    print("SIZE      : %.2f MB" % (os.path.getsize(path) / 1e6))
    print("PAGES     : %d" % doc.page_count)
    r = doc[0].rect
    print("PAGE BOX  : %.1f x %.1f pt  (%.2f x %.2f in)  ratio %.3f"
          % (r.width, r.height, r.width / 72, r.height / 72, r.width / r.height))
    print("ENCRYPTED : %s" % doc.is_encrypted)
    print("=" * 78)

    # ---- text extractability -------------------------------------------------
    total_chars = 0
    pages_with_text = 0
    for p in doc:
        t = p.get_text("text").strip()
        total_chars += len(t)
        if len(t) > 20:
            pages_with_text += 1
    print("TEXT: %d chars over %d/%d pages (%.0f chars/page)"
          % (total_chars, pages_with_text, doc.page_count, total_chars / max(1, doc.page_count)))
    if total_chars < 100:
        print("  >>> Raster scan: text is NOT extractable. Enhancement must be")
        print("      image-based, and fonts cannot be reused programmatically.")
    else:
        print("  >>> Real text: fonts/sizes/colours are recoverable.")

    # ---- fonts ---------------------------------------------------------------
    fonts = collections.Counter()
    sizes = collections.Counter()
    colors = collections.Counter()
    for p in doc:
        for blk in p.get_text("dict")["blocks"]:
            for ln in blk.get("lines", []):
                for sp in ln["spans"]:
                    if sp["text"].strip():
                        fonts[(sp["font"], round(sp["size"], 1))] += len(sp["text"])
                        sizes[round(sp["size"], 1)] += len(sp["text"])
                        colors["#%06x" % (sp["color"] & 0xFFFFFF)] += len(sp["text"])
    print("\nFONTS (font, size) by character count:")
    for (f, s), c in fonts.most_common(20):
        print("   %-38s %6.1f pt   %6d chars" % (f, s, c))
    print("\nTOP SIZES:", ", ".join("%s pt (%d)" % (s, c) for s, c in sizes.most_common(8)))
    print("TOP COLOURS:", ", ".join("%s (%d)" % (c, n) for c, n in colors.most_common(10)))

    # ---- images --------------------------------------------------------------
    print("\nIMAGES")
    per_page = []
    full_page_bg = 0
    for i, p in enumerate(doc, 1):
        imgs = p.get_images(full=True)
        per_page.append(len(imgs))
        for im in imgs:
            try:
                bbox = p.get_image_bbox(im)
            except Exception:
                continue
            if bbox.width > p.rect.width * 0.9 and bbox.height > p.rect.height * 0.9:
                full_page_bg += 1
    print("  total image placements : %d" % sum(per_page))
    print("  pages with >=1 image   : %d" % sum(1 for n in per_page if n))
    print("  full-page background(s): %d" % full_page_bg)
    if full_page_bg >= doc.page_count * 0.8:
        print("  >>> A reusable full-page background exists on most pages.")
        print("      That background can be lifted straight out and reused.")
    print("  per page: %s" % per_page[:60])

    # ---- drawings / vector art ----------------------------------------------
    print("\nVECTOR PATHS (page 1): %d" % len(doc[0].get_drawings()))

    # ---- free space estimate -------------------------------------------------
    print("\nFREE SPACE (fraction of each page covered by text or images)")
    for i in range(min(doc.page_count, 60)):
        p = doc[i]
        used = pymupdf.Rect()
        for blk in p.get_text("blocks"):
            used |= pymupdf.Rect(blk[:4])
        area = 0.0
        if not used.is_empty:
            area = used.get_area() / (p.rect.width * p.rect.height)
        print("  p%02d  text covers %4.1f%%   imgs=%d" % (i + 1, area * 100, per_page[i]))

    # ---- per-page first line, to compare against the rebuilt deck ------------
    print("\n" + "=" * 78)
    print("FIRST LINE OF EACH PAGE")
    for i, p in enumerate(doc, 1):
        lines = [l.strip() for l in p.get_text("text").split("\n") if l.strip()]
        head = " | ".join(lines[:3])[:110] if lines else "(no text)"
        print("  %2d  %s" % (i, head))


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "uploads/KATON.pdf")
