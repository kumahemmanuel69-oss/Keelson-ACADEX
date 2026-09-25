#!/usr/bin/env python3
"""
Build the KATON 2026 'Effective Assessment Practices' workshop deck.

Source of truth: Teacher Assessment Manual and Toolkit (NaCCA / Ministry of Education, Ghana).
Audience: Senior High School teachers.
Design goal: every technical term gets a PLAIN ENGLISH meaning + a GHANAIAN CLASSROOM PICTURE.
"""
import os
import glob
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

ROOT = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(ROOT, "assets", "opt")
DIAGRAMS = os.path.join(ROOT, "assets", "diagrams")
SRC_ASSETS = os.path.join(ROOT, "assets")


def ensure_assets(max_w=1600, quality=86):
    """Regenerate the width-optimised JPEGs from the source PNGs when missing.

    assets/opt is a derived folder and is gitignored, so a fresh checkout (or a
    sandbox reset) leaves it empty. Without this the deck silently builds with
    placeholder boxes instead of illustrations.
    """
    os.makedirs(ASSETS, exist_ok=True)
    made = 0
    for src in sorted(glob.glob(os.path.join(SRC_ASSETS, "*.png"))):
        stem = os.path.splitext(os.path.basename(src))[0]
        dst = os.path.join(ASSETS, stem + ".jpg")
        if os.path.exists(dst):
            continue
        im = Image.open(src).convert("RGB")
        w, h = im.size
        if w > max_w:
            im = im.resize((max_w, int(h * max_w / w)), Image.LANCZOS)
        im.save(dst, "JPEG", quality=quality, optimize=True)
        made += 1
    if made:
        print(f"generated {made} optimised image(s) into assets/opt/")
    return made


def ensure_diagrams():
    """Redraw the teaching diagrams if assets/diagrams is missing.

    Same reasoning as ensure_assets(): a sandbox reset can strip untracked
    files, and a build with missing diagrams would silently fall back to
    placeholder boxes.
    """
    want = ["dartboard", "dok_staircase", "cycle", "split_70_30", "transcript"]
    missing = [w for w in want
               if not os.path.exists(os.path.join(DIAGRAMS, w + ".png"))]
    if not missing:
        return 0
    script = os.path.join(ROOT, "tools", "make_diagrams.py")
    if not os.path.exists(script):
        print(f"WARNING: {len(missing)} diagram(s) missing and no generator found")
        return 0
    import subprocess, sys as _sys
    subprocess.run([_sys.executable, script], check=False)
    print(f"regenerated diagrams: {', '.join(missing)}")
    return len(missing)


ensure_diagrams()
ensure_assets()
OUT = os.path.join(ROOT, "deliverables", "Effective-Assessment-Practices-KATON-2026.pptx")

# ---------------------------------------------------------------- design system
GREEN      = RGBColor(0x0B, 0x6E, 0x4F)
GREEN_DK   = RGBColor(0x07, 0x4A, 0x35)
GREEN_LT   = RGBColor(0xE3, 0xF1, 0xEA)
GOLD       = RGBColor(0xF4, 0xB4, 0x00)
GOLD_LT    = RGBColor(0xFD, 0xF3, 0xD8)
TERRA      = RGBColor(0xC1, 0x44, 0x0E)
TERRA_LT   = RGBColor(0xFB, 0xE9, 0xE0)
INK        = RGBColor(0x14, 0x26, 0x2E)
GREY       = RGBColor(0x5A, 0x6B, 0x72)
CREAM      = RGBColor(0xFD, 0xF8, 0xF0)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
PAPER      = RGBColor(0xF7, 0xF9, 0xF8)

HEAD_FONT = "Trebuchet MS"
BODY_FONT = "Calibri"

SW, SH = 13.333, 7.5          # slide size, inches
M = 0.62                      # side margin
CW = SW - 2 * M               # content width

prs = Presentation()
prs.slide_width = Inches(SW)
prs.slide_height = Inches(SH)
BLANK = prs.slide_layouts[6]


# ---------------------------------------------------------------- helpers
def solid(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def rect(slide, l, t, w, h, color, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, Inches(l), Inches(t), Inches(w), Inches(h))
    return solid(s, color)


def textbox(slide, l, t, w, h):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    return tf


def para(tf, text, size, color, bold=False, font=BODY_FONT, first=False,
         space_before=0, space_after=6, align=PP_ALIGN.LEFT, italic=False,
         line_spacing=1.0):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    p.line_spacing = line_spacing
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    r.font.name = font
    return p


def rich(tf, parts, size, first=False, space_before=0, space_after=6,
         align=PP_ALIGN.LEFT, line_spacing=1.0):
    """parts = [(text, color, bold, italic), ...]"""
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    p.line_spacing = line_spacing
    for text, color, bold, italic in parts:
        r = p.add_run()
        r.text = text
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = color
        r.font.name = BODY_FONT
    return p


def picture_cover(slide, name, l, t, w, h):
    """Place image filling the box exactly, cropping the overflow."""
    path = os.path.join(ASSETS, name)
    if not os.path.exists(path):
        # graceful fallback while an illustration is still being generated
        fb = rect(slide, l, t, w, h, GREEN_LT)
        tf = fb.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        para(tf, "[ illustration pending ]", 12, GREEN, first=True,
             align=PP_ALIGN.CENTER, space_after=0)
        return fb
    iw, ih = Image.open(path).size
    box_ar, img_ar = w / h, iw / ih
    pic = slide.shapes.add_picture(path, Inches(l), Inches(t), Inches(w), Inches(h))
    if img_ar > box_ar:
        f = (1 - box_ar / img_ar) / 2
        pic.crop_left = f
        pic.crop_right = f
    else:
        f = (1 - img_ar / box_ar) / 2
        pic.crop_top = f
        pic.crop_bottom = f
    return pic


def picture_fit(slide, name, l, t, w, h):
    """Place an image fully inside the box, centred, never cropped.

    picture_cover() crops, which is right for photographs but destroys a
    diagram whose labels run to the edge.
    """
    path = os.path.join(DIAGRAMS, name)
    if not os.path.exists(path):
        fb = rect(slide, l, t, w, h, GREEN_LT)
        tf = fb.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        para(tf, "[ diagram pending ]", 12, GREEN, first=True,
             align=PP_ALIGN.CENTER, space_after=0)
        return fb
    iw, ih = Image.open(path).size
    ar = iw / ih
    if w / h > ar:                 # box wider than image -> fit height
        nh, nw = h, h * ar
    else:                          # box taller than image -> fit width
        nw, nh = w, w / ar
    return slide.shapes.add_picture(path, Inches(l + (w - nw) / 2),
                                    Inches(t + (h - nh) / 2),
                                    Inches(nw), Inches(nh))


def slide_diagram(img, title, sub=None, caption=None, note_text="", accent=GREEN):
    """A full-width teaching diagram with its own notes."""
    s = new()
    y = title_bar(s, title, sub)
    box_h = SH - y - 1.34
    picture_fit(s, img, M, y + 0.04, CW, box_h)
    if caption:
        tf = textbox(s, M, y + box_h + 0.12, CW, 0.44)
        para(tf, caption, 12, GREY, first=True, italic=True,
             align=PP_ALIGN.CENTER, space_after=0, line_spacing=1.06)
    footer(s)
    notes(s, note_text)
    return s


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def new():
    return prs.slides.add_slide(BLANK)


def footer(slide, dark=False):
    """Slide number + running brand."""
    n = len(prs.slides._sldIdLst)
    col = RGBColor(0xB9, 0xCF, 0xC7) if dark else RGBColor(0x9A, 0xA8, 0xAE)
    tf = textbox(slide, SW - M - 2.4, SH - 0.46, 2.4, 0.28)
    para(tf, f"Effective Assessment Practices  ·  {n}", 9, col,
         first=True, align=PP_ALIGN.RIGHT, space_after=0)


def chip(slide, l, t, text, bg, fg, w=None, size=10.5):
    w = w or (0.26 + 0.105 * len(text))
    s = rect(slide, l, t, w, 0.30, bg, MSO_SHAPE.ROUNDED_RECTANGLE)
    s.adjustments[0] = 0.45
    tf = s.text_frame
    tf.margin_left = tf.margin_right = Inches(0.06)
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    para(tf, text.upper(), size, fg, bold=True, font=HEAD_FONT, first=True,
         align=PP_ALIGN.CENTER, space_after=0)
    return s


def title_bar(slide, text, sub=None):
    """Standard content-slide header with green accent."""
    rect(slide, 0, 0, SW, 0.115, GREEN)
    rect(slide, M, 0.50, 0.075, 0.52, GOLD)
    tf = textbox(slide, M + 0.22, 0.44, CW - 0.22, 0.66)
    para(tf, text, 27, INK, bold=True, font=HEAD_FONT, first=True, space_after=0)
    y = 1.14
    if sub:
        tf2 = textbox(slide, M + 0.22, 1.10, CW - 0.22, 0.32)
        para(tf2, sub, 13, GREY, first=True, italic=True, space_after=0)
        y = 1.50
    return y


def card(slide, l, t, w, h, title, body, fill=WHITE, accent=GREEN, tsize=13,
         bsize=11.5):
    rect(slide, l, t, w, h, fill, MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.06
    rect(slide, l, t, 0.055, h, accent)
    tf = textbox(slide, l + 0.26, t + 0.19, w - 0.5, h - 0.34)
    if title:
        para(tf, title, tsize, accent, bold=True, font=HEAD_FONT, first=True,
             space_after=4)
        for b in body:
            para(tf, b, bsize, INK, space_after=5, line_spacing=1.06)
    else:
        for i, b in enumerate(body):
            para(tf, b, bsize, INK, first=(i == 0), space_after=5, line_spacing=1.06)


def table(slide, l, t, w, headers, rows, col_w=None, row_h=0.44, hdr_h=0.42,
          fsize=11.5, hsize=11.5):
    """Hand-built table (full styling control, no theme surprises)."""
    n = len(headers)
    col_w = col_w or [w / n] * n
    # header
    x = l
    rect(slide, l, t, w, hdr_h, GREEN_DK)
    for i, htext in enumerate(headers):
        tf = textbox(slide, x + 0.13, t + 0.09, col_w[i] - 0.26, hdr_h - 0.1)
        para(tf, htext, hsize, WHITE, bold=True, font=HEAD_FONT, first=True,
             space_after=0)
        x += col_w[i]
    # body
    y = t + hdr_h
    for ri, row in enumerate(rows):
        bg = WHITE if ri % 2 == 0 else PAPER
        rect(slide, l, y, w, row_h, bg)
        x = l
        for i, cell in enumerate(row):
            tf = textbox(slide, x + 0.13, y + 0.07, col_w[i] - 0.26, row_h - 0.1)
            para(tf, cell, fsize, INK, bold=(i == 0), first=True, space_after=0,
                 line_spacing=1.0)
            x += col_w[i]
        y += row_h
    return y


# ================================================================ SLIDES
def slide_title():
    s = new()
    picture_cover(s, "01_hero_classroom.jpg", 0, 0, SW, SH)
    ov = rect(s, 0, 0, SW, SH, GREEN_DK)
    ov.fill.fore_color.rgb = GREEN_DK
    ov.fill.transparency = 0.0
    # transparency needs XML tinkering; use a semi-opaque band instead
    ov._element.getparent().remove(ov._element)
    rect(s, 0, 0, SW, SH, INK).fill.transparency = 0.0
    # proper alpha overlay
    from pptx.oxml.ns import qn
    ov2 = rect(s, 0, 0, SW, SH, GREEN_DK)
    solid_el = ov2.fill._xPr.find(qn('a:solidFill'))
    clr = solid_el.find(qn('a:srgbClr'))
    alpha = clr.makeelement(qn('a:alpha'), {'val': '82000'})
    clr.append(alpha)
    rect(s, 0, 0, SW, 0.13, GOLD)
    tf = textbox(s, M + 0.15, 1.85, CW - 0.3, 3.4)
    para(tf, "Effective Assessment Practices", 50, WHITE, bold=True,
         font=HEAD_FONT, first=True, space_after=10)
    para(tf, "Plain words. Everyday classrooms. Real Ghanaian schools.", 21, GOLD,
         font=HEAD_FONT, space_after=20)
    para(tf, "A practical workshop for Senior High School teachers — every technical "
             "term explained in plain English, with a classroom picture you already know.",
         14, RGBColor(0xD8, 0xE8, 0xE1), space_after=0, line_spacing=1.15)
    tf2 = textbox(s, M + 0.15, SH - 1.05, CW - 0.3, 0.65)
    para(tf2, "KATON 2026", 15, WHITE, bold=True, font=HEAD_FONT, first=True,
         space_after=2)
    para(tf2, "Based on the Teacher Assessment Manual and Toolkit — NaCCA, "
              "Ministry of Education, Republic of Ghana", 10.5,
         RGBColor(0xC6, 0xDB, 0xD3), space_after=0)
    notes(s, "WELCOME AND OPENING (2 min)\n\n"
             "Good morning, colleagues. Thank you for coming.\n\n"
             "Let me start with a promise: over the next session I am not going to bury "
             "you under big words. Every technical term we meet, I will give you two things — "
             "the plain meaning, and a picture from a classroom like ours.\n\n"
             "This workshop is built on the Teacher Assessment Manual and Toolkit produced by "
             "NaCCA under the Ministry of Education. It is the document that guides how we are "
             "expected to assess learners under the new SHS curriculum — so this is not our own "
             "invention. It is the standard we are working to.\n\n"
             "By the time we finish, my hope is that assessment stops feeling like paperwork, "
             "and starts feeling like something you already do — just done more deliberately.")
    return s


def slide_promises():
    s = new()
    y = title_bar(s, "Three Promises for Today",
                  "What you should be able to do before you leave this room")
    items = [
        ("1", "You will not be drowned in jargon",
         "Every technical term gets a plain-English meaning — one sentence, no showing off.",
         GREEN),
        ("2", "Every idea comes with a classroom picture",
         "A Ghanaian SHS situation you will recognise, so the term sticks.",
         GOLD),
        ("3", "You leave with something to use on Monday",
         "Three-minute tools you can try in your very next lesson, no budget required.",
         TERRA),
    ]
    x, w, gap = M, 3.85, 0.30
    for num, head, body, col in items:
        rect(s, x, y + 0.10, w, 2.45, WHITE, MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.05
        rect(s, x, y + 0.10, w, 0.075, col)
        c = rect(s, x + 0.32, y + 0.52, 0.62, 0.62, col, MSO_SHAPE.OVAL)
        tfc = c.text_frame
        tfc.vertical_anchor = MSO_ANCHOR.MIDDLE
        para(tfc, num, 22, WHITE, bold=True, font=HEAD_FONT, first=True,
             align=PP_ALIGN.CENTER, space_after=0)
        tf = textbox(s, x + 0.32, y + 1.30, w - 0.64, 1.0)
        para(tf, head, 14, INK, bold=True, font=HEAD_FONT, first=True, space_after=6,
             line_spacing=1.05)
        para(tf, body, 11.5, GREY, space_after=0, line_spacing=1.1)
        x += w + gap
    tf = textbox(s, M, y + 2.80, CW, 0.78)
    rich(tf, [("A word before we start:  ", GREEN, True, False),
              ("assessment is not extra work added on top of teaching. It is the thing that tells you "
               "whether your teaching landed. When it is done well, it actually ", INK, False, False),
              ("reduces", INK, True, True),
              (" your work — because you stop guessing what to re-teach.", INK, False, False)],
         12.5, first=True, line_spacing=1.15)
    footer(s)
    notes(s, "SET THE FRAME (1 min)\n\n"
             "Ask the room: 'How many of us were taught how to assess, versus how many of us were "
             "just told to assess?' Let hands go up. This normalises the gap and buys you goodwill.\n\n"
             "Then make the key point at the bottom of the slide: assessment is not an add-on. "
             "A teacher who assesses well does LESS wasted work, because they know exactly what "
             "to re-teach instead of re-teaching everything.")
    return s


def slide_section(num, title, sub, img=None, notes_text=""):
    s = new()
    if img:
        picture_cover(s, img, SW / 2, 0, SW / 2, SH)
    rect(s, 0, 0, SW / 2, SH, GREEN_DK)
    rect(s, 0, 0, 0.14, SH, GOLD)
    tf = textbox(s, M + 0.35, 2.15, SW / 2 - M - 1.0, 2.6)
    para(tf, f"PART {num}", 13, GOLD, bold=True, font=HEAD_FONT, first=True,
         space_after=10)
    para(tf, title, 34, WHITE, bold=True, font=HEAD_FONT, space_after=12,
         line_spacing=1.03)
    para(tf, sub, 13, RGBColor(0xC6, 0xDB, 0xD3), space_after=0, line_spacing=1.15)
    if not img:
        rect(s, SW / 2, 0, SW / 2, SH, CREAM)
    notes(s, notes_text)
    return s


def slide_concept(title, sub, plain, classroom, extra=None, accent=GREEN,
                  notes_text=""):
    """The workhorse slide: plain-English meaning + classroom picture."""
    s = new()
    y = title_bar(s, title, sub)
    # plain-English panel
    rect(s, M, y + 0.02, CW, 1.62, GREEN_LT, MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.05
    rect(s, M, y + 0.02, 0.055, 1.62, accent)
    chip(s, M + 0.26, y + 0.20, "Plain English", accent, WHITE)
    tf = textbox(s, M + 0.26, y + 0.63, CW - 0.6, 0.95)
    para(tf, plain, 13.5, INK, first=True, line_spacing=1.16, space_after=0)
    # classroom picture panel
    cy = y + 1.78
    rect(s, M, cy, CW, 2.28, GOLD_LT, MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.05
    rect(s, M, cy, 0.055, 2.28, GOLD)
    chip(s, M + 0.26, cy + 0.18, "Classroom picture", TERRA, WHITE, size=10)
    tf = textbox(s, M + 0.26, cy + 0.61, CW - 0.6, 1.6)
    para(tf, classroom, 13.5, INK, first=True, line_spacing=1.16, space_after=0)
    if extra:
        tf = textbox(s, M, cy + 2.42, CW, 0.62)
        rich(tf, extra, 11.5, first=True, line_spacing=1.12)
    footer(s)
    notes(s, notes_text)
    return s


def slide_image_text(title, sub, img, bullets, img_side="right", accent=GREEN,
                     note_text="", img_caption=None):
    s = new()
    y = title_bar(s, title, sub)
    img_w = 5.35
    txt_w = CW - img_w - 0.42
    if img_side == "right":
        il, tl = M + txt_w + 0.42, M
    else:
        il, tl = M, M + img_w + 0.42
    picture_cover(s, img, il, y + 0.05, img_w, 5.10)
    rect(s, il, y + 0.05, img_w, 0.065, accent)
    if img_caption:
        tf = textbox(s, il, y + 5.22, img_w, 0.4)
        para(tf, img_caption, 10.5, GREY, italic=True, first=True, space_after=0,
             align=PP_ALIGN.CENTER, line_spacing=1.05)
    layout_cards(s, bullets, tl, y + 0.06, txt_w, accent, footer_cb=footer)
    notes(s, note_text)
    return s


def _lines_needed(text, pt, body_w_pt):
    """Estimate wrapped line count the same way the renderer will."""
    cpl = max(8, body_w_pt / (0.485 * pt))
    n = 0
    for chunk in text.split('\n'):
        n += max(1, -(-len(chunk) // int(cpl)))
    return n


def _measure(bullets, txt_w, tsize, bsize):
    """Height each card needs, in inches (incl. its internal padding)."""
    inner_w = (txt_w - 0.50) * 72.0
    out = []
    for b in bullets:
        lt = _lines_needed(b[0], tsize, inner_w)
        lb = _lines_needed(b[1], bsize, inner_w)
        out.append(lt * (tsize * 1.298 / 72.0) + lb * (bsize * 1.298 / 72.0)
                   + 3.0 / 72.0 + 0.30)
    return out


def layout_cards(slide, bullets, l, t, w, accent, footer_cb=None):
    """Stack cards down a column, shrinking type until everything fits."""
    track = 6.95 - t
    gap = 0.12
    n = len(bullets)
    avail = track - gap * (n - 1)
    for tsize, bsize in ((12.5, 11.5), (12, 11), (11.5, 10.5), (11, 10),
                         (10.5, 9.5), (10, 9)):
        needs = _measure(bullets, w, tsize, bsize)
        if sum(needs) <= avail:
            break
    total = sum(needs)
    if total <= avail:
        slack = (avail - total) / n
        heights = [x + slack for x in needs]
    else:
        heights = [x * avail / total for x in needs]
    by = t
    for b, ch in zip(bullets, heights):
        rect(slide, l, by, w, ch, WHITE, MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.07
        rect(slide, l, by, 0.05, ch, accent)
        tf = textbox(slide, l + 0.22, by + 0.15, w - 0.44, ch - 0.30)
        para(tf, b[0], tsize, accent, bold=True, font=HEAD_FONT, first=True, space_after=3)
        para(tf, b[1], bsize, INK, space_after=0, line_spacing=1.1)
        by += ch + gap
    if footer_cb:
        footer_cb(slide)


def diagram_flow(slide, l, t, w, h, steps, accent=GREEN, caption=None):
    """A vertical 'flow' diagram drawn from native shapes — fully editable in PowerPoint."""
    cap_h = 0.42 if caption else 0.0
    n = len(steps)
    arrow_h = 0.26
    avail = h - cap_h - arrow_h * (n - 1)
    step_h = avail / n
    rect(slide, l, t, w, h - cap_h, GREEN_LT, MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.05
    y = t + 0.16
    inner = w - 0.5
    for i, (label, sub) in enumerate(steps):
        rect(slide, l + 0.24, y, inner, step_h - 0.10, WHITE,
             MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.10
        rect(slide, l + 0.24, y, 0.07, step_h - 0.10, accent)
        num = rect(slide, l + 0.44, y + (step_h - 0.10 - 0.36) / 2, 0.36, 0.36,
                   accent, MSO_SHAPE.OVAL)
        tfn = num.text_frame
        tfn.vertical_anchor = MSO_ANCHOR.MIDDLE
        para(tfn, str(i + 1), 12, WHITE, bold=True, font=HEAD_FONT, first=True,
             align=PP_ALIGN.CENTER, space_after=0)
        tf = textbox(slide, l + 0.94, y + 0.10, inner - 0.82, step_h - 0.30)
        para(tf, label, 12.5, accent, bold=True, font=HEAD_FONT, first=True,
             space_after=2, line_spacing=1.02)
        para(tf, sub, 11, INK, space_after=0, line_spacing=1.08)
        y += step_h - 0.10
        if i < n - 1:
            rect(slide, l + w / 2 - 0.115, y + 0.02, 0.23, arrow_h - 0.04,
                 accent, MSO_SHAPE.DOWN_ARROW)
            y += arrow_h
    if caption:
        tf = textbox(slide, l, t + h - cap_h + 0.05, w, cap_h - 0.05)
        para(tf, caption, 10.5, GREY, italic=True, first=True, space_after=0,
             align=PP_ALIGN.CENTER, line_spacing=1.05)


def slide_diagram_text(title, sub, steps, bullets, diagram_side="right",
                       accent=GREEN, note_text="", caption=None):
    """Same as slide_image_text but the visual is a native, editable diagram."""
    s = new()
    y = title_bar(s, title, sub)
    dw = 5.35
    txt_w = CW - dw - 0.42
    if diagram_side == "right":
        dl, tl = M + txt_w + 0.42, M
    else:
        dl, tl = M, M + dw + 0.42
    diagram_flow(s, dl, y + 0.05, dw, 5.10, steps, accent=accent, caption=caption)
    layout_cards(s, bullets, tl, y + 0.06, txt_w, accent, footer_cb=footer)
    notes(s, note_text)
    return s


def slide_grid_cards(title, sub, img, bullets, accent=GREEN, note_text="",
                     img_caption=None, cols=2):
    """Image on the left, cards in a 2-column grid on the right."""
    s = new()
    y = title_bar(s, title, sub)
    img_w = 4.30
    picture_cover(s, img, M, y + 0.05, img_w, 5.10)
    rect(s, M, y + 0.05, img_w, 0.065, accent)
    if img_caption:
        tf = textbox(s, M, y + 5.22, img_w, 0.4)
        para(tf, img_caption, 10.5, GREY, italic=True, first=True, space_after=0,
             align=PP_ALIGN.CENTER, line_spacing=1.05)
    gl = M + img_w + 0.40
    gw = SW - M - gl
    colw = (gw - 0.22) / cols
    rows = -(-len(bullets) // cols)
    track = 6.95 - y - 0.06
    gap = 0.12
    ch = (track - gap * (rows - 1)) / rows
    for i, b in enumerate(bullets):
        r, c = divmod(i, cols)
        cl = gl + c * (colw + 0.22)
        ct = y + 0.06 + r * (ch + gap)
        rect(s, cl, ct, colw, ch, WHITE, MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.07
        rect(s, cl, ct, colw, 0.055, accent)
        tf = textbox(s, cl + 0.22, ct + 0.20, colw - 0.44, ch - 0.36)
        para(tf, b[0], 13, accent, bold=True, font=HEAD_FONT, first=True, space_after=4)
        para(tf, b[1], 11.5, INK, space_after=0, line_spacing=1.1)
    footer(s)
    notes(s, note_text)
    return s
    notes(s, note_text)
    return s


def slide_full_image(img, kicker, title, caption, note_text=""):
    s = new()
    picture_cover(s, img, 0, 0, SW, SH)
    rect(s, 0, SH - 2.55, SW, 2.55, GREEN_DK, MSO_SHAPE.RECTANGLE)
    from pptx.oxml.ns import qn
    ov = s.shapes[-1]
    solid_el = ov.fill._xPr.find(qn('a:solidFill'))
    clr = solid_el.find(qn('a:srgbClr'))
    clr.append(clr.makeelement(qn('a:alpha'), {'val': '90000'}))
    rect(s, 0, 0, SW, 0.13, GOLD)
    tf = textbox(s, M, SH - 2.30, CW, 2.0)
    para(tf, kicker, 12, GOLD, bold=True, font=HEAD_FONT, first=True, space_after=8)
    para(tf, title, 32, WHITE, bold=True, font=HEAD_FONT, space_after=10,
         line_spacing=1.0)
    para(tf, caption, 14.5, RGBColor(0xD8, 0xE8, 0xE1), space_after=0,
         line_spacing=1.12)
    notes(s, note_text)
    return s


def slide_table(title, sub, headers, rows, col_w=None, intro=None, note_text="",
                row_h=0.46, fsize=11.5, callout=None, callout_color=GREEN_LT,
                callout_accent=GREEN, callout_title=None):
    s = new()
    y = title_bar(s, title, sub)
    if intro:
        tf = textbox(s, M, y + 0.02, CW, 0.42)
        para(tf, intro, 12, GREY, italic=True, first=True, space_after=0,
             line_spacing=1.1)
        y += 0.48
    end = table(s, M, y + 0.06, CW, headers, rows, col_w=col_w, row_h=row_h,
                fsize=fsize)
    if callout:
        cy = end + 0.34
        ch = 7.02 - cy
        if ch > 0.55:
            rect(s, M, cy, CW, ch, callout_color,
                 MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.07
            rect(s, M, cy, 0.055, ch, callout_accent)
            tf = textbox(s, M + 0.28, cy + 0.17, CW - 0.56, ch - 0.34)
            if callout_title:
                para(tf, callout_title, 11, callout_accent, bold=True,
                     font=HEAD_FONT, first=True, space_after=5)
                para(tf, callout, 12, INK, space_after=0, line_spacing=1.14)
            else:
                para(tf, callout, 12, INK, first=True, space_after=0,
                     line_spacing=1.14)
    footer(s)
    notes(s, note_text)
    return s


# ================================================================ DECK SEQUENCE
# Structured to match the KATON 2026 workshop agenda exactly:
#   Welcome 5 · Icebreaker 5 · S1 20 · S2 15 · Break 15 · S3 15 · S4 20
#   Practical 25 · Gallery Walk 10 · Data 10 · Challenges 5 · Action & Close 5


def slide_session(num, title, mins, sub, img=None, notes_text=""):
    """Session divider carrying its agenda timing."""
    s = new()
    if img:
        picture_cover(s, img, SW / 2, 0, SW / 2, SH)
    rect(s, 0, 0, SW / 2, SH, GREEN_DK)
    rect(s, 0, 0, 0.14, SH, GOLD)
    tf = textbox(s, M + 0.35, 1.85, SW / 2 - M - 1.0, 3.2)
    kicker = num if isinstance(num, str) else f"SESSION {num}"
    para(tf, kicker, 13, GOLD, bold=True, font=HEAD_FONT, first=True,
         space_after=6)
    badge = rect(s, M + 0.35, 2.28, 1.30, 0.40, GOLD, MSO_SHAPE.ROUNDED_RECTANGLE)
    badge.adjustments[0] = 0.4
    tfb = badge.text_frame
    tfb.vertical_anchor = MSO_ANCHOR.MIDDLE
    para(tfb, f"{mins} MINUTES", 11, GREEN_DK, bold=True, font=HEAD_FONT, first=True,
         align=PP_ALIGN.CENTER, space_after=0)
    tf2 = textbox(s, M + 0.35, 2.95, SW / 2 - M - 1.0, 2.0)
    para(tf2, title, 32, WHITE, bold=True, font=HEAD_FONT, first=True, space_after=12,
         line_spacing=1.03)
    para(tf2, sub, 13, RGBColor(0xC6, 0xDB, 0xD3), space_after=0, line_spacing=1.15)
    if not img:
        rect(s, SW / 2, 0, SW / 2, SH, CREAM)
    notes(s, notes_text)
    return s


def slide_activity(title, sub, mins, steps, img=None, note_text="", accent=GOLD,
                   brief=None):
    """A workshop activity brief: what to do, in how long."""
    s = new()
    y = title_bar(s, title, sub)
    if img:
        img_w = 4.05
        picture_cover(s, img, M, y + 0.05, img_w, 5.10)
        rect(s, M, y + 0.05, img_w, 0.065, accent)
        tl = M + img_w + 0.42
        tw = CW - img_w - 0.42
    else:
        tl, tw = M, CW
    # timing badge
    b = rect(s, tl, y + 0.02, 1.85, 0.44, accent, MSO_SHAPE.ROUNDED_RECTANGLE)
    b.adjustments[0] = 0.4
    tfb = b.text_frame
    tfb.vertical_anchor = MSO_ANCHOR.MIDDLE
    para(tfb, f"{mins}", 12.5, GREEN_DK, bold=True, font=HEAD_FONT, first=True,
         align=PP_ALIGN.CENTER, space_after=0)
    by = y + 0.60
    if brief:
        tf = textbox(s, tl, by, tw, 0.72)
        para(tf, brief, 12.5, INK, first=True, line_spacing=1.16, space_after=0)
        by += 0.80
    n = len(steps)
    avail = 6.98 - by
    gap = 0.09
    ch = (avail - gap * (n - 1)) / n
    for i, (label, body) in enumerate(steps):
        ct = by + i * (ch + gap)
        rect(s, tl, ct, tw, ch, WHITE, MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.09
        rect(s, tl, ct, 0.055, ch, accent)
        num = rect(s, tl + 0.20, ct + (ch - 0.34) / 2, 0.34, 0.34, accent,
                   MSO_SHAPE.OVAL)
        tfn = num.text_frame
        tfn.vertical_anchor = MSO_ANCHOR.MIDDLE
        para(tfn, str(i + 1), 11.5, GREY if accent is GOLD else WHITE, bold=True,
             font=HEAD_FONT, first=True, align=PP_ALIGN.CENTER, space_after=0)
        tf = textbox(s, tl + 0.68, ct + 0.10, tw - 0.92, ch - 0.20)
        para(tf, label, 12.5, accent, bold=True, font=HEAD_FONT, first=True,
             space_after=3, line_spacing=1.02)
        para(tf, body, 11.5, INK, space_after=0, line_spacing=1.08)
    footer(s)
    notes(s, note_text)
    return s


def slide_solutions(title, sub, rows, intro=None, note_text="", row_h=0.60,
                    fsize=11.5):
    """Challenge | What it looks like | What actually works."""
    s = new()
    y = title_bar(s, title, sub)
    if intro:
        tf = textbox(s, M, y + 0.02, CW, 0.40)
        para(tf, intro, 12, GREY, italic=True, first=True, space_after=0,
             line_spacing=1.10)
        y += 0.46
    headers = ["The challenge", "What it looks like in school", "What actually works"]
    col_w = [3.35, 3.95, 4.79]
    table(s, M, y + 0.06, CW, headers, rows, col_w=col_w, row_h=row_h, fsize=fsize)
    footer(s)
    notes(s, note_text)
    return s


def slide_agenda():
    s = new()
    y = title_bar(s, "Workshop Agenda", "Two hours thirty minutes \u00b7 with a break")
    rows = [
        ["00:00", "Welcome & Workshop Objectives", "5 min", GREEN],
        ["00:05", "Icebreaker \u2014 One Word for Assessment", "5 min", GREEN],
        ["00:10", "Session 1: Understanding Assessment \u2014 Purpose & Types", "20 min", GOLD],
        ["00:30", "Session 2: Ghana's Assessment Policy Landscape", "15 min", GOLD],
        ["00:45", "Break", "15 min", GREY],
        ["01:00", "Session 3: Principles of Effective Assessment", "15 min", GOLD],
        ["01:15", "Session 4: Assessment Strategies Across Levels", "20 min", GOLD],
        ["01:35", "Practical Workshop: Design a Task & Rubric", "25 min", TERRA],
        ["02:00", "Gallery Walk & Group Share-Outs", "10 min", TERRA],
        ["02:10", "Using Assessment Data: Feedback & Records", "10 min", GREEN],
        ["02:20", "Common Challenges & Practical Solutions", "5 min", GREEN],
        ["02:25", "Action Planning & Closing", "5 min", GREEN],
    ]
    ty = y + 0.04
    rh = 0.415
    for i, (t, label, dur, col) in enumerate(rows):
        bg = WHITE if i % 2 == 0 else PAPER
        if label == "Break":
            bg = RGBColor(0xEE, 0xF2, 0xF1)
        rect(s, M, ty, CW, rh, bg)
        rect(s, M, ty, 0.055, rh, col)
        tf = textbox(s, M + 0.24, ty + 0.09, 1.20, rh - 0.14)
        para(tf, t, 11.5, GREY, bold=True, font=HEAD_FONT, first=True, space_after=0)
        tf = textbox(s, M + 1.50, ty + 0.09, 8.20, rh - 0.14)
        para(tf, label, 12.5, INK, bold=(label == "Break" or label.startswith("Session")),
             font=HEAD_FONT if label.startswith("Session") or label == "Break" else BODY_FONT,
             first=True, space_after=0)
        tf = textbox(s, SW - M - 1.55, ty + 0.09, 1.55, rh - 0.14)
        para(tf, dur, 11.5, col, bold=True, font=HEAD_FONT, first=True,
             align=PP_ALIGN.RIGHT, space_after=0)
        ty += rh
    footer(s)
    notes(s, "THE RUNNING ORDER (1 min)\n\n"
             "Walk the room through the shape of the day quickly. The two things worth flagging:\n\n"
             "1. THE PRACTICAL WORKSHOP (1:35) is the heart of the day - 25 minutes for you to "
             "actually design a task and a rubric you can use. Everything before it is preparation "
             "for that.\n\n"
             "2. THE BREAK IS AT 00:45, after Session 2. Have tea and water ready; people lose "
             "concentration if the break slips.\n\n"
             "Then move straight into the objectives - do not dwell on the agenda itself.")
    return s


def slide_objectives():
    s = new()
    y = title_bar(s, "By the End of This Workshop, You Will Be Able To\u2026",
                  "Four things. That is the whole promise.")
    items = [
        ("1", "Distinguish & Apply",
         "Distinguish formative and summative assessment, and use each purposefully at the "
         "right point in a lesson or term.", GREEN),
        ("2", "Navigate Policy",
         "Explain Ghana's continuous assessment requirements, and locate your grade level "
         "within the Standards-Based Curriculum.", GOLD),
        ("3", "Design Fair Tools",
         "Create valid, fair and inclusive assessment tasks and simple rubrics suited to your "
         "learners' level.", TERRA),
        ("4", "Use Data for Learning",
         "Turn assessment results into targeted feedback, better teaching decisions, and clear "
         "records.", GREEN),
    ]
    x, w, gap = M, 2.95, 0.22
    for num, head, body, col in items:
        rect(s, x, y + 0.18, w, 3.55, WHITE, MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.05
        rect(s, x, y + 0.18, w, 0.075, col)
        c = rect(s, x + 0.30, y + 0.52, 0.60, 0.60, col, MSO_SHAPE.OVAL)
        tfc = c.text_frame
        tfc.vertical_anchor = MSO_ANCHOR.MIDDLE
        para(tfc, num, 21, WHITE, bold=True, font=HEAD_FONT, first=True,
             align=PP_ALIGN.CENTER, space_after=0)
        tf = textbox(s, x + 0.30, y + 1.28, w - 0.60, 2.25)
        para(tf, head, 14.5, col, bold=True, font=HEAD_FONT, first=True, space_after=7,
             line_spacing=1.02)
        para(tf, body, 11.5, INK, space_after=0, line_spacing=1.14)
        x += w + gap
    tf = textbox(s, M, y + 3.98, CW, 0.60)
    rich(tf, [("Note the verb in objective 3:  ", GREEN, True, False),
              ("\u201cdesign\u201d, not \u201cdiscuss\u201d. By 1:35 this afternoon you will have "
               "built something \u2014 not just talked about it.", INK, False, False)],
         12.5, first=True, line_spacing=1.15)
    footer(s)
    notes(s, "OBJECTIVES (4 min)\n\n"
             "Read each objective aloud, then translate it into teacher-speak:\n\n"
             "1. DISTINGUISH & APPLY - 'You will know which kind of assessment to reach for, and "
             "when. Not just the names.'\n"
             "2. NAVIGATE POLICY - 'You will know what Ghana actually requires of you, and where "
             "your subject sits. No more guessing.'\n"
             "3. DESIGN FAIR TOOLS - 'You will build a task and a rubric today. Actually build it, "
             "not plan to build it.'\n"
             "4. USE DATA FOR LEARNING - 'Marks will stop being a record of the past and start "
             "informing what you teach next.'\n\n"
             "POINT AT THE NOTE AT THE BOTTOM. Objective 3 says DESIGN. That is deliberate and it "
             "is the promise of the afternoon. Flag it now so people know the workshop is not "
             "going to be all talk.\n\n"
             "Then ask for a show of hands: 'Who has ever been to a workshop where you designed "
             "something you actually used the following week?' Usually very few hands. That is "
             "what we are trying to change today.")
    return s


def slide_icebreaker():
    s = new()
    y = title_bar(s, "Icebreaker \u2014 One Word for Assessment",
                  "Five minutes \u00b7 honesty welcome, formality not required")
    img_w = 5.10
    picture_cover(s, "18_icebreaker.jpg", M, y + 0.05, img_w, 5.10)
    rect(s, M, y + 0.05, img_w, 0.065, GOLD)
    tl = M + img_w + 0.42
    tw = CW - img_w - 0.42
    b = rect(s, tl, y + 0.04, 1.55, 0.42, GOLD, MSO_SHAPE.ROUNDED_RECTANGLE)
    b.adjustments[0] = 0.4
    tfb = b.text_frame
    tfb.vertical_anchor = MSO_ANCHOR.MIDDLE
    para(tfb, "5 MINUTES", 11.5, GREEN_DK, bold=True, font=HEAD_FONT, first=True,
         align=PP_ALIGN.CENTER, space_after=0)
    steps = [
        ("Think", "When you hear the word \u201cassessment\u201d, what is the FIRST word that "
                  "comes into your mind? Not the textbook word. Your honest word."),
        ("Share", "Turn to the person beside you. One word each. Thirty seconds. No explaining, "
                  "no defending \u2014 just the word."),
        ("Collect", "We will hear a few out loud. Every word is welcome \u2014 including the "
                    "uncomfortable ones."),
    ]
    by = y + 0.62
    ch = 1.20
    for i, (label, body) in enumerate(steps):
        ct = by + i * (ch + 0.10)
        rect(s, tl, ct, tw, ch, WHITE, MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.08
        rect(s, tl, ct, 0.055, ch, GOLD)
        tf = textbox(s, tl + 0.24, ct + 0.14, tw - 0.48, ch - 0.28)
        para(tf, label, 12.5, TERRA, bold=True, font=HEAD_FONT, first=True, space_after=3)
        para(tf, body, 11.5, INK, space_after=0, line_spacing=1.10)
    # debrief card
    dy = by + 3 * (ch + 0.10) - 0.10
    rect(s, tl, dy, tw, 1.72, GREEN_LT, MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.06
    rect(s, tl, dy, 0.055, 1.72, GREEN)
    tf = textbox(s, tl + 0.24, dy + 0.16, tw - 0.48, 1.44)
    para(tf, "WHY WE START HERE", 10.5, GREEN, bold=True, font=HEAD_FONT, first=True,
         space_after=5)
    para(tf, "You will hear words like stress, marks, fear, pressure, paperwork, ranking. "
             "You may also hear growth, feedback, help, understanding. Both sets are honest "
             "\u2014 and the difference between them is the whole point of today.",
         11, INK, space_after=0, line_spacing=1.12)
    footer(s)
    notes(s, "ICEBREAKER (5 min) - this sets the emotional tone for the whole day\n\n"
             "DO NOT SKIP THIS. It looks like a warm-up; it is actually the most important five "
             "minutes of the workshop, because it surfaces how teachers really feel about "
             "assessment before you start teaching anything.\n\n"
             "HOW TO RUN IT:\n"
             "  - Give 30 seconds of silent thinking first. Do not let people answer immediately.\n"
             "  - Then 60 seconds in pairs. Keep it strictly to one word each - the constraint is "
             "what makes it work.\n"
             "  - Then take 6-8 words from the floor. Write them up on the board or flip chart, "
             "uncommented, in two loose columns as they fall.\n\n"
             "WHAT YOU WILL HEAR, almost certainly: stress, marks, fear, pressure, exams, "
             "paperwork, ranking, punishment, workload. Occasionally: feedback, growth, help, "
             "understanding, progress.\n\n"
             "THE DEBRIEF - deliver this gently, without judgement:\n"
             "'Look at what we just wrote. Almost none of these words are about learning. They are "
             "about being judged. And that is not because we are bad teachers - it is because most "
             "of us were assessed that way ourselves, and we inherited it. By the end of today I "
             "want to move some of these words across to the other side.'\n\n"
             "DO NOT let anyone feel accused. Frame it as something we inherited, not something we "
             "chose. Then move on briskly - do not let the icebreaker run past five minutes.")
    return s


# ---------------------------------------------------------------- WELCOME
slide_title()
slide_agenda()
slide_objectives()
slide_icebreaker()


# ---------------------------------------------------------------- SESSION 1
slide_session(1, "Understanding Assessment",
              20,
              "Purpose & types \u2014 what assessment is for, and the family of tools.",
              img="07_feedback_loop.jpg",
              notes_text="SESSION 1 (20 min)\n\n"
              "This session answers two questions: what is assessment actually FOR, and what "
              "kinds are there?\n\n"
              "We start with the definition, then the purposes, then the big split between "
              "formative and summative using the soup metaphor. We finish with the full family of "
              "assessment types and a quick classification exercise.\n\n"
              "PACE NOTE: this is 20 minutes and it is the foundation for everything else. Do not "
              "rush the soup metaphor - it will be referenced all day.")

slide_concept(
    "Assessment \u2014 The Plain Meaning",
    "The single most important definition of the day",
    "Finding out what a learner knows and can do, and how far they have travelled towards the "
    "goal you set. It is gathering information about a learner in order to make a decision "
    "about their learning.",
    "It is not only the end-of-term paper. It is the question you asked in the middle of the "
    "lesson and the blank faces that answered it. It is the exercise book you marked at 9pm. "
    "It is watching a group argue over a titration and realising they have misunderstood the "
    "indicator. Every one of those is assessment \u2014 and every one of them changed what you "
    "did next.",
    extra=[("The official wording says the same thing:  ", INK, False, False),
           ("\u201cthe process of collecting information on the learner to help in deciding the "
            "degree to which the learner has achieved the expected learning outcomes/standards.\u201d",
            GREY, False, True)],
    notes_text="THE CORE DEFINITION (3 min)\n\n"
    "Read the plain version aloud. Then ask the room to name one thing they did last week that "
    "was assessment but NOT a test. Collect two or three answers.\n\n"
    "KEY POINT to land: assessment is not a document, it is a DECISION-MAKING process. If you "
    "gathered information and changed nothing, you collected data but you did not really assess.\n\n"
    "Note the careful wording in the official definition - the goal is not to rank learners, but "
    "to decide 'the degree to which' they have achieved a standard. That phrase matters, and it "
    "links directly to objective 1 of today's workshop.")

slide_image_text(
    "Why We Assess \u2014 The Purposes",
    "Assessment is not one thing doing one job",
    "04_pantry_market.jpg",
    [("To improve learning", "The main purpose. Information gathered DURING learning, used to adjust teaching while there is still time to help."),
     ("To diagnose", "Before teaching, to find out what learners already bring and where the gaps are."),
     ("To inform teaching", "Your planning for next week should be shaped by what last week's assessment revealed."),
     ("To certify and select", "At the end, to summarise achievement for transcripts, placement and employers. A legitimate purpose \u2014 just not the only one."),
     ("To motivate and involve", "Assessment that learners understand and act on builds ownership. Assessment that only judges breeds fear \u2014 as the icebreaker showed.")],
    img_side="right",
    accent=GREEN,
    img_caption="Checking what learners already bring, before you build on it.",
    note_text="THE PURPOSES (3 min)\n\n"
             "THE REFRAME: most teachers experience assessment as one thing - judging. The manual "
             "gives it at least five distinct jobs, and judging is only one of them.\n\n"
             "WALK THE LIST and ask for a show of hands on each: 'Who assessed to IMPROVE learning "
             "last week?' (most hands) 'Who assessed to CERTIFY?' (all hands) 'Who assessed to "
             "DIAGNOSE before starting a topic?' (usually far fewer).\n\n"
             "LINK BACK TO THE ICEBREAKER: 'The words we wrote up earlier - stress, fear, pressure "
             "- come from assessment being used for only one of these five purposes. When it is "
             "used for the first three, it stops feeling like a threat.'\n\n"
             "IMPORTANT BALANCE: do not let this sound like summative assessment is the villain. "
             "Certification and selection are legitimate and necessary. The problem is summative "
             "assessment being the ONLY kind learners experience.")

slide_full_image(
    "02_soup_tasting.jpg",
    "FORMATIVE ASSESSMENT  \u00b7  ASSESSMENT FOR LEARNING + AS LEARNING",
    "Tasting the Soup",
    "You taste while it is still on the fire. No salt? Add some. Too much pepper? Add water. "
    "You have not failed the soup \u2014 you have just improved it before anyone sits down to "
    "eat. That is assessment FOR learning, and it saves the meal.",
    note_text="THE SOUP (3 min)\n\n"
    "Ask the room: 'Does anybody here wait until the food is served at the table before they "
    "first taste it?' Wait for the laughter and the shaking heads. Of course not.\n\n"
    "Then the turn: 'So why do we do exactly that with our learners? Why do we find out what "
    "they misunderstood in the end-of-term exam, when it is too late to fix it?'\n\n"
    "KEY POINT: the cook who tastes is not being nosy or wasting time. Tasting is not an "
    "interruption to cooking - tasting IS cooking.\n\n"
    "The manual's words: formative assessment 'enables teachers to modify or improve teaching "
    "and learning'. Notice the word IMPROVE. Not PROVE.\n\n"
    "Hold this image up for the rest of the day. Refer back to it in Session 3 and in the "
    "practical workshop.")

slide_full_image(
    "03_meal_served.jpg",
    "SUMMATIVE ASSESSMENT  \u00b7  ASSESSMENT OF LEARNING",
    "Serving the Meal",
    "The cooking is finished. This is the meal as it is. No amount of tasting will help now. "
    "You serve it, the family eats, and the meal is judged as served. That is assessment OF "
    "learning \u2014 end of term, WASSCE, certification.",
    note_text="THE MEAL (3 min)\n\n"
    "Summative assessment is not the enemy. It has a real and necessary job: selection, "
    "certification and placement. University admissions and employers need a final verdict. "
    "The manual says it provides 'a summary of the learner's overall achievement for selection, "
    "certification and placement purposes.'\n\n"
    "The point is not that summative is bad. The point is that summative ALONE is a poor diet. "
    "If you only ever serve meals and never taste, you will keep serving the same mistakes.\n\n"
    "Examples to name: end-of-term exams, WASSCE, end-of-programme exams, term papers, "
    "industrial attachment assessment.\n\n"
    "NOTE FOR SESSION 2: this is where the 70/30 split becomes relevant - the summative "
    "judgement is partly YOURS. Flag it and move on.")

slide_table(
    "The Three Sisters: AfL, AaL and AoL",
    "These three acronyms cause the most confusion \u2014 here they are side by side",
    ["Acronym", "Full name", "Who is doing the work", "In one plain sentence"],
    [["AfL", "Assessment for Learning", "The TEACHER", "I check as we go, so I can teach better."],
     ["AaL", "Assessment as Learning", "The LEARNER", "The student checks themselves, so they learn better."],
     ["AoL", "Assessment of Learning", "The SYSTEM / teacher", "A final summary of what has been achieved."]],
    col_w=[1.35, 2.85, 2.35, 5.54],
    intro="AfL and AaL together make up what the manual calls FORMATIVE assessment. AoL is SUMMATIVE.",
    row_h=0.62,
    callout_title="A SIMPLE TEST TO USE ON YOUR OWN PRACTICE",
    callout="Ask: \u201cWho is going to change what they do as a result of this information?\u201d "
            "If the answer is the TEACHER, it is AfL. If it is the LEARNER, it is AaL. "
            "If nobody changes anything and it just gets recorded, it is AoL.\n"
            "AaL is the one we most often forget: we tell learners their marks, but we rarely "
            "teach them to check themselves \u2014 and that is the skill that carries them into "
            "university and into work.",
    note_text="THE THREE SISTERS (3 min)\n\n"
    "This table is your anchor for the rest of the workshop. Refer back to it constantly.\n\n"
    "Say it this way: 'For' and 'as' are both about improving. 'Of' is about reporting.\n\n"
    "A simple test to give teachers - ask: 'Who is going to change what they do as a result of "
    "this information?' If the answer is the teacher, it is AfL. If it is the learner, it is AaL. "
    "If nobody changes anything and it just gets recorded, it is AoL.\n\n"
    "IMPORTANT: AaL is the one we forget. We tell them their marks, but we never teach them to "
    "check themselves. That is the skill that carries them into university and work.\n\n"
    "WHERE AaL LIVES IN PRACTICE: self-assessment against a rubric, reflection journals, "
    "learners tracking their own progress. We will build this into the rubric exercise later.")

slide_table(
    "Formative vs Summative \u2014 Side by Side",
    "The same skills, two different jobs",
    ["", "FORMATIVE (the tasting)", "SUMMATIVE (the serving)"],
    [["When", "During teaching \u2014 before, during and after the lesson", "After teaching \u2014 end of unit, term, year, programme"],
     ["Purpose", "To IMPROVE learning while there is still time", "To PROVE and record overall achievement"],
     ["Who uses it", "Teacher and learner", "School, WAEC and other external bodies"],
     ["What it sounds like", "\u201cI see where you got stuck \u2014 let us try it this way\u201d", "\u201cThis is your final grade for the term\u201d"],
     ["Feels like", "A conversation, a question, a quick check, a show of hands", "A formal paper, a scheduled exam, a transcript"],
     ["Examples", "Questions, exit cards, class exercises, observation, peer assessment, drafts", "End-of-term exam, WASSCE, projects, portfolios, practicals"]],
    col_w=[1.85, 4.90, 5.34],
    row_h=0.56,
    note_text="SIDE BY SIDE (2 min)\n\n"
             "Walk down the rows. The 'What it sounds like' row is the one that lands with "
             "teachers - read both sentences out loud in the two different tones of voice.\n\n"
             "Ask: 'Which of these two voices do our learners hear most from us?' Be honest that "
             "for many of us, it is the second voice far more often than the first.\n\n"
             "Reassure: this is not about working harder. It is about redistributing the same "
             "effort. A ten-minute exit card can replace an hour of re-teaching the wrong thing.")

slide_concept(
    "The Golden Rule Most Teachers Miss",
    "The two do not live in separate boxes \u2014 they feed each other",
    "Summative results can be used FORMATIVELY, and formative work can count SUMMATIVELY. "
    "The label describes the PURPOSE the information is put to \u2014 not the test paper itself.",
    "You mark the end-of-term paper as usual. But instead of only recording the scores, you "
    "hand the scripts back, put learners in groups to find the answers to the questions they "
    "missed, and you re-teach the two topics the whole class failed. That exam just became a "
    "teaching tool. Equally: the class exercises and the group project you grade against a "
    "clear rubric \u2014 that is the 30% school-based assessment. Your everyday work is already "
    "counting towards the final grade.",
    extra=[("Why this matters for you:  ", TERRA, True, False),
           ("the same piece of work can do two jobs. You are not being asked to double your "
            "marking load \u2014 you are being asked to use the information twice.", GREY, False, False)],
    accent=TERRA,
    notes_text="THE GOLDEN RULE (3 min)\n\n"
    "This is the highest-value concept in Session 1. Many teachers believe formative and "
    "summative are two separate systems competing for time. They are not.\n\n"
    "The manual has a whole section on this (3.5.3 and 3.5.4). Read the example aloud if you "
    "have the manual open: 'instead of scoring a summative test, the teacher gives back the test "
    "to learners to discuss in groups to find answers to the questions that were wrongly "
    "answered.'\n\n"
    "PRACTICAL COMMITMENT TO ASK FOR: challenge every teacher to pick ONE end-of-term paper this "
    "year and give it back for group correction instead of just recording the marks. One paper. "
    "We will return to this in Action Planning.\n\n"
    "BRIDGE TO SESSION 2: 'This is also where policy becomes personal. That 30% I just mentioned "
    "- that is you. Let us look at exactly what Ghana requires.'")

slide_diagram(
    "cycle.png",
    "Assessment Is a Cycle, Not an Event",
    "The loop that turns marks into learning \u2014 and where it usually breaks",
    "Most of us run PLAN \u2192 TEACH \u2192 ASSESS and then stop. The learning happens in the last two steps.",
    note_text="THE CYCLE (3 min)\n\n"
    "WHY THIS SLIDE EXISTS: teachers often experience assessment as something that happens TO "
    "them at the end of term - a deadline, a stack of scripts, a set of marks to enter. This "
    "diagram reframes it as a loop the teacher controls.\n\n"
    "TRACE IT WITH YOUR FINGER on the screen, or better, walk it around the room. Five steps, "
    "and the arrow always comes back to PLAN.\n\n"
    "THE KEY INSIGHT, and the one to say slowly: 'Most schools do the first three steps and "
    "stop. You plan, you teach, you assess - and then the marks go in the book and nothing "
    "changes. That is not assessment. That is record-keeping.'\n\n"
    "ANALYSE is the step that is almost always skipped. Ask the room: 'When did you last sit "
    "down with a set of results and look for a PATTERN, rather than just a mark?' We come back "
    "to how to do that at 2:10.\n\n"
    "THE TIMING POINT: the gap between ASSESS and ACT should be days. Feedback that arrives "
    "three weeks later is history, not feedback.\n\n"
    "LINK FORWARD: everything in Session 3 (the seven pillars) exists to make this loop "
    "trustworthy enough to act on.")

slide_table(
    "Quick Check \u2014 Formative or Summative?",
    "Two minutes \u00b7 call it out \u00b7 then we discuss",
    ["The situation", "Formative or summative?", "Why"],
    [["You give a five-question quiz halfway through a lesson to check understanding.", "?", "?"],
     ["You set the end-of-term examination paper.", "?", "?"],
     ["You stand beside a group and listen to how they are solving a problem.", "?", "?"],
     ["You mark a learner's draft essay and ask them to rewrite one paragraph.", "?", "?"],
     ["You submit learners' WASSCE registration grades.", "?", "?"],
     ["You grade a group project against a rubric and it counts towards their 30%.", "?", "?"]],
    col_w=[6.80, 2.65, 2.64],
    intro="Do not think about the paperwork. Ask: who is going to change what they do as a result of this information?",
    row_h=0.58,
    note_text="QUICK CHECK (4 min) - the interactive close of Session 1\n\n"
             "RUN IT FAST. Read each situation and take a show of hands or shout-outs. Do not "
             "labour each one for more than 30 seconds.\n\n"
             "THE ANSWERS, with the reasoning you want them to hear:\n"
             "1. Mid-lesson quiz -> FORMATIVE. The purpose is to find out and adjust. Note: an "
             "UNANNOUNCED quiz can still be formative - the label is about use, not formality.\n"
             "2. End-of-term paper -> SUMMATIVE. It sums up a period and is recorded.\n"
             "3. Listening to a group -> FORMATIVE. Observation is one of the manual's named "
             "strategies. This surprises people - it does not need paper to count.\n"
             "4. Marking a draft and asking for a rewrite -> FORMATIVE. The learner acts on it. "
             "Notice this is written work but still formative.\n"
             "5. WASSCE registration grades -> SUMMATIVE. External, recorded, for certification.\n"
             "6. Group project graded against a rubric counting to the 30% -> THE GOLDEN RULE. It "
             "is formative work USED summatively. This is the one that matters most.\n\n"
             "IF THE ROOM IS SPLIT on number 6, that is a good outcome - it means the distinction "
             "between the TOOL and its USE is the thing to dwell on. Say: 'This is exactly the "
             "point. The work is formative while they are doing it and summative once it is "
             "counted. Same work, two jobs.'\n\n"
             "CLOSE SESSION 1 by returning to the soup: 'Taste the soup. Serve the meal. Now - "
             "who says you may taste, and what must you record? That is Session 2.'")


# ---------------------------------------------------------------- SESSION 2
slide_session(2, "Ghana's Assessment\nPolicy Landscape",
              15,
              "What you are required to do \u2014 and who requires it.",
              img="05_pillars.jpg",
              notes_text="SESSION 2 (15 min)\n\n"
              "This session answers: what does Ghana actually require of me, and who decides?\n\n"
              "Teachers often carry out assessment requirements without knowing where they come "
              "from or what they are for. This session connects the dots: the policy actors, the "
              "two frameworks, the 70/30 split, continuous assessment requirements, the "
              "Standards-Based Curriculum, and the school-based transcript.\n\n"
              "PACE NOTE: keep this brisk and concrete. The 70/30 slide is the one that lands - "
              "make sure you reach it with time to spare.")

slide_image_text(
    "Who Does What \u2014 The Policy Actors",
    "Six bodies, one system. Knowing who is who makes circulars make sense.",
    "17_practical_lab.jpg",
    [("NaCCA", "National Council for Curriculum and Assessment. Sets the curriculum AND the assessment guidance. This workshop's manual is theirs."),
     ("Ministry of Education", "Sets national policy and funds the system. Owns the reforms behind the new curriculum."),
     ("Ghana Education Service", "Runs public pre-tertiary schools. Your direct employer; issues day-to-day directives."),
     ("WAEC", "West African Examinations Council. Runs WASSCE \u2014 the external 70%. Also advises on item quality."),
     ("NaSIA", "National School Inspectorate Authority. Inspects schools \u2014 and will look at your assessment records."),
     ("National Teaching Council", "Regulates the teaching profession: licensure, professional standards and continuous professional development.")],
    img_side="right",
    accent=GREEN,
    img_caption="Different bodies, different jobs \u2014 all pointing at your classroom.",
    note_text="THE POLICY ACTORS (3 min)\n\n"
             "WHY THIS SLIDE EXISTS: teachers receive circulars from NaCCA, GES, NaSIA and the NTC "
             "and often cannot tell who is asking for what or why. That confusion is what makes "
             "assessment requirements feel like arbitrary paperwork.\n\n"
             "TEACH IT AS A SIMPLE DIVISION OF LABOUR: NaCCA says WHAT should be learned and how "
             "it should be assessed. The Ministry sets policy. GES runs the schools. WAEC examines "
             "externally. NaSIA checks. The NTC governs your profession.\n\n"
             "THE TWO THAT TOUCH YOU MOST DIRECTLY, and worth naming explicitly:\n"
             "  - GES - because they are your employer and issue operational directives.\n"
             "  - NaSIA - because inspectors will ask to see your assessment records. Make sure "
             "your records are in order.\n\n"
             "REASSURANCE: you are not expected to memorise this. The point is simply that when a "
             "circular arrives, you can tell which body it came from and what kind of instruction "
             "it is.")

slide_table(
    "The Two Frameworks \u2014 NPLAF and SEAF",
    "Where the rules in the manual actually come from",
    ["Framework", "Full name", "What it covers", "Why you should care"],
    [["NPLAF", "National Pre-tertiary Learning and Assessment Framework",
      "The national master document. Defines the principles and forms of assessment for ALL pre-tertiary levels.",
      "The root source. Every principle in Session 3 traces back here."],
     ["SEAF", "Secondary Education Assessment Framework",
      "How SHS / SHTS / STEM learners are assessed within and across grade levels.",
      "The direct rulebook for your classroom. It proposes the two forms: formative and summative."],
     ["SBC", "Standards-Based Curriculum",
      "The current SHS curriculum. Organised by strands and sub-strands with learning outcomes expressed as standards.",
      "Your teaching and assessment must be aligned to its learning outcomes \u2014 that is what validity means in practice."]],
    col_w=[1.35, 3.30, 4.30, 3.14],
    intro="Three acronyms that govern everything you do. You do not need to memorise them \u2014 you need to know which one answers your question.",
    row_h=1.02,
    note_text="THE FRAMEWORKS (4 min)\n\n"
             "TEACH THE RELATIONSHIP, not the definitions. It is a hierarchy:\n"
             "NPLAF is the root - the national framework for all pre-tertiary assessment.\n"
             "SEAF sits under it and it is the one that speaks directly to SHS.\n"
             "The SBC is the curriculum itself, which tells you the learning outcomes you must "
             "assess against.\n\n"
             "THE ONE SENTENCE THAT MATTERS: 'Your assessment must be aligned to the learning "
             "outcomes in the Standards-Based Curriculum.' That single requirement is where "
             "validity comes from - you cannot validly assess a standard you never taught, and you "
             "cannot teach a standard if you have not located it in the curriculum.\n\n"
             "PRACTICAL DIRECTION: tell teachers to locate their subject's learning outcomes in "
             "the SBC for the term ahead. If they have never done this, that is the single most "
             "useful thing they can do after today.\n\n"
             "DO NOT get dragged into a debate about the merits of the reform. If someone raises "
             "it, acknowledge it and note that the manual is the guidance for the current "
             "curriculum, whatever we think of it.")

slide_image_text(
    "The 70 / 30 Split \u2014 Your Marks Now Have Weight",
    "Why school-based assessment is no longer soft data",
    "21_records.jpg",
    [("The split", "Thirty per cent of a learner's final grade comes from work done in YOUR school and marked by YOU. Seventy per cent comes from WAEC's final WASSCE."),
     ("What makes up the 30%", "Formative work (portfolios, performance and project work) plus summative internal work (end-of-term examinations), recorded in a school-based transcript."),
     ("What changed", "Internal marks are no longer treated as a formality at the back of the mark book. The 30% now sits on a school-based transcript with far greater transparency and quality assurance."),
     ("The consequence", "If your school's 30% is not credible, transparent and defensible, the learner is the one who suffers when the transcript is scrutinised \\u2014 and the school is the one that is exposed."),
     ("So what?", "Reliable rubrics applied consistently are no longer just good practice. They are protection \\u2014 for your learners and for your school.")],
    img_side="right",
    accent=TERRA,
    img_caption="Your 30% is now on the transcript, and it has to be defensible.",
    note_text="THE 70/30 SPLIT (4 min) - the highest-impact slide in this session\n\n"
             "THIS IS WHERE POLICY BECOMES PERSONAL. Principles persuade; accountability "
             "compels.\n\n"
             "THE HISTORY, briefly: for years, continuous assessment marks were often handled "
             "casually - generous scoring, lost record books, marks reconstructed at the end of "
             "term. The reform now records the 30% on a school-based transcript with far greater "
             "transparency and quality assurance.\n\n"
             "STATE THE CONSEQUENCE PLAINLY: 'If a learner's transcript is questioned and your "
             "30% cannot be defended, it is the learner who loses a place - and your school whose "
             "credibility is damaged.'\n\n"
             "THE SOLUTION IS NOT MORE PAPERWORK. It is the simple disciplines we are covering "
             "today: a rubric written BEFORE the task, applied consistently, kept on record, and "
             "the same across all classes teaching the same subject.\n\n"
             "RAISE IT AS A DEPARTMENTAL QUESTION, not just individual: 'Do all of us teaching the "
             "same subject at the same level use the same criteria? Could we produce the evidence "
             "if NaSIA asked?'\n\n"
             "This is the natural bridge into Session 3 on principles.")

slide_diagram(
    "split_70_30.png",
    "The 70 / 30 Split \u2014 In One Picture",
    "Why your internal marks carry more weight than they used to",
    "Thirty per cent marked by you, seventy by WAEC \u2014 and the 30 is now the part that follows the learner.",
    note_text="THE 30/70 SLIDE IN ONE PICTURE (2 min)\n\n"
    "USE THIS IMMEDIATELY AFTER the previous slide. Some teachers will not have absorbed the "
    "split from the text; the ring makes it visual in about four seconds.\n\n"
    "POINT AT THE GREEN BAND: 'Thirty per cent. That is yours. Nobody else sets it, nobody "
    "else marks it. That is roughly one full subject grade worth of a learner's future.'\n\n"
    "POINT AT THE GOLD: 'Seventy per cent is WAEC. Same paper, same day, everybody in the "
    "country. Nothing you can do about that in this room.'\n\n"
    "THE LINE THAT LANDS: 'The 70 is fixed. The 30 is the part you actually control - and it "
    "is now written down in a way that follows this child out of your school.'\n\n"
    "HONEST CAVEAT TO GIVE: internal marks used to be soft data that nobody checked. That is "
    "no longer true. Mark generously and you inflate a transcript. Mark harshly and you damage "
    "a real future. Neither is fair.\n\n"
    "DO NOT let this become a session about WAEC's grading. Keep it on the 30 and what it "
    "demands of the teacher.")

slide_table(
    "Continuous Assessment \u2014 What Is Required of You",
    "The practical obligations, in plain terms",
    ["Requirement", "What it means in practice", "Where teachers commonly slip"],
    [["Assess against curriculum learning outcomes",
      "Your tasks must test the standards in the Standards-Based Curriculum for your level.",
      "Setting questions on content you taught but which is not in the curriculum standards \u2014 or missing standards you did not teach."],
     ["Use a variety of assessment methods",
      "Not just written tests. Include projects, portfolios, performance, practicals, observation and oral work.",
      "Relying on class tests alone, then finding the 30% is thin and undefensible."],
     ["Record results systematically",
      "Every learner's results kept accurately across the term, with the evidence behind them.",
      "Losing the record book, or reconstructing marks at the end of term. This is the most common quality-assurance failure."],
     ["Give feedback learners can act on",
      "Results must feed back into teaching and into the learner's next steps \u2014 not just a mark.",
      "Marking and recording without ever handing work back or re-teaching the gaps."],
     ["Accommodate learners with SEN",
      "Extra time, alternative formats and appropriate accommodations, applied fairly.",
      "Treating accommodations as favours rather than requirements \u2014 and having no record of what was provided."],
     ["Be transparent and consistent",
      "Learners know the criteria in advance; the same rubric is used across classes.",
      "Two teachers marking the same subject to different standards, which undermines the whole transcript."]],
    col_w=[3.15, 4.80, 4.14],
    intro="Six obligations that come with the 30%. None of them is optional, and all of them are auditable.",
    row_h=0.94, fsize=11,
    note_text="CONTINUOUS ASSESSMENT REQUIREMENTS (3 min)\n\n"
             "This slide turns policy into a checklist. Go down the list quickly - do not read "
             "every word aloud. Instead, for each row, ask: 'How many of us are confident we could "
             "produce the evidence for this tomorrow?'\n\n"
             "THE TWO MOST IMPORTANT ROWS are rows 3 and 6:\n"
             "  - Row 3, RECORDS: the most common quality-assurance failure is not bad teaching, it "
             "is missing evidence. Marks reconstructed at the end of term cannot be defended.\n"
             "  - Row 6, CONSISTENCY: if two teachers marking the same subject apply different "
             "standards, the transcript loses meaning. This is a departmental responsibility, not "
             "an individual one.\n\n"
             "ROW 5, SEN: this is a legal and ethical requirement, not a courtesy. And note the "
             "second half - you need a RECORD of what accommodations you provided.\n\n"
             "PRACTICAL ASK: 'Before you leave today, check that your record for this term is "
             "complete and current. Not at the end of term - now.'")

slide_concept(
    "The School-Based Transcript \u2014 Why Records Matter",
    "The document that turns your mark book into a learner's future",
    "The school-based transcript records the internal 30% formally, with far greater "
    "transparency and quality assurance than the old system. It accompanies the learner beyond "
    "your school \u2014 to further study, to placements, to employers.",
    "Imagine two learners with identical WASSCE grades. One school's 30% rests on rubrics, "
    "marked work and a clean record book. The other's rests on marks entered the week before "
    "submission with no evidence behind them. If both transcripts are scrutinised, only one "
    "survives. The learner does not lose a place because they knew less \u2014 they lose it "
    "because somebody's record-keeping was loose. Your record book is not admin. It is a "
    "young person's evidence.",
    extra=[("The test to apply to yourself:  ", TERRA, True, False),
           ("if an inspector asked you to justify one learner's 30% mark today, could you show the "
            "work, the rubric and the date? If not, the record needs attention.", GREY, False, False)],
    accent=TERRA,
    notes_text="THE TRANSCRIPT (2 min)\n\n"
             "MAKE THIS CONCRETE. The transcript sounds like an administrative document; it is "
             "actually a young person's evidence base for the next stage of their life.\n\n"
             "USE THE CONTRAST on the slide - two learners, same exam grades, different quality of "
             "internal records. Ask: 'Which 30% would you want to be defending?'\n\n"
             "THE SELF-TEST IS THE TAKEAWAY: 'If an inspector asked you to justify one learner's "
             "30% mark today, could you show the work, the rubric and the date?' Most teachers "
             "will realise the answer is 'not entirely' for at least some learners. That is not "
             "an accusation - it is exactly the gap the reform is designed to close.\n\n"
             "DO NOT make this feel like a threat. Frame it as protecting the learner and "
             "protecting yourself. Then hand over to the break.")

slide_diagram(
    "transcript.png",
    "What the Transcript Actually Looks Like",
    "A worked example \u2014 and the question an inspector will ask about it",
    "Every number in the internal columns must be traceable to a piece of marked work.",
    note_text="THE TRANSCRIPT, CONCRETELY (3 min)\n\n"
    "WHY A WORKED EXAMPLE: 'transcript' is an abstract word, and most teachers have never seen "
    "one. Showing a table with real numbers makes the obligation concrete.\n\n"
    "WALK DOWN ONE ROW, e.g. Core Mathematics: portfolio 8, project 9, exam 7 - internal total "
    "24. Then WASSCE 52, final 76, grade B2. Show that the internal 24 is a QUARTER of the "
    "final 76. It is not a rounding detail.\n\n"
    "THE INSPECTOR QUESTION, at the bottom of the diagram: 'Show me the work behind this 24.' "
    "Read it out. Then ask the room: 'If NaSIA walked in today and asked me that about learner "
    "number 14 in this class, what could I actually put on the table?'\n\n"
    "WHAT GOOD LOOKS LIKE: the portfolio piece, the project, the rubric it was marked against, "
    "and a date. Four things. If you have those, the mark is defensible.\n\n"
    "BE GENTLE HERE. Most teachers will privately realise the answer is 'not entirely'. That "
    "is not an accusation - it is precisely the gap this reform is designed to close, and the "
    "reason this workshop exists.\n\n"
    "BREAK HANDOVER: thank them, give the time, and say you will be at the front during the "
    "break for anyone with a specific subject question.")

def slide_break():
    s = new()
    picture_cover(s, "01_hero_classroom.jpg", 0, 0, SW, SH)
    from pptx.oxml.ns import qn
    ov = rect(s, 0, 0, SW, SH, GREEN_DK)
    solid_el = ov.fill._xPr.find(qn('a:solidFill'))
    clr = solid_el.find(qn('a:srgbClr'))
    clr.append(clr.makeelement(qn('a:alpha'), {'val': '88000'}))
    rect(s, 0, 0, SW, 0.13, GOLD)
    tf = textbox(s, M + 0.15, 2.35, CW - 0.3, 2.6)
    para(tf, "Break", 56, WHITE, bold=True, font=HEAD_FONT, first=True, space_after=12)
    para(tf, "15 minutes  \u00b7  back at 01:00", 20, GOLD, font=HEAD_FONT, space_after=20)
    para(tf, "When we return: the seven principles that hold up every good assessment \u2014 and "
             "then you will design your own task and rubric.", 14,
         RGBColor(0xD8, 0xE8, 0xE1), space_after=0, line_spacing=1.15)
    footer(s, dark=True)
    notes(s, "BREAK (15 min)\n\n"
             "ANNOUNCE THE RETURN TIME CLEARLY - people lose track. Write it on the board.\n\n"
             "USE THE BREAK WELL YOURSELF:\n"
             "  - Check that the materials for the practical workshop are ready: A3 or flip-chart "
             "paper, marker pens, sticky notes, and printed copies of the task and rubric "
             "templates if you have them.\n"
             "  - Set up the room for group work if the layout needs to change - four or five "
             "tables of four to six people.\n"
             "  - Prepare the wall space for the gallery walk later.\n\n"
             "IF PEOPLE DRIFT BACK LATE, start Session 3 with the room rather than waiting - it "
             "pulls the rest in.\n\n"
             "COFFEE CHAT OPPORTUNITY: this is your best chance to hear what people really think "
             "about the 70/30 slide. Listen more than you talk.")
    return s


slide_break()


# ---------------------------------------------------------------- SESSION 3
slide_session(3, "Principles of\nEffective Assessment",
              15,
              "Seven pillars. Remove one, and the roof leans.",
              img="05_pillars.jpg",
              notes_text="SESSION 3 (15 min)\n\n"
              "Section 2 of the manual lists seven principles of effective assessment. This is the "
              "part people skim - and it is the part that explains why some assessments work and "
              "others do not.\n\n"
              "Frame it as seven pillars holding up a roof. You do not need all seven perfect; you "
              "need all seven present.\n\n"
              "PACE NOTE: 15 minutes for seven principles plus an overview means roughly 90 seconds "
              "each. Move briskly. The examples matter more than the definitions.")

slide_grid_cards(
    "Seven Pillars \u2014 Hold Up the Roof",
    "Everything that follows rests on these",
    "05_pillars.jpg",
    [("1 \u00b7 Validity", "Am I measuring the RIGHT thing?"),
     ("2 \u00b7 Reliability", "Do I get the SAME result every time, whoever marks it?"),
     ("3 \u00b7 Fairness and ethics", "Does every learner get an EQUITABLE chance?"),
     ("4 \u00b7 Transparency", "Do learners know WHAT they are judged on and HOW?"),
     ("5 \u00b7 Inclusivity", "Is it accessible to ALL learners, without exception?"),
     ("6 \u00b7 Practicability", "Can I ACTUALLY do this with my time and my class size?"),
     ("7 \u00b7 Assessment utility", "Is it USEFUL \u2014 does anybody learn anything from it?")],
    accent=GREEN,
    img_caption="Seven pillars. Remove one, and the roof leans.",
    note_text="THE SEVEN PILLARS (2 min overview)\n\n"
             "Read all seven quickly to give the shape of the session - do not dwell yet, each one "
             "gets its own slide next.\n\n"
             "Ask the room to guess which pillar is most often broken in real practice. Common "
             "answers: transparency and practicability. Both are true - we keep the criteria secret "
             "and then design assessments nobody has time to mark.\n\n"
             "Promise: 'By the end of the next twelve minutes you will recognise which of these you "
             "already do well, and which one is quietly costing your learners marks.'")

slide_concept(
    "1 \u00b7 Validity \u2014 Measuring the Right Thing",
    "The pillar most often broken without anyone noticing",
    "Are you actually measuring the thing you say you are measuring? A valid assessment measures "
    "the important learning outcomes of the curriculum \u2014 not just whatever was easiest to set.",
    "A trader uses a weighing scale to sell yam. If she used that scale to measure the LENGTH of "
    "the yam, the reading would be worthless \u2014 not because the scale is broken, but because it "
    "is measuring the wrong thing. Now think of your last class test. Your learning outcome said "
    "learners should ANALYSE data. But every question you set asked them to RECALL a definition. "
    "Your test was neat. It was well marked. It was simply measuring the wrong thing. That is "
    "an invalid assessment \u2014 and the learners who studied hard still lost marks for it.",
    extra=[("The teacher's quick validity check:  ", GREEN, True, False),
           ("\u201cDoes this question actually test the skill I said I was teaching?\u201d  "
            "If the answer is no, the marks it produces mean very little.", GREY, False, False)],
    notes_text="VALIDITY (2 min)\n\n"
             "THE TRAP: validity fails silently. An invalid test still produces marks, still fills "
             "the mark book, still looks professional. Nobody complains. But the numbers do not "
             "mean what everyone thinks they mean.\n\n"
             "THE CLEAREST TEST: compare your learning outcome to your questions. If the outcome "
             "says 'analyse' and your questions say 'state', you have an invalid assessment. Ask "
             "teachers to silently do this check on the last test they set. Most will find at least "
             "one problem.\n\n"
             "THE FIX - and this is the highest-value practical tip in this session: build a TABLE "
             "OF SPECIFICATION. A simple grid mapping each question to a topic and a thinking "
             "level. The manual asks teachers to 'create a learning and assessment plan (i.e. table "
             "of test specification)'. Without it, we all default to recall questions, because "
             "recall questions are the easiest to write. The grid forces balance.\n\n"
             "LINK FORWARD: we will use exactly this grid in the practical workshop at 1:35.")

slide_image_text(
    "2 \u00b7 Reliability \u2014 The Same Result Every Time",
    "Reliability is about trust in the numbers",
    "06_scales.jpg",
    [("Plain English",
      "Do you get the same result every time, whoever is doing the marking? Reliable results are "
      "dependable enough to make decisions with."),
     ("The classroom picture",
      "A market scale that gives the same weight for the same pile of tomatoes \u2014 on Monday or "
      "Saturday, whoever is standing at the stall. Now imagine the scale was generous on Monday "
      "and harsh on Saturday. Nobody would trust that market again."),
     ("Where we break it",
      "If this term's 30 marks in your subject are generously given and next term's are harsh, "
      "learners are being ranked against a ruler that keeps changing length. With the 30% now on "
      "a transcript, that is no longer just unfair \u2014 it is indefensible."),
     ("How the manual says to fix it",
      "Clear learning outcomes \u00b7 a colleague reviews your marking \u00b7 more than one "
      "assessment method for the same outcome \u00b7 proper marking schemes with weightings \u00b7 "
      "rubrics given OUT before the task \u00b7 a calm, suitable room.")],
    img_side="right",
    accent=GREEN,
    img_caption="Same pile, same reading \u2014 every time.",
    note_text="RELIABILITY (2 min)\n\n"
             "OPEN WITH THE MANUAL'S OWN WARNING - it is blunt and it stings: the 'connoisseur' "
             "approach to assessment - that is, 'I know it when I see it but I cannot put it into "
             "words' - is NOT acceptable. Read that line out loud. It is a direct challenge to the "
             "way many of us mark.\n\n"
             "LINK BACK TO SESSION 2: why does reliability suddenly matter so much? Because the 30% "
             "is now on a transcript. Unreliable marking is unfair in a way learners can feel but "
             "not prove.\n\n"
             "THE MOST POWERFUL HABIT on this slide: give your rubric or marking scheme OUT BEFORE "
             "the task, not after. The manual says this explicitly for performance and practical "
             "assessment.\n\n"
             "CHEAP RELIABILITY HACK to offer teachers: swap a set of scripts with a colleague in "
             "the same subject and mark each other's sample. The manual recommends having another "
             "teacher review the work. You will be shocked how differently two people read the same "
             "answer. This is also a departmental consistency measure.")

slide_diagram(
    "dartboard.png",
    "Validity and Reliability Together",
    "The two pillars side by side \u2014 the target explains both in one picture",
    "Validity is aiming at the right target. Reliability is hitting the same place every time. You need both.",
    note_text="THE TARGET (3 min)\n\n"
    "THIS IS THE BEST-REMEMBERED DIAGRAM IN ASSESSMENT TRAINING. Give it the time.\n\n"
    "TOP LEFT - valid and reliable. Tight cluster on the centre. You measured the right thing "
    "and you would get it again tomorrow. That is the goal.\n\n"
    "TOP RIGHT - reliable but not valid. Tight cluster, wrong place. 'This is the teacher who "
    "is very consistent and very wrong. Their marks are tidy, their record book is neat, and "
    "they are confidently measuring the wrong thing. Consistency does not rescue you from "
    "assessing the wrong outcome.'\n\n"
    "BOTTOM LEFT - valid but not reliable. Spread around the centre. Right target on average, "
    "but no single mark can be trusted. 'If this learner is ranked 12th today, is that real? "
    "The average is fine. The individual mark is a coin toss.'\n\n"
    "BOTTOM RIGHT - neither. Scatter. No use at all - not for a mark, not for feedback, not "
    "for planning.\n\n"
    "THE TEACHING MOVE: ask the room which panel is the most dangerous. Most will say bottom "
    "right. The answer is the TOP RIGHT - because it looks competent. Nobody investigates a "
    "neat, consistent record book.\n\n"
    "CONNECT TO THE 30%: on a transcript, the top-right panel is the one that damages a "
    "learner without anyone noticing.\n\n"
    "TRANSITION: 'Validity and reliability are the two technical pillars. The next five are "
    "about ethics and practicality.'")

slide_concept(
    "3 \u00b7 Fairness and Ethics \u2014 An Equitable Chance",
    "Fairness is not treating everyone identically \u2014 it is giving everyone a real chance",
    "Assessment strategies should give learners an equitable opportunity to demonstrate what they "
    "know and can do, taking into account their ability, learning styles, gender and Special "
    "Educational Needs. It also means never assessing what you have not taught.",
    "A test question asks learners to explain the rules of cricket. A learner from a village that "
    "has never played cricket \u2014 and never seen it \u2014 cannot answer, and fails. That "
    "question did not measure physics. It measured cricket. Now swap it for a question about the "
    "physics of a loaded trotro braking at a junction, and every learner in the room has something "
    "to think with. Same physics. Same difficulty. Fair.",
    extra=[("The manual's red flags:  ", TERRA, True, False),
           ("culturally biased content  \u00b7  unfamiliar words and examples  \u00b7  assessing "
            "content you never taught  \u00b7  failing to give SEN learners extra time or "
            "alternative formats.", GREY, False, False)],
    accent=TERRA,
    notes_text="FAIRNESS AND ETHICS (2 min)\n\n"
             "START WITH THE CRICKET EXAMPLE. It is memorable and it is exactly the failure mode "
             "the manual warns against - 'culturally biased or discriminatory content, unfamiliar "
             "words, questioning or examples'.\n\n"
             "THEN ASK THE HARD QUESTION: 'How many of us have set questions using examples from "
             "books written far from here, about things our learners have never seen?' Let that "
             "sit. The fix is not easier questions - it is LOCAL, FAMILIAR CONTEXTS for the same "
             "rigour. That is what the trotro example shows.\n\n"
             "THE SECOND BIG IDEA: do not assess what you have not taught. The manual lists this "
             "under fairness, and it is the single most common fairness complaint learners "
             "make.\n\n"
             "LINK BACK TO THE ICEBREAKER: 'Fairness is where the word FEAR from this morning "
             "comes from. Learners are not afraid of being measured. They are afraid of being "
             "measured unfairly.'")

slide_concept(
    "4 \u00b7 Transparency \u2014 No Secret Exams",
    "If learners have to guess what you want, you have already made it harder than it needs to be",
    "Making the assessment process and criteria clear and understandable to learners. They should "
    "know what is being judged, how it is judged, and what counts as a pass.",
    "Two teachers set the same project. The first teacher says \u201cdo the project, I will mark "
    "it\u201d \u2014 and the learners spend the week guessing. The second teacher hands out the "
    "rubric on day one: what is expected, the marks for each part, what a top answer looks like. "
    "Both teachers mark with the same rubric. The second class almost always scores higher \u2014 "
    "not because they worked harder, but because they knew where to aim. Assessment should not be "
    "a guessing game you are the referee of.",
    extra=[("Transparency also requires you to:  ", GREEN, True, False),
           ("share the learning outcomes being assessed  \u00b7  tell learners the pass mark  \u00b7  "
            "share results openly with learners and parents/guardians  \u00b7  give learners a route "
            "to seek review and redress if they disagree with a score.", GREY, False, False)],
    accent=GREEN,
    notes_text="TRANSPARENCY (2 min)\n\n"
             "THE KILLER LINE: 'Assessment should not be a guessing game where the teacher is the "
             "only person who knows the answer.'\n\n"
             "THE MOST COMMON OFFENCE: hiding the rubric until after the marking, usually to 'stop "
             "them cheating'. Point out that a learner who aims perfectly at a published rubric is "
             "not cheating. They are learning. That is exactly what we want.\n\n"
             "THE REDRESS POINT is the one nobody expects. The manual requires you to 'provide an "
             "opportunity for learners to seek review and redress'. Ask: 'If a learner in this "
             "school disputes a mark, what actually happens? Is there a route?' In many schools "
             "there is none. That is a gap worth naming.\n\n"
             "LINK FORWARD: transparency and validity are partners. Learners cannot demonstrate a "
             "learning outcome they were never told about. This is also the fastest pillar to fix "
             "- hand out the rubric next week.")

slide_concept(
    "5 \u00b7 Inclusivity \u2014 Fair and Accessible to ALL",
    "Not an extra duty. It is a standing requirement of the framework",
    "Creating assessment practices that are fair and accessible to every learner \u2014 regardless "
    "of gender, disability, poverty, background or learning difference. The framework names three "
    "lenses you must apply: GESI, SEL and SEN.",
    "A learner with low vision in your class cannot read an ordinary small-print paper \u2014 but "
    "they understand the work perfectly. A learner who is hard of hearing misses the spoken "
    "instructions you gave once, quickly, at the front. A learner who is dyslexic knows the "
    "science but loses marks to their handwriting. In all three cases we did not measure their "
    "learning. We measured their eyes, their ears and their handwriting. Inclusivity means "
    "removing that barrier so the mark reflects the learning.",
    extra=[("GESI", GREEN, True, False), (" = Gender Equality and Social Inclusion.   ", GREY, False, False),
           ("SEL", GREEN, True, False), (" = Socio-Emotional Learning.   ", GREY, False, False),
           ("SEN", GREEN, True, False), (" = Special Educational Needs.", GREY, False, False)],
    accent=GREEN,
    notes_text="INCLUSIVITY (2 min)\n\n"
             "THE KEY REFRAME, and the single most memorable line in this session: 'When a learner "
             "with low vision fails your test, you have not discovered that they do not know the "
             "work. You have discovered that they cannot see your paper. Those are completely "
             "different findings, and only one of them is about learning.'\n\n"
             "DECODE THE THREE ACRONYMS - teachers see these letters constantly and rarely have "
             "them explained:\n"
             "  - GESI - Gender Equality and Social Inclusion. No learner disadvantaged by gender, "
             "disability, poverty or background.\n"
             "  - SEL - Socio-Emotional Learning. The learner's feelings, confidence and "
             "relationships matter and are assessed, not just their marks.\n"
             "  - SEN - Special Educational Needs. Learners needing specific support.\n\n"
             "THE MANUAL'S PRACTICAL TOOLS to name out loud: braille, oral translation, "
             "text-to-speech and other AI supports, sign language interpretation, assistive "
             "technology, extra time, alternative formats.\n\n"
             "A WARNING WORTH GIVING: build fairness into your RUBRIC. If you are assessing "
             "science understanding, do not dock marks for handwriting or grammar by accident. "
             "Decide deliberately. We will act on this in the practical workshop.\n\n"
             "CONNECT TO THE 30%: if accommodations are provided but never recorded, you cannot "
             "defend the mark later.")

slide_concept(
    "6 \u00b7 Practicability \u2014 Can You Actually Do It?",
    "The pillar that kills the most beautiful assessment plans",
    "The assessment must be feasible, convenient, efficient and successful with the real "
    "resources you have \u2014 your actual class size, your actual timetable, your actual materials.",
    "You design a wonderful one-hour one-on-one oral assessment for your Government class. It is "
    "rigorous. It is valid. It is inclusive. It is also completely impossible with 60 learners "
    "\u2014 and by Thursday it will have collapsed into chaos, you will be three weeks behind, and "
    "you will abandon it. The plan failed not because it was bad, but because it was not "
    "practicable. A good assessment is one you can still run in week eight, when you are tired and "
    "the term is crowded.",
    extra=[("The three practicability questions:  ", TERRA, True, False),
           ("Do I have the materials and the security?  \u00b7  Does this format suit my class size "
            "and their level?  \u00b7  Do I have the TIME to set it, mark it, AND give constructive "
            "feedback?", GREY, False, False)],
    accent=TERRA,
    notes_text="PRACTICABILITY (2 min)\n\n"
             "OPEN WITH THE HONEST OBSERVATION: 'We have all designed a beautiful assessment in "
             "January that we quietly stopped running by March.' This slide is about why that "
             "happens - and it is not a character flaw. It is a design flaw.\n\n"
             "THE THREE QUESTIONS on the slide are the whole test. Walk through them for a real "
             "example. A 20-item test: markable in one evening, feedback possible. A "
             "one-hour-per-learner oral: mathematically impossible.\n\n"
             "THE POINT TEACHERS MISS: set, mark AND give feedback. Many assessments are "
             "practicable to administer but NOT practicable to give feedback on. Since feedback is "
             "where the learning actually happens, an assessment you cannot give feedback on has "
             "already failed.\n\n"
             "VALIDATE THEIR REALITY: yes, class sizes are large and time is short. That is exactly "
             "why practicability is written into the national framework as a PRINCIPLE, not an "
             "excuse. The framework expects you to choose methods that fit your actual conditions.\n\n"
             "PRACTICAL TIP for large classes: use strategies that scale - whole-class response, "
             "peer assessment with a rubric, exit cards, group projects with individual "
             "accountability. These are the tools for the next session.")

slide_concept(
    "7 \u00b7 Assessment Utility \u2014 Is It Actually Useful?",
    "The question that separates data collection from real assessment",
    "The assessment must be useful. It must give you and the learner information worth having, "
    "and that information must lead to action.",
    "You set a class test. You mark it. You enter the scores in the mark book. You file the "
    "papers. Term ends. Now ask honestly: did anybody learn anything from that test? If the test "
    "did not change what you taught next, and did not change what the learner did next, then all "
    "that happened was data collection. Utility is the difference between a test that teaches and "
    "a test that merely records.",
    extra=[("Utility in practice:  ", GREEN, True, False),
           ("state the intended use of the result up front  \u00b7  produce credible information for "
            "learners and parents  \u00b7  give constructive feedback  \u00b7  and then ACT on it "
            "\u2014 re-teach, regroup, intervene, adapt.", GREY, False, False)],
    accent=GREEN,
    notes_text="ASSESSMENT UTILITY (2 min)\n\n"
             "THIS IS THE SLIDE THAT TIES THE SEVEN TOGETHER. You can satisfy validity, "
             "reliability, fairness, transparency, inclusivity and practicability - and still "
             "produce an assessment that nobody uses. That is the most common failure of all, "
             "because it is invisible.\n\n"
             "THE TEST TO OFFER: 'Name one thing you changed last term because of an assessment "
             "result.' Give people a moment. Ask two or three to answer out loud. The answers show "
             "what utility looks like in practice.\n\n"
             "LINK BACK TO THE SOUP: a cook who tastes the soup and then does nothing about it has "
             "wasted the taste. Tasting without adjusting is not cooking. Assessing without acting "
             "is not assessment.\n\n"
             "CLOSE THE SESSION: 'Seven pillars. Validity, reliability, fairness, transparency, "
             "inclusivity, practicability, utility. You will not get all seven perfect every week. "
             "But when an assessment disappoints you, come back to this list - the answer is almost "
             "always hiding in one of the seven.'\n\n"
             "BRIDGE TO SESSION 4: 'Now let us look at the tools that help you satisfy all seven "
             "at once.'")
# ---------------------------------------------------------------- SESSION 4
slide_session(4, "Assessment Strategies\nAcross Levels",
              20,
              "Choosing the right tool \u2014 and matching it to your learners.",
              img="10_rubric_card.jpg",
              notes_text="SESSION 4 (20 min)\n\n"
              "This session is the toolkit. Session 3 gave the principles; this gives the methods.\n\n"
              "We cover: how to choose a strategy, the Depth of Knowledge staircase, the three "
              "families of strategy, four strategies in detail (portfolio, debate, practical, "
              "rubric), and how to pitch the same strategy at SHS 1, 2 and 3.\n\n"
              "PACE NOTE: this session sets up the practical workshop. Every strategy here is "
              "available for teachers to use in their own design at 1:35. Say so explicitly.")

slide_image_text(
    "Choosing a Strategy \u2014 Match the Tool to the Job",
    "The most common mistake is using a test for everything",
    "09_plan_journey.jpg",
    [("Start with the outcome, not the tool",
      "Ask first: what should the learner be able to DO? Then choose the method that lets them show it. If the outcome is \u201cdismantle and reassemble\u201d, a written test cannot show it."),
     ("Three families of strategy",
      "INFORMAL (questions, observation, exit cards) for quick in-lesson checks. FORMAL WRITTEN (tests, essays, MCQs) for recording. PERFORMANCE (projects, practicals, portfolios, debates) for real application."),
     ("Vary your methods deliberately",
      "The manual requires a variety of assessment methods. If all your methods are written tests, you are assessing recall and missing everything else \u2014 and your 30% is thin."),
     ("Match the level, not the label",
      "The same strategy pitches differently at SHS 1, 2 and 3. Scaffold heavily in Form 1; expect independence by Form 3. We will work through this shortly."),
     ("Check the seven pillars before you commit",
      "Will it measure the right thing (validity)? Can you mark it consistently (reliability)? Can you actually run it in week eight (practicability)?")],
    img_side="right",
    accent=GREEN,
    img_caption="Start with what learners must do, then pick the tool.",
    note_text="CHOOSING A STRATEGY (3 min)\n\n"
             "THE CENTRAL MISTAKE this slide corrects: teachers default to the written test because "
             "it is familiar, not because it fits the outcome. That is a validity failure, and it is "
             "the single most common one in Ghanaian classrooms.\n\n"
             "THE REFRAME: 'Do not ask what assessment you will set. Ask what the learner must be "
             "able to DO. Then ask what task would let them show it.'\n\n"
             "THE THREE FAMILIES are worth naming so teachers see the range:\n"
             "  - INFORMAL - questions, observation, exit cards. Zero preparation, immediate.\n"
             "  - FORMAL WRITTEN - tests, essays, MCQs. Markable, recordable, auditable.\n"
             "  - PERFORMANCE - projects, practicals, portfolios, debates. Real application.\n\n"
             "LINK BACK TO SESSION 2: the manual requires a VARIETY of methods, and the 30% has to "
             "be defensible. If your entire 30% is class tests at three points in the term, it is "
             "both thin evidence and a narrow picture of the learner.\n\n"
             "POINT FORWARD: 'Everything on this slide is available to you in the workshop at 1:35. "
             "You are not designing in the abstract.'")

slide_diagram(
    "dok_staircase.png",
    "Depth of Knowledge \u2014 Climbing the Staircase",
    "How DEEP is the thinking? Not how difficult the words look.",
    "The words can stay simple while the thinking goes deep \u2014 that is the whole point.",
    note_text="THE STAIRCASE (3 min)\n\n"
    "THEN the table on the next slide gives the detail. This slide is the shape of the idea.\n\n"
    "READ THE BANNER ALOUD: 'The thinking gets deeper, not the words harder.' This is the "
    "single most important sentence about DoK. Teachers commonly think a difficult question is "
    "one with big words. DoK says the opposite: you can ask a very deep question in very simple "
    "language.\n\n"
    "WALK UP THE STAIRS with the Social Studies examples on the next slide. DoK 1 is recall - "
    "state three causes. DoK 3 is a farmer whose yield is falling: diagnose it and justify it. "
    "Same subject, same simple words, completely different thinking.\n\n"
    "WHERE WASSCE SITS: the external paper increasingly rewards DoK 3 and 4. If every question "
    "you set is DoK 1, your learners pass your class tests comfortably and then meet a paper "
    "that asks them to reason - and they have never been asked to. Check your last test: count "
    "how many questions were recall.\n\n"
    "THE PRACTICAL RULE: a good SHS paper is not all DoK 4. It is a STAIRCASE - some recall to "
    "secure the basics, most at DoK 2-3, and at least one genuinely extended task where you "
    "can.\n\n"
    "LINK FORWARD: in the workshop at 1:35 you will write a DoK level into your task "
    "deliberately, instead of discovering it afterwards.")

slide_table(
    "Depth of Knowledge \u2014 The Staircase of Thinking",
    "How DEEP is the thinking? Not how difficult the words look.",
    ["Level", "Name", "Plain English", "SHS example \u2014 Social Studies"],
    [["DoK 1", "Recall / Reproduction",
      "Give back a fact you were taught.",
      "\u201cState three causes of climate change.\u201d"],
     ["DoK 2", "Skills and Concepts",
      "Use your knowledge in a familiar routine.",
      "\u201cExplain how the greenhouse effect works, using a diagram.\u201d"],
     ["DoK 3", "Strategic Thinking",
      "Reason, justify, plan. There is more than one route to the answer.",
      "\u201cA cocoa farmer's yield is falling. Using the data given, diagnose the likely cause and justify your answer.\u201d"],
     ["DoK 4", "Extended Thinking",
      "Investigate over time, combine ideas, produce something original.",
      "\u201cConduct a week-long survey in your community on waste disposal and present recommendations to the school assembly.\u201d"]],
    col_w=[1.15, 2.05, 3.55, 5.34],
    intro="DoK stands for Depth of Knowledge. It is a ladder, not a difficulty rating \u2014 the words can stay simple while the thinking goes deep.",
    row_h=0.78, fsize=11,
    note_text="DEPTH OF KNOWLEDGE (4 min)\n\n"
             "THE BIGGEST MISCONCEPTION, and the reason this slide exists: teachers think a question "
             "is DoK 4 because it uses big words or long sentences. It does not. DoK measures the "
             "TYPE of thinking required.\n\n"
             "SAY THIS CLEARLY: 'A long question with big vocabulary that asks you to remember "
             "something is still DoK 1. A short, simple question that asks you to judge a real "
             "situation and defend your judgement can be DoK 3.'\n\n"
             "WALK UP THE STAIRCASE using the climate change example on the slide - it is drawn "
             "directly from the ministry's own curriculum resource, so teachers will meet it "
             "again.\n\n"
             "WHY IT MATTERS FOR VALIDITY: if your learning outcome is at DoK 3 ('analyse', "
             "'evaluate') and your test questions are all at DoK 1 ('state', 'list', 'define'), "
             "your assessment is invalid - no matter how well marked. This is exactly the validity "
             "failure we described in Session 3.\n\n"
             "INTERACTIVE, 60 SECONDS: ask teachers to take one question from their last test and "
             "label it. Most will find their whole paper sits at DoK 1 and 2. That is a diagnosis, "
             "not a criticism - and it is easy to fix with a table of specification.\n\n"
             "USE THIS IN THE WORKSHOP: when they design their task at 1:35, they should aim at "
             "least one question at DoK 3.")

slide_image_text(
    "Portfolio \u2014 The Learner's Own Story",
    "Not a folder of everything. A chosen collection, with reflection.",
    "15_portfolio.jpg",
    [("Plain English",
      "A carefully chosen collection of a learner's work across time, showing progress towards the learning outcomes. Formative as it grows, summative when a rubric is applied at the end."),
     ("The classroom picture",
      "Two things separate a real portfolio from a pile of papers. First it is CURATED \u2014 the learner chooses what goes in, and learns by choosing. Second it carries REFLECTION \u2014 beside each piece, a short note: \u201cI chose this because\u2026\u201d and \u201cnext time I would\u2026\u201d."),
     ("Why it is worth the effort",
      "It is the clearest evidence you will ever have of progress over time \u2014 and it is pure Assessment as Learning. The learner is judging their own growth, with evidence."),
     ("The rule that saves your evenings",
      "You do not mark every piece. You mark the REFLECTIONS and the final selection. That is where the thinking is visible."),
     ("Where the manual helps",
      "Appendix C.1 gives worked portfolio exemplars for Science and Mathematics, including the rubric. You do not have to invent it.")],
    img_side="right",
    accent=GREEN,
    img_caption="Curated, with reflection \u2014 otherwise it is just a folder.",
    note_text="PORTFOLIO (2 min)\n\n"
             "THE MISCONCEPTION: most teachers think a portfolio is a folder of every exercise the "
             "learner has ever done. That is not a portfolio; it is a filing cabinet.\n\n"
             "THE TWO WORDS THAT MATTER: CURATED and REFLECTION.\n"
             "  - Curated - the learner chooses. Choosing is itself an act of assessment: they must "
             "judge which work best shows their learning.\n"
             "  - Reflection - a short written note beside each piece. 'I chose this because...' and "
             "'next time I would...'.\n\n"
             "THIS IS ASSESSMENT AS LEARNING in its purest form - the AaL from Session 1.\n\n"
             "PRACTICAL NOTE: portfolios take a term to build. Start small - one subject, one class, "
             "three pieces across one term.\n\n"
             "TIME-SAVING TIP TO OFFER: do not mark every piece. Mark the reflections and the final "
             "selection. That is where the thinking is visible.\n\n"
             "WHERE TO START: the manual's Appendix C.1 has Science and Mathematics exemplars with "
             "rubrics already written. Point teachers there - it removes the biggest barrier.")

slide_image_text(
    "Debate \u2014 Thinking Out Loud",
    "A strategy that assesses reasoning, not just recall",
    "16_debate.jpg",
    [("Plain English",
      "Learners take a position on an issue, argue it with evidence, and respond to the opposing case. You assess the quality of the thinking, not just the conclusion."),
     ("The classroom picture",
      "A Government class debates: \u201cShould Ghana raise the minimum wage?\u201d Half the room argues for it, half against. The Economics or English teacher uses the same format for their own content."),
     ("What you are actually assessing",
      "Clarity of the argument  \u00b7  use of evidence, not just opinion  \u00b7  quality of rebuttal \u2014 do they answer the other side?  \u00b7  listening  \u00b7  respectful disagreement."),
     ("The rule that makes it work",
      "Give the rubric BEFORE the debate, and give learners time to research. A debate sprung on them with no preparation assesses confidence, not competence \u2014 and it rewards the loudest voices.")],
    img_side="left",
    accent=TERRA,
    img_caption="Assessing the reasoning, not the volume.",
    note_text="DEBATE (2 min)\n\n"
             "WHY DEBATE IS UNDERRATED: it is one of the few strategies that assesses THINKING IN "
             "PUBLIC. A written answer hides the reasoning; a debate exposes it. You hear whether "
             "the learner actually understands, or has merely memorised.\n\n"
             "THE FAIRNESS WARNING - and this is the important bit: a debate sprung on learners with "
             "no preparation does not assess competence. It assesses confidence. It rewards the "
             "loudest voices and punishes the shy but thoughtful learner. This is a FAIRNESS failure "
             "(pillar 3) and a VALIDITY failure (pillar 1) at the same time.\n\n"
             "WHAT TO PUT IN THE RUBRIC: clarity of argument, use of evidence, quality of rebuttal, "
             "listening and response, respectful disagreement. Notice that 'volume' and 'confidence' "
             "do not appear.\n\n"
             "INCLUSIVITY NOTE: offer a written or small-group alternative for learners who genuinely "
             "cannot perform in front of a large audience. The learning outcome is the reasoning, not "
             "the performance.")

slide_image_text(
    "Practical and Performance \u2014 Doing It For Real",
    "The manual calls it authentic assessment: it looks like real life",
    "17_practical_lab.jpg",
    [("Plain English",
      "Learners demonstrate a skill in a real or simulated setting, and you judge the DOING \u2014 not just the writing about it."),
     ("The classroom picture",
      "The Chemistry learner stands at the bench and titrates. The Technical Skills learner dismantles and reassembles an engine. The Visual Arts learner throws a pot. The Music learner performs. You are watching the skill, not reading about it."),
     ("What the teacher must do first",
      "Design a task requiring real application  \u00b7  provide resources, guidance and support  \u00b7  evaluate against PREDETERMINED criteria  \u00b7  model the skill  \u00b7  act as coach, not just examiner."),
     ("Why the rubric matters most here",
      "In a practical, judgement happens live with no chance to re-read. Without a rubric written in advance you will mark the confident performer higher than the careful one."),
     ("No lab? Still possible",
      "The manual explicitly allows real OR simulated contexts. An improvised set-up still counts.")],
    img_side="right",
    accent=GREEN,
    img_caption="You are judging the doing, not the writing about it.",
    note_text="PRACTICAL AND PERFORMANCE (2 min)\n\n"
             "THIS IS WHAT THE MANUAL CALLS 'performance-based / practical / authentic' assessment. "
             "'Authentic' simply means: it looks like real life.\n\n"
             "THE KEY DIFFERENCE: you do not ask them to describe a titration on paper. They stand at "
             "the bench and titrate. You judge what they DO.\n\n"
             "THE TEACHER'S FIVE JOBS, straight from the manual: design a task requiring real-life "
             "application, provide resources and support, evaluate against predetermined criteria, "
             "model the skill, and serve as coach throughout.\n\n"
             "WHY THE RUBRIC IS NON-NEGOTIABLE HERE: in a written test you can re-read a script. In a "
             "practical, the judgement happens live and once. Without criteria fixed in advance, you "
             "will unconsciously mark the confident performer higher than the careful one - and that "
             "is a reliability failure, not an opinion.\n\n"
             "SMALL-SCHOOL REALITY: if you have no laboratory, a simulated or improvised setting "
             "still counts. The manual explicitly allows 'real or simulated' contexts. Do not let "
             "lack of equipment stop you.")

slide_image_text(
    "Rubric \u2014 The Recipe Card",
    "The single most useful document in your assessment toolkit",
    "10_rubric_card.jpg",
    [("Plain English",
      "A table showing exactly what counts and how many marks each part is worth, at each level of quality."),
     ("The classroom picture",
      "It is a recipe card with measurements. Not \u201cadd enough salt\u201d \u2014 but \u201cone teaspoon\u201d. Not \u201cwrite a good essay\u201d \u2014 but \u201cthree clear arguments, each supported by one example, maximum two pages.\u201d"),
     ("Why it changes everything",
      "It removes the guesswork for the learner AND the bias from the marker. It is the fastest route to both transparency and reliability at the same time."),
     ("The rule nobody follows",
      "Give the rubric OUT BEFORE the task \u2014 never after. The manual says this explicitly for performance and practical assessment."),
     ("A simple three-band template",
      "EXCELLENT (full marks)  \u00b7  GOOD (most marks)  \u00b7  DEVELOPING (some marks). Three bands is enough to start. Four criteria is plenty.")],
    img_side="left",
    accent=GOLD,
    img_caption="A recipe card with measurements, not guesses.",
    note_text="RUBRIC (3 min) - this is the direct setup for the practical workshop\n\n"
             "THE ANALOGY TO STRESS: 'Add enough salt' is not a recipe. 'One teaspoon' is a recipe. "
             "Most of our instructions to learners are the first kind - and then we are surprised "
             "when thirty learners produce thirty different things.\n\n"
             "THE DOUBLE BENEFIT: a rubric improves BOTH transparency and reliability simultaneously. "
             "It tells the learner where to aim, and it stops two markers giving wildly different "
             "scores to the same script. It is the best-value document in this whole workshop.\n\n"
             "THE RULE TO REPEAT: handout BEFORE the task. A rubric revealed after marking is not "
             "transparency. It is a post-mortem.\n\n"
             "KEEP IT SIMPLE - this is important for the workshop: three bands (Excellent, Good, "
             "Developing) and four criteria is plenty. Teachers over-engineer rubrics and then do "
             "not use them. A simple rubric used consistently beats a detailed one abandoned in "
             "week three.\n\n"
             "TRANSITION INTO THE WORKSHOP: 'In ten minutes you are going to build one of these for "
             "your own subject. Not a perfect one - a usable one.'")

slide_table(
    "Matching Strategy to Level",
    "The same strategy, pitched differently at SHS 1, 2 and 3",
    ["Strategy", "SHS 1 (Form 1)", "SHS 2 (Form 2)", "SHS 3 (Form 3)"],
    [["Project work",
      "Short, tightly scaffolded, one clear deliverable. Heavy guidance.",
      "Multi-step with milestones. Learner plans more of the process.",
      "Extended and largely independent. Expects synthesis and original thinking."],
     ["Debate",
      "Structured roles, given positions, short speaking time.",
      "Learners research their own evidence. Rebuttal expected.",
      "Full format, independent research, higher expectation of rebuttal quality."],
     ["Practical / performance",
      "Follow a modelled procedure step by step with close supervision.",
      "Carry out the procedure with less prompting; record results accurately.",
      "Design or adapt a procedure; analyse and evaluate outcomes."],
     ["Portfolio",
      "Teacher-guided selection. Simple reflection prompts supplied.",
      "Learner selects with increasing independence. Reflection in more depth.",
      "Self-directed curation with critical reflection on growth."],
     ["DoK level to expect",
      "Mainly DoK 1\u20132, with guided DoK 3.",
      "DoK 2\u20133 regularly.",
      "DoK 3\u20134, including extended tasks."]],
    col_w=[2.55, 3.18, 3.18, 3.18],
    intro="Same tool, different scaffolding. Pitching Form 1 work at Form 3 level is not rigour \u2014 it is a validity failure.",
    row_h=0.85, fsize=11,
    note_text="MATCHING STRATEGY TO LEVEL (3 min)\n\n"
             "WHY THIS SLIDE EXISTS: teachers often reuse the same task across all year groups, or "
             "pitch everything at the middle. Both are validity failures - you are not assessing the "
             "standard for that level.\n\n"
             "READ DOWN THE DoK ROW ESPECIALLY: SHS 1 should be mainly DoK 1-2 with GUIDED DoK 3. By "
             "SHS 3 you should be expecting DoK 3-4. If a Form 1 teacher is setting DoK 4 extended "
             "research tasks, the learners will fail - not because they are weak, but because the "
             "task was pitched wrong.\n\n"
             "THE SCAFFOLDING PRINCIPLE: same strategy, different amount of support. A Form 1 debate "
             "can work brilliantly if you give them the positions and short speaking time. The same "
             "format with independent research is a Form 3 task.\n\n"
             "LINK TO EVIDENCE: this connects directly to the manual's point about ensuring 'the load "
             "or the length of the tasks are appropriate to the level of the learner'.\n\n"
             "ASK THE ROOM: 'Which of these levels do you mostly teach? Does your last assessment "
             "match that row?' Give them thirty seconds to look.")

slide_image_text(
    "Quick Wins \u2014 Three-Minute Tools You Can Use Tomorrow",
    "No budget. No printing. No extra marking load.",
    "14_peer_assessment.jpg",
    [("Exit card",
      "Last three minutes of the lesson: on a slip, \u201cwrite one thing you learned today and one thing you are still confused about.\u201d Collect at the door. You now know exactly where to start next lesson."),
     ("Entry ticket",
      "First three minutes: one question that shows what they already bring. Free diagnostic assessment, every single day."),
     ("Think-Pair-Share",
      "Think alone, then discuss with a partner, then share with the class. Three minutes of real thinking time means the quiet learners at the back have something to say when you ask."),
     ("K-W-L chart",
      "Three columns: what I Know, what I Want to know, what I Learned. Makes thinking visible and makes the learner reflective."),
     ("Self- and peer-assessment",
      "Learners check their own or each other's work AGAINST THE RUBRIC. Only works if they have the criteria in hand \u2014 otherwise it becomes a popularity contest.")],
    img_side="right",
    accent=GOLD,
    img_caption="Learners assessing each other \u2014 with the rubric in hand.",
    note_text="QUICK WINS (2 min) - the close of Session 4\n\n"
             "THIS IS THE 'TAKE IT HOME' SLIDE. Its job is to prove that formative assessment is not "
             "extra work. Every tool here costs three minutes and no money.\n\n"
             "PICK ONE AND COMMIT THE ROOM. Do not ask them to adopt all five. Ask each teacher to "
             "choose ONE tool to try in their very next lesson and tell the person next to them which "
             "one they chose. Public commitment increases follow-through dramatically.\n\n"
             "WHY EXIT CARDS ARE THE BEST STARTING POINT: three minutes, scrap paper, works with 60 "
             "learners, and tells you precisely where to begin tomorrow. If you only ever adopt one "
             "thing from today, adopt the exit card.\n\n"
             "THE PEER-ASSESSMENT WARNING is worth saying: peer assessment fails when learners do not "
             "have the rubric. Give them the criteria, and train them to be specific and kind.\n\n"
             "TRANSITION INTO THE WORKSHOP: 'That is the toolkit. Now let us use it. In a moment you "
             "will design your own task and rubric - and you will have 25 minutes to do it properly.'")


# ---------------------------------------------------------------- WORKSHOP
slide_session("PRACTICAL WORKSHOP", "Design a Task & Rubric",
              25,
              "Something you will actually use \u2014 not something you plan to use.",
              img="22_workshop_groups.jpg",
              notes_text="PRACTICAL WORKSHOP (25 min) - THE HEART OF THE DAY\n\n"
              "Everything before this point has been preparation. Everything after it is reflection. "
              "This is where teachers produce something real.\n\n"
              "WHAT THEY PRODUCE: one learning outcome, one assessment task, and one rubric - for a "
              "topic they will teach in the next month.\n\n"
              "BEFORE YOU START, CHECK:\n"
              "  - Groups formed (4-6 people, ideally by subject or level so they can help each "
              "other)\n"
              "  - A3 or flip-chart paper and marker pens at every table\n"
              "  - Sticky notes available\n"
              "  - The templates visible (slide 38) and, if possible, printed\n\n"
              "TIMING IS TIGHT. Use the time cues on slide 39 and keep to them. If a group is "
              "running behind, that is fine - the point is a usable draft, not a perfect document.\n\n"
              "YOUR ROLE: circulate constantly. Sit with each table for two minutes. Ask two "
              "questions: 'What is your learning outcome?' and 'How will you know a good one when "
              "you see it?' Those two questions catch nearly every problem.")

slide_activity(
    "Your Brief \u2014 Design a Task & Rubric",
    "25 minutes \u00b7 in groups of four to six \u00b7 use a topic you teach this term",
    "25 MINUTES",
    [("Work in groups",
      "Groups of four to six. If possible, group by subject or level so you can help each other. One person keeps time."),
     ("Use a REAL topic",
      "Choose something you will actually teach in the next month. A real topic produces a real tool. An invented one produces a nice-looking document you will never use."),
     ("Produce three things",
      "ONE learning outcome  \u00b7  ONE assessment task  \u00b7  ONE simple rubric. That is all. Not a scheme of work \u2014 one task."),
     ("Put it on the paper",
      "Use the A3 sheet. Write large. This goes on the wall for the gallery walk at 2:00, so it must be readable from a distance."),
     ("Finish with the checklist",
      "At minute 20 you will check your work against the six-point quality checklist. Fix the easiest problem, then stop.")],
    img="22_workshop_groups.jpg",
    accent=GOLD,
    note_text="SETTING UP THE WORKSHOP (3 min of the 25)\n\n"
             "READ THE BRIEF ALOUD, then repeat the three deliverables clearly: ONE outcome, ONE "
             "task, ONE rubric. Teachers will try to design a whole term's assessment. Stop that "
             "kindly.\n\n"
             "WHY A REAL TOPIC: this is the most important instruction on the slide. Invented topics "
             "produce documents that never get used. Ask them to open their scheme of work and pick "
             "something from the next four weeks.\n\n"
             "GROUPING: by subject or level if possible, so they can share subject-specific "
             "knowledge. Mixed groups work too, but take longer to get going.\n\n"
             "TIMEKEEPING: appoint one timekeeper per group and give them the cue times. You will "
             "also call them out from the front.\n\n"
             "WHAT GOOD LOOKS LIKE at the end: a clear learning outcome with an action verb, a task "
             "that genuinely lets learners show that outcome, and a rubric with three bands and "
             "around four criteria. That is it. Perfection is not the goal.")

slide_table(
    "Step 1 & 2 \u2014 The Outcome and the Strategy",
    "Minutes 0\u201310 \u00b7 do these before you think about the task",
    ["Step", "What to decide", "The test to apply"],
    [["1. Write the learning outcome",
      "ONE sentence. \u201cBy the end of this, the learner should be able to\u2026\u201d Use an ACTION VERB: explain, calculate, design, evaluate, demonstrate.",
      "Could I show someone else this sentence and would they know exactly what to teach and assess? If not, it is too vague."],
     ["1b. Check the DoK level",
      "Is your verb at the right thinking level for your class? State = DoK 1. Explain = DoK 2. Analyse / justify = DoK 3. Design / evaluate = DoK 4.",
      "Does the thinking level match the standard in the curriculum for this level?"],
     ["2. Choose the assessment strategy",
      "Which family? INFORMAL (observation, questioning) \u00b7 WRITTEN (test, essay) \u00b7 PERFORMANCE (project, practical, portfolio, debate).",
      "Could a learner who genuinely has this skill fail this task for a reason that has nothing to do with the skill?"],
     ["2b. Sanity-check the seven pillars",
      "VALID (measures the outcome) \u00b7 RELIABLE (markable the same way twice) \u00b7 PRACTICABLE (runnable in week eight with your class size).",
      "If any of the three fails, change the strategy now \u2014 not after you have built the task."]],
    col_w=[2.65, 5.45, 3.99],
    intro="Ten minutes on these two steps. Skipping them is why assessment tasks fail \u2014 the task is always easy once the outcome is clear.",
    row_h=1.10, fsize=11,
    note_text="MINUTES 0-10: OUTCOME AND STRATEGY\n\n"
             "CALL THIS OUT AT THE START: 'For the first ten minutes, do not think about the task. "
             "Outcome first. The task is easy once the outcome is clear - and almost impossible "
             "before it is.'\n\n"
             "CIRCULATE AND LISTEN FOR THESE PROBLEMS:\n\n"
             "PROBLEM 1 - VAGUE OUTCOMES. 'Understand photosynthesis' is not an outcome you can "
             "assess - 'understand' is invisible. Push them to a verb that produces evidence: "
             "explain, describe, calculate, demonstrate, compare, evaluate.\n\n"
             "PROBLEM 2 - WRONG DoK LEVEL. A Form 1 class given a DoK 4 design task will fail. Ask: "
             "'Is that the level your curriculum expects for this year group?'\n\n"
             "PROBLEM 3 - STRATEGY MISMATCH. The outcome is 'demonstrate safe titration technique' "
             "but the chosen strategy is a written test. Ask the test on the slide: 'Could a learner "
             "who genuinely has this skill fail this task for an unrelated reason?'\n\n"
             "TWO QUESTIONS TO ASK EVERY TABLE:\n"
             "  1. 'What is your learning outcome?' - if they cannot say it in one sentence, that is "
             "the problem.\n"
             "  2. 'How will you know a good one when you see it?' - this is the beginning of the "
             "rubric.\n\n"
             "DO NOT LET GROUPS SKIP AHEAD. If a group has not written a clear outcome by minute 10, "
             "sit with them.")

slide_table(
    "Step 3 & 4 \u2014 The Task and the Rubric",
    "Minutes 10\u201320 \u00b7 build the task, then the marking criteria",
    ["Step", "What to decide", "The test to apply"],
    [["3. Write the task",
      "Exactly what the learner must do. Include the instructions you would give them, and what they must produce. One task \u2014 not a list of five.",
      "Could a learner who has the skill show it here? And would a learner who only memorised definitions be unable to fake it?"],
     ["3b. Consider your SEN learners",
      "What accommodations would this task need? Extra time, larger print, oral alternative, a written alternative to a performance.",
      "Have I decided this in advance \u2014 rather than improvising on the day and forgetting to record it?"],
     ["4. Build the rubric",
      "THREE bands: Excellent / Good / Developing. FOUR criteria at most. Each cell says what the work must show to earn that band.",
      "Could a colleague pick this up and mark the same script the same way I would? That is reliability."],
     ["4b. Write the learner-facing version",
      "One or two sentences telling learners what \u201cexcellent\u201d looks like. This is the part handed out BEFORE the task.",
      "Have I described the WORK, not the learner? \u201cClear argument with evidence\u201d \u2014 not \u201cvery good student\u201d."]],
    col_w=[2.65, 5.45, 3.99],
    intro="Ten minutes. Keep the rubric simple \u2014 three bands and four criteria. A simple rubric used consistently beats a detailed one abandoned in week three.",
    row_h=1.10, fsize=11,
    note_text="MINUTES 10-20: TASK AND RUBRIC\n\n"
             "THE MOST IMPORTANT INSTRUCTION: keep the rubric simple. Three bands, four criteria. "
             "Teachers over-engineer rubrics and then do not use them. Tell them plainly: 'A simple "
             "rubric you actually use beats a detailed one you abandon in week three.'\n\n"
             "CIRCULATE AND WATCH FOR:\n\n"
             "PROBLEM 1 - THE UNFAKEABLE TEST. Ask: 'Would a learner who only memorised definitions "
             "be unable to fake their way through this task?' If they could fake it, the task is "
             "testing recall, not the outcome.\n\n"
             "PROBLEM 2 - RUBRIC DESCRIBING THE LEARNER RATHER THAN THE WORK. This is very common. "
             "'Excellent: a very good student' tells nobody anything. Push them to: 'Excellent: "
             "three clear arguments, each supported by one piece of evidence.'\n\n"
             "PROBLEM 3 - TOO MANY CRITERIA. If a group has eight criteria, ask them which four "
             "matter most. Marking with eight criteria is slow and inconsistent.\n\n"
             "PROBLEM 4 - FORGETTING SEN. Ask directly: 'Who in your class would struggle to access "
             "this task, and what will you do about it?' This is the inclusivity pillar made "
             "practical - and note on the slide that accommodations must be decided in advance and "
             "recorded.\n\n"
             "AT MINUTE 20, CALL TIME and move them to the quality checklist. Do not let groups "
             "drift into minute 25 - they need the checklist time.")

slide_table(
    "Quality Checklist \u2014 Before You Finish",
    "Minutes 20\u201325 \u00b7 check your own work \u00b7 fix the easiest problem \u00b7 stop",
    ["Check", "Ask yourself", "If it fails"],
    [["Validity", "Does my task actually measure the learning outcome I wrote \u2014 or has it drifted into something easier to set?",
      "Change the task, not the outcome."],
     ["Reliability", "Could a colleague mark this the same way I would, using only my rubric?",
      "Add specificity to the vague rubric cells."],
     ["Fairness", "Does the task assume experiences or contexts some of my learners have not had?",
      "Re-context it to something familiar \u2014 the physics of a trotro, not of cricket."],
     ["Inclusivity", "Who cannot access this task, and what have I put in place for them?",
      "Add an accommodation and write it down."],
     ["Transparency", "Could I hand this rubric to learners BEFORE the task, and would it help them aim?",
      "Rewrite the language so a learner can understand it."],
     ["Practicability", "Can I run this in week eight, mark it, AND give feedback \u2014 with my real class size?",
      "Scale it down. A smaller task you complete beats a grand one you abandon."]],
    col_w=[2.15, 6.35, 3.59],
    intro="Six questions. Five minutes. This is the step that separates a nice-looking document from a usable tool.",
    row_h=0.83, fsize=11,
    note_text="MINUTES 20-25: QUALITY CHECK\n\n"
             "THIS IS THE STEP TEACHERS WANT TO SKIP. Do not let them. Five minutes here is what "
             "makes the difference between a document and a tool.\n\n"
             "INSTRUCT THEM: 'Read your own work against these six questions. Find the ONE problem "
             "that is easiest to fix. Fix it. Then stop.' Do not attempt to fix everything - there "
             "is no time and it is not necessary.\n\n"
             "THE THREE MOST COMMON FAILURES you will see in this check:\n"
             "  - VALIDITY: the task drifted into recall because recall is easy to write.\n"
             "  - RELIABILITY: rubric cells that say 'good', 'satisfactory', 'shows understanding' "
             "without describing what would earn them.\n"
             "  - PRACTICABILITY: a task that would take two weeks to mark. This is the one to catch "
             "most gently, because the enthusiasm is real but the plan will collapse.\n\n"
             "CLOSE THE WORKSHOP SEGMENT: 'Put your sheet on the wall. You have five minutes' break "
             "before the gallery walk - stretch, then come back ready to read other people's work.'\n\n"
             "IF A GROUP FINISHED EARLY: ask them to write the learner-facing version of their "
             "rubric, or to write the SEN accommodation, or to predict what the weakest learner in "
             "their class would produce.")
# ---------------------------------------------------------------- GALLERY WALK
slide_activity(
    "Gallery Walk \u2014 Read Each Other's Work",
    "10 minutes \u00b7 walk, read, leave written feedback",
    "10 MINUTES",
    [("Stick your work on the wall",
      "One sheet per group. Write your subject and level at the top so people know the context."),
     ("Walk and read \u2014 6 minutes",
      "Move around the room reading other groups' work. No talking for the first three minutes \u2014 just read. Then you may discuss quietly."),
     ("Leave feedback on sticky notes",
      "One note per sheet. Write ONE thing that would make this task better or fairer. Be specific and kind \u2014 \u201cthe rubric cell for Good is still vague\u201d, not \u201cneeds work\u201d."),
     ("Collect your own notes \u2014 2 minutes",
      "Go back to your own sheet. Read the notes people left for you. You do not have to agree with them all \u2014 but read them."),
     ("Share-outs \u2014 2 minutes",
      "We will hear from two or three groups at the front: what did you design, and what did the feedback make you reconsider?")],
    img="19_gallery_walk.jpg",
    accent=TERRA,
    note_text="GALLERY WALK (10 min)\n\n"
             "PURPOSE: teachers learn more from critiquing each other's work than from listening to "
             "you. This also builds the habit of the manual's 'critiquing as an assessment "
             "strategy' - they are doing exactly what they should have learners do.\n\n"
             "SET THE NORMS BEFORE YOU START - this matters, or the feedback is either useless or "
             "hurtful:\n"
             "  - ONE note per sheet. Not a page of comments.\n"
             "  - SPECIFIC, not general. 'The Good column does not say how many examples' is useful. "
             "'Needs improvement' is not.\n"
             "  - KIND. These are colleagues' drafts, not final products.\n\n"
             "ENFORCE THE QUIET READING PHASE. Three minutes of silence at the start is what makes "
             "people actually read rather than chat.\n\n"
             "IF YOU ARE SHORT OF TIME: cut the share-outs, not the feedback notes. The written "
             "feedback is where the learning is.\n\n"
             "SHARE-OUT QUESTIONS to ask the two or three groups at the front:\n"
             "  - 'What did you design, in one sentence?'\n"
             "  - 'What did the feedback make you reconsider?'\n"
             "  - 'What is the one thing you will change before you use this?'\n\n"
             "LISTEN FOR a group that says they realised their rubric was too vague, or their task "
             "was not really testing their outcome. Name that publicly - it validates the process.\n\n"
             "USE THE GALLERY WALK AS YOUR EVIDENCE: take photographs of the wall. These are real "
             "assessment tools designed by your staff, and they are useful for the school's "
             "professional development record.")


# ---------------------------------------------------------------- USING DATA
slide_session("USING ASSESSMENT DATA", "Feedback & Records",
              10,
              "Turning marks into learning \u2014 and into records you can defend.",
              img="21_records.jpg",
              notes_text="USING ASSESSMENT DATA (10 min)\n\n"
              "Two halves: FEEDBACK (what you give back to the learner) and RECORDS (what you keep "
              "and why).\n\n"
              "This is objective 4 of the workshop: turn assessment results into targeted feedback, "
              "better teaching decisions, and clear records.\n\n"
              "PACE NOTE: 10 minutes is tight. Keep the feedback slide warm and human; keep the "
              "records slide crisp and practical. Teachers already know records matter \u2014 they "
              "need the system, not the lecture.")

slide_image_text(
    "Feedback \u2014 The Good Mechanic",
    "Why \u201cgood work, 7/10\u201d is not feedback",
    "12_feedback_mechanic.jpg",
    [("The useless mechanic",
      "Says \u201cthe car is bad\u201d and takes your money. You leave knowing nothing, and the car is no better."),
     ("The good mechanic",
      "Opens the bonnet, points at ONE exact part and says: \u201cThis hose is cracked. Replace it and the overheating stops.\u201d You now know what to do next."),
     ("Plain English",
      "Feedback is information that helps a learner close the gap between where they are and where they need to be."),
     ("What it is NOT",
      "It is not a score. It is not praise. \u201cGood work, 7/10\u201d tells a learner nothing about what to do differently. A mark is a verdict; feedback is a map."),
     ("The rule teachers break",
      "Feedback without TIME TO ACT ON IT is decoration. If you hand back corrections and immediately move to the next topic, the feedback dies in the exercise book.")],
    img_side="right",
    accent=TERRA,
    img_caption="Point at the exact part, and say what to do next.",
    note_text="FEEDBACK (4 min)\n\n"
             "THE MECHANIC ANALOGY is the heart of this slide. Everyone has met both mechanics. Ask "
             "the room which one they would return to, and why. Then make the turn: 'Which mechanic "
             "are we to our learners?'\n\n"
             "THE HARD TRUTH: most of what we call feedback is a number plus a word. 'Good work, "
             "7/10' gives the learner a verdict, not a next step. The learner knows they got 7. They "
             "do not know what to do to get 9.\n\n"
             "THE MOST ACTIONABLE RULE - and the one teachers find genuinely relieving: feedback "
             "needs TIME. Ten minutes at the start of the next lesson for learners to read and act "
             "on your corrections is worth more than two hours of extra marking. Without that time, "
             "all your careful marking is decoration.\n\n"
             "LINK BACK TO THE PILLARS: feedback without time to act on it fails the UTILITY test "
             "from Session 3. You satisfied six pillars and produced something nobody used.\n\n"
             "SELF-CHECK TO OFFER: 'When you hand back a marked script, does the learner do anything "
             "with it other than look at the score?'")

slide_table(
    "Records \u2014 What to Keep, and Why",
    "Because the 30% is now on a transcript that follows the learner",
    ["What to keep", "Why it matters", "How to keep it manageable"],
    [["The assessment task itself",
      "You cannot defend a mark without showing what was asked. Keep a copy of the paper or task sheet, not just the marks.",
      "One folder per class per term. Digital photo of the task sheet if you prefer."],
     ["The rubric or marking scheme",
      "This is your evidence that marking was consistent \u2014 the reliability pillar made auditable.",
      "Keep the actual rubric, including any changes you made after marking started."],
     ["Marks, with a date",
      "Marks without dates cannot show progress over the term, and cannot be audited.",
      "A single clean mark book, updated after each assessment \u2014 not reconstructed at the end."],
     ["Evidence of a sample of learner work",
      "For any mark that could be questioned, you need at least a sample of the actual work.",
      "Keep the scripts for the assessments that feed the 30%, or a representative sample."],
     ["SEN accommodations provided",
      "Required by the framework \u2014 and needed to explain a mark later if it is questioned.",
      "A short note in the record against the learner's name: what was provided, and when."],
     ["Feedback given",
      "The manual requires that results feed back into learning. Evidence that they did.",
      "A note of the date feedback was returned is usually enough."]],
    col_w=[2.80, 5.30, 3.99],
    intro="Six things. If you can produce all six for any learner, your 30% is defensible.",
    row_h=0.82, fsize=11,
    note_text="RECORDS (4 min)\n\n"
             "OPEN BY REASSURING THEM: 'This is not about creating more paperwork. It is about "
             "keeping six things you probably already keep, in a way that can be found.'\n\n"
             "THE TEST TO OFFER: 'Pick one learner at random. Could you produce all six of these for "
             "them right now?' Most teachers will find two or three are missing or scattered. That "
             "is the gap - and it is fixable in an afternoon.\n\n"
             "THE ONE THAT IS MOST OFTEN MISSED is SEN accommodations. The framework requires both "
             "that they are provided AND, practically, you need the note to explain a mark later.\n\n"
             "THE DATE COLUMN on marks matters more than people expect. Marks without dates cannot "
             "show progress across the term, which is exactly what a portfolio or a transcript is "
             "meant to demonstrate.\n\n"
             "PRACTICAL ADVICE, keep it simple: 'One folder per class. One mark book. Date "
             "everything. Keep the rubric with the task. That is the whole system.'\n\n"
             "BRIDGE: 'Records are about the past. The next slide is about turning them into the "
             "future.'")

slide_image_text(
    "From Data to Action \u2014 The Whole Point",
    "Assessment that does not change anything was just data collection",
    "07_feedback_loop.jpg",
    [("Look for the pattern, not the person",
      "Do not just record 45 individual scores. Ask: which questions did the class fail together? Eighteen learners failing the same question is one teaching problem, not eighteen learning problems."),
     ("The three responses to a pattern",
      "RE-TEACH the whole class if most failed \u00b7  REGROUP and target support if a few failed \u00b7  MOVE ON if everyone passed. Decide deliberately, not by habit."),
     ("Close the loop in the next lesson",
      "The gap between the assessment and the response should be days, not weeks. Feedback that arrives three weeks later is history, not feedback."),
     ("Use it for your own planning",
      "The manual is explicit: assessment information should modify teaching and learning strategies. Your next week's lesson plan should be visibly shaped by last week's assessment."),
     ("Make it visible to learners",
      "Tell them what you noticed and what you are changing. 'I can see we all struggled with this \u2014 we will do it a different way today.' That single sentence turns assessment from judgement into teaching.")],
    img_side="right",
    accent=GREEN,
    img_caption="From mark book to next lesson plan.",
    note_text="FROM DATA TO ACTION (2 min)\n\n"
             "THIS IS OBJECTIVE 4 OF THE WORKSHOP, and it is the slide that closes the loop on the "
             "whole day.\n\n"
             "THE MOST USEFUL REFRAME: 'Look for the pattern, not the person.' The mark book "
             "encourages us to see 45 individual results. But 18 learners failing the same question "
             "is ONE teaching problem, not 18 learning problems. That insight alone changes how much "
             "work you have to do.\n\n"
             "THE THREE RESPONSES are worth naming clearly, because teachers often default to only "
             "one: if most failed, re-teach the whole class. If a few failed, target support. If "
             "everyone passed, move on. Deciding deliberately beats deciding by habit.\n\n"
             "CLOSE THE LOOP IN DAYS, NOT WEEKS. Give the number: 'Three weeks later is not "
             "feedback, it is history.'\n\n"
             "THE LAST BULLET is the one that changes classroom culture: tell learners what you "
             "noticed and what you are changing. 'I can see we all struggled with this - we will do "
             "it a different way today.' That single sentence turns assessment from judgement into "
             "teaching, which brings us back to the words from the icebreaker this morning.")


# ---------------------------------------------------------------- CHALLENGES
slide_solutions(
    "Common Challenges & Practical Solutions",
    "The five things you are all thinking about \u2014 with honest answers",
    [["\u201cMy class is too large to assess properly.\u201d",
      "60 learners, one teacher, no time. Individual oral assessment is impossible.",
      "Do not assess everyone individually. Use methods that scale: whole-class response, exit cards collected at the door, peer assessment with a rubric, group projects with individual accountability. Sample deeply rather than test everyone shallowly."],
     ["\u201cThere is not enough time in the term.\u201d",
      "The syllabus is crowded. Assessment feels like an addition on top of teaching.",
      "Assessment IS teaching, not an addition. Three minutes of exit cards replaces an hour of re-teaching material the class already knew. Start small \u2014 one exit card a week, not a full new system."],
     ["\u201cMarking already takes my evenings.\u201d",
      "Every script marked in full detail, late at night, then filed.",
      "Mark the drafts formatively and the final version summatively. Use rubrics so marking is a judgement against criteria, not a rewrite. And remember: feedback needs time for learners to act on it more than it needs length."],
     ["\u201cI was never trained to design rubrics.\u201d",
      "Rubrics feel like something other people know how to do.",
      "You do not have to start from a blank page. The NaCCA manual's Appendix C has worked exemplars and rubrics for portfolios, projects, debates, practicals, questioning and more. Adapt one."],
     ["\u201cThe learners only care about marks.\u201d",
      "Learners skip the comments and look at the score.",
      "That is learned behaviour, and it is rational \u2014 marks are what they have always been rewarded for. Change it by giving them time to act on feedback and by grading a draft formatively. When feedback visibly improves their next mark, they start reading it."],
     ["\u201cWhat about learners who will not engage?\u201d",
      "Some learners will not speak in a debate, will not participate in peer assessment.",
      "Offer alternative formats rather than excuses. A written alternative to a debate still assesses the reasoning. And check your own task first \u2014 often the barrier is the format, not the learner."]],
    intro="These are the real constraints \u2014 not excuses. Every answer below is something you can start this term without extra budget.",
    row_h=0.83, fsize=10,
    note_text="COMMON CHALLENGES (5 min)\n\n"
             "THIS SLIDE IS YOUR CREDIBILITY. If you do not address the real constraints honestly, "
             "everything else sounds like theory from someone who has not taught a large class.\n\n"
             "HOW TO RUN IT: read the first line of each challenge - the \u201cmy class is too "
             "large\u201d type opening - and let the room murmur in recognition. Then give the "
             "practical answer. Do not read the whole table aloud.\n\n"
             "VALIDATE FIRST, SOLVE SECOND. Say plainly: 'These are not excuses. These are real "
             "constraints, and the framework expects us to design around them - that is exactly "
             "what the practicability pillar means.'\n\n"
             "THE CHALLENGE MOST LIKELY TO COME UP, and the one to spend most time on: TIMB. "
             "Teachers feel assessment is added on top of an already crowded syllabus. The reframe: "
             "'Assessment is not on top of teaching. It is what tells you which teaching to stop "
             "doing.'\n\n"
             "ON RUBRICS - this is where you get the biggest return. Many teachers avoid rubrics "
             "because they think they must invent them. Tell them clearly: Appendix C of the NaCCA "
             "manual has exemplars for almost every strategy we covered today. Point them there.\n\n"
             "ON LEARNERS ONLY CARING ABOUT MARKS: this connects back to the icebreaker. If learners "
             "have only ever experienced assessment as judgement, of course they only look at the "
             "number. That is learned. Change the experience and the behaviour follows.\n\n"
             "IF YOU ARE SHORT OF TIME: cover the first three challenges only. Those are the ones "
             "in the room.")


# ---------------------------------------------------------------- CLOSING
slide_image_text(
    "Action Planning \u2014 One Commitment Each",
    "Five minutes \u00b7 write it down \u00b7 say it to the person beside you",
    "20_action_plan.jpg",
    [("Pick ONE thing \u2014 not five",
      "Choose the single change you will actually make. Not the ideal system \u2014 the realistic next step."),
     ("Write it specifically",
      "Weak: \u201cI will use more formative assessment.\u201d Strong: \u201cI will use an exit card at the end of every double period with Form 2 Science, starting Monday.\u201d"),
     ("Name the first step and the date",
      "What is the very first action? When will it happen? \u201cPrint the slips this weekend, use them Monday period 3.\u201d"),
     ("Say it out loud to your neighbour",
      "Thirty seconds each. Spoken commitment to a peer is much more likely to happen than a private resolution."),
     ("Choose from today's toolkit",
      "An exit card  \u00b7  a rubric written BEFORE the task  \u00b7  giving back one end-of-term paper for group correction  \u00b7  a peer-assessment cycle with the rubric  \u00b7  checking your record book is complete.")],
    img_side="right",
    accent=GREEN,
    img_caption="Not the ideal system. The realistic next step.",
    note_text="ACTION PLANNING (3 min)\n\n"
             "THE DISCIPLINE OF THIS SLIDE: ONE commitment. Teachers leave workshops with good "
             "intentions and change nothing because they try to change everything. Force the "
             "constraint.\n\n"
             "DRILL DOWN ON SPECIFICITY. When someone says 'I will do more formative assessment', "
             "push back kindly: 'With which class? How often? Starting when? What will you actually "
             "do at the end of the lesson?'\n\n"
             "THE STRONGEST VERSION names all four: WHAT, with WHICH CLASS, HOW OFTEN, STARTING "
             "WHEN. Read the model on the slide aloud as the standard.\n\n"
             "SPOKEN COMMITMENT: give exactly 30 seconds each, then move on. Do not skip this \u2014 "
             "publicly saying it to a peer genuinely increases follow-through.\n\n"
             "OFFER THE MENU at the bottom so nobody has to invent a starting point. Five options, "
             "all from today's content, all achievable this term without budget.\n\n"
             "IF YOU WANT TO STRENGTHEN FOLLOW-UP: collect the slips, or take a photograph of the "
             "board. Then at your next staff meeting, ask how the commitments went. That single "
             "follow-up question is what turns a workshop into a change.")

def slide_closing():
    s = new()
    picture_cover(s, "01_hero_classroom.jpg", 0, 0, SW, SH)
    from pptx.oxml.ns import qn
    ov = rect(s, 0, 0, SW, SH, GREEN_DK)
    solid_el = ov.fill._xPr.find(qn('a:solidFill'))
    clr = solid_el.find(qn('a:srgbClr'))
    clr.append(clr.makeelement(qn('a:alpha'), {'val': '89000'}))
    rect(s, 0, 0, SW, 0.13, GOLD)
    tf = textbox(s, M + 0.15, 1.30, CW - 0.3, 4.9)
    para(tf, "Thank You", 40, WHITE, bold=True, font=HEAD_FONT, first=True, space_after=18)
    para(tf, "This morning we wrote one word each for assessment. Most of them were about being "
             "judged.", 15, RGBColor(0xD8, 0xE8, 0xE1), space_after=10, line_spacing=1.16)
    para(tf, "Everything we did today was about moving those words \u2014 from judgement toward "
             "learning.", 15, RGBColor(0xD8, 0xE8, 0xE1), space_after=20, line_spacing=1.16)
    para(tf, "Taste the soup. Serve the meal. Know which one you are doing, and why.",
         19, GOLD, bold=True, font=HEAD_FONT, space_after=18)
    para(tf, "Your handouts have every term, every principle and every strategy we covered. "
             "The NaCCA manual is free online, and Appendix C has ready-made rubrics you can "
             "adapt tonight.", 13, RGBColor(0xC6, 0xDB, 0xD3), space_after=0, line_spacing=1.15)
    footer(s, dark=True)
    notes(s, "CLOSING (2 min)\n\n"
             "DO NOT SUMMARISE THE WHOLE DAY AGAIN. You have just done the action planning. Close "
             "with the callback to the icebreaker instead - it is the emotional bookend.\n\n"
             "DELIVER IT SLOWLY:\n"
             "'This morning I asked you for one word for assessment. Most of the words on that board "
             "were about being judged - stress, fear, marks, pressure. And I said then that the "
             "difference between those words and words like growth, feedback and help is the whole "
             "point of today.'\n\n"
             "'Everything we did was about moving those words across. Not by working harder - by "
             "tasting the soup. By knowing which form of assessment you are using and why. By writing "
             "the rubric before the task. By giving feedback they have time to act on.'\n\n"
             "'Taste the soup. Serve the meal. Know which one you are doing, and why.'\n\n"
             "THEN POINT TO THE HANDOUTS:\n"
             "  - The PDF handbook has every term, principle and strategy, with the classroom "
             "pictures, so you do not have to remember it all.\n"
             "  - The NaCCA Teacher Assessment Manual and Toolkit is free at "
             "curriculumresources.edu.gh. Appendix C is the part to go to first - ready-made "
             "rubrics and exemplars for almost every strategy we covered.\n\n"
             "FINALLY, THANK THEM PROPERLY. They gave up two and a half hours. Acknowledge that "
             "they designed real tasks today, not just listened. Then invite informal questions - "
             "people often ask the best questions once the formality drops.\n\n"
             "IF ANYONE STAYS BEHIND: those conversations are worth more than the session. Listen "
             "to what they say about the 70/30 slide especially - that is where the real anxiety is.")
    return s


slide_closing()

prs.save(OUT)
print("saved:", OUT)
print("slides:", len(prs.slides._sldIdLst))
