#!/usr/bin/env python3
"""
Build the KATON 2026 'Effective Assessment Practices' workshop deck.

Source of truth: Teacher Assessment Manual and Toolkit (NaCCA / Ministry of Education, Ghana).
Audience: Senior High School teachers.
Design goal: every technical term gets a PLAIN ENGLISH meaning + a GHANAIAN CLASSROOM PICTURE.
"""
import os
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

ROOT = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(ROOT, "assets", "opt")
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
                row_h=0.46, fsize=11.5):
    s = new()
    y = title_bar(s, title, sub)
    if intro:
        tf = textbox(s, M, y + 0.02, CW, 0.42)
        para(tf, intro, 12, GREY, italic=True, first=True, space_after=0,
             line_spacing=1.1)
        y += 0.48
    table(s, M, y + 0.06, CW, headers, rows, col_w=col_w, row_h=row_h, fsize=fsize)
    footer(s)
    notes(s, note_text)
    return s


# ---------------------------------------------------------------- 1. OPENING
slide_title()
slide_promises()

# ---------------------------------------------------------------- 2. WHAT IS ASSESSMENT
slide_section(1, "What Assessment\nReally Means",
              "Before we can do it well, we need to agree on what it actually is.",
              img="01_hero_classroom.jpg",
              notes_text="PART 1 (4 min)\n\n"
              "Most teachers hear 'assessment' and think 'exams and marking'. That is only "
              "one small part of it. In this section we reset the definition, then meet the "
              "three acronyms that confuse everybody — AfL, AaL and AoL.")

slide_concept(
    "Assessment — The Plain Meaning",
    "The single most important definition of the day",
    "Finding out what a learner knows and can do, and how far they have travelled "
    "towards the goal you set. It is gathering information about a learner in order "
    "to make a decision about their learning.",
    "It is not only the end-of-term paper. It is the question you asked in the middle "
    "of the lesson and the blank faces that answered it. It is the exercise book you "
    "marked at 9pm. It is watching a group argue over a titration and realising they "
    "have misunderstood the indicator. Every one of those is assessment — and every "
    "one of them changed what you did next.",
    extra=[("The official wording says the same thing:  ", INK, False, False),
           ("\u201cthe process of collecting information on the learner to help in deciding the degree "
            "to which the learner has achieved the expected learning outcomes/standards.\u201d",
            GREY, False, True)],
    notes_text="THE CORE DEFINITION (3 min)\n\n"
    "Read the plain version aloud. Then ask the room to name one thing they did last week "
    "that was assessment but NOT a test. Collect two or three answers.\n\n"
    "KEY POINT to land: assessment is not a document, it is a DECISION-MAKING process. "
    "If you gathered information and changed nothing, you collected data but you did not "
    "really assess.\n\n"
    "Note the careful wording in the official definition — the goal is not to rank learners, "
    "but to decide 'the degree to which' they have achieved a standard. That phrase matters.")

slide_table(
    "The Three Sisters: AfL, AaL and AoL",
    "These three acronyms cause the most confusion — here they are side by side",
    ["Acronym", "Full name", "Who is doing the work", "In one plain sentence"],
    [["AfL", "Assessment for Learning", "The TEACHER", "I check as we go, so I can teach better."],
     ["AaL", "Assessment as Learning", "The LEARNER", "The student checks themselves, so they learn better."],
     ["AoL", "Assessment of Learning", "The SYSTEM / teacher", "A final summary of what has been achieved."]],
    col_w=[1.35, 2.85, 2.35, 5.54],
    intro="AfL and AaL together make up what the manual calls FORMATIVE assessment. AoL is SUMMATIVE.",
    row_h=0.60,
    note_text="THE THREE SISTERS (4 min)\n\n"
    "This table is your anchor for the rest of the workshop. Refer back to it constantly.\n\n"
    "Say it this way: 'For' and 'as' are both about improving. 'Of' is about reporting.\n\n"
    "A simple test to give teachers — ask: 'Who is going to change what they do as a result "
    "of this information?' If the answer is the teacher, it is AfL. If it is the learner, it "
    "is AaL. If nobody changes anything and it just gets recorded, it is AoL.\n\n"
    "Important: AaL is the one we forget. We tell them their marks, but we never teach them "
    "to check themselves. That is the skill that carries them into university and work.")

# ---------------------------------------------------------------- 3. THE SOUP AND THE MEAL
slide_section(2, "The Soup and\nthe Meal",
              "The simplest way to understand formative and summative assessment.",
              img="02_soup_tasting.jpg",
              notes_text="PART 2 (6 min)\n\n"
              "This is the section that makes everything click. Ghanaian teachers remember this "
              "metaphor. Use it all the way through.\n\n"
              "FORMATIVE = you taste the soup while it is still on the fire.\n"
              "SUMMATIVE = you serve the meal at the table.\n\n"
              "Hold this image up for the rest of the workshop.")

slide_full_image(
    "02_soup_tasting.jpg",
    "FORMATIVE ASSESSMENT  ·  ASSESSMENT FOR LEARNING + AS LEARNING",
    "Tasting the Soup",
    "You taste while it is still on the fire. No salt? Add some. Too much pepper? Add water. "
    "You have not failed the soup — you have just improved it before anyone sits down to eat. "
    "That is assessment FOR learning, and it saves the meal.",
    note_text="THE SOUP (3 min)\n\n"
    "Ask the room: 'Does anybody here wait until the food is served at the table before they "
    "first taste it?' Wait for laughter and shaking heads. Of course not.\n\n"
    "Then the turn: 'So why do we do exactly that with our learners? Why do we find out what "
    "they misunderstood in the end-of-term exam, when it is too late to fix it?'\n\n"
    "KEY POINT: the cook who tastes is not being nosy. The cook who tastes is doing the "
    "professional thing. Tasting is not an interruption to cooking — it IS cooking.\n\n"
    "The manual's words: formative assessment 'enables teachers to modify or improve teaching "
    "and learning'. Notice the word IMPROVE. Not PROVE.")

slide_full_image(
    "03_meal_served.jpg",
    "SUMMATIVE ASSESSMENT  ·  ASSESSMENT OF LEARNING",
    "Serving the Meal",
    "The cooking is finished. This is the meal as it is. No amount of tasting will help now. "
    "You serve it, the family eats, and the meal is judged as served. That is assessment OF "
    "learning — end of term, WASSCE, certification.",
    note_text="THE MEAL (3 min)\n\n"
    "Summative assessment is not the enemy. It has a real and necessary job: selection, "
    "certification and placement. University admissions and employers need a final verdict. "
    "The manual says it provides 'a summary of the learner's overall achievement for selection, "
    "certification and placement purposes.'\n\n"
    "The point is not that summative is bad. The point is that summative ALONE is a poor diet. "
    "If you only ever serve meals and never taste, you will keep serving the same mistakes.\n\n"
    "Examples to name: end-of-term exams, WASSCE, end-of-programme exams, term papers, "
    "industrial attachment assessment.")

slide_table(
    "Formative vs Summative — Side by Side",
    "The same skills, two different jobs",
    ["", "FORMATIVE (the tasting)", "SUMMATIVE (the serving)"],
    [["When", "During teaching — before, during and after the lesson", "After teaching — end of unit, term, year, programme"],
     ["Purpose", "To IMPROVE learning while there is still time", "To PROVE and record overall achievement"],
     ["Who uses it", "Teacher and learner", "School, WAEC and other external bodies"],
     ["What it sounds like", "\u201cI see where you got stuck — let us try it this way\u201d", "\u201cThis is your final grade for the term\u201d"],
     ["Feels like", "A conversation, a question, a quick check, a show of hands", "A formal paper, a scheduled exam, a transcript"],
     ["Examples", "Questions, exit cards, class exercises, observation, peer assessment, drafts", "End-of-term exam, WASSCE, projects, portfolios, practicals"]],
    col_w=[1.85, 4.90, 5.34],
    row_h=0.56,
    note_text="SIDE BY SIDE (3 min)\n\n"
    "Walk down the rows. The 'What it sounds like' row is the one that lands with teachers — "
    "read both sentences out loud in the two different tones of voice.\n\n"
    "Ask: 'Which of these two voices do our learners hear most from us?' Be honest that for "
    "many of us, it is the second voice far more often than the first.\n\n"
    "Reassure: this is not about working harder. It is about redistributing the same effort.")

slide_concept(
    "The Golden Rule Most Teachers Miss",
    "The two do not live in separate boxes — they feed each other",
    "Summative results can be used FORMATIVELY, and formative work can count SUMMATIVELY. "
    "The label describes the PURPOSE the information is put to — not the test paper itself.",
    "You mark the end-of-term paper as usual. But instead of only recording the scores, you "
    "hand the scripts back, put learners in groups to find the answers to the questions they "
    "missed, and you re-teach the two topics the whole class failed. That exam just became a "
    "teaching tool. Equally: the class exercises and the group project you grade against a "
    "clear rubric — that is the 30% school-based assessment. Your everyday work is already "
    "counting towards the final grade.",
    extra=[("Why this matters for you:  ", TERRA, True, False),
           ("the same piece of work can do two jobs. You are not being asked to double your "
            "marking load — you are being asked to use the information twice.", GREY, False, False)],
    accent=TERRA,
    notes_text="THE GOLDEN RULE (4 min)\n\n"
    "This is the highest-value slide in the deck. Many teachers believe formative and summative "
    "are two separate systems that compete for time. They are not.\n\n"
    "The manual has a whole section on this (3.5.3 and 3.5.4). Read the second example in "
    "section 3.5.3 aloud if you have the manual open: 'instead of scoring a summative test, the "
    "teacher gives back the test to learners to discuss in groups to find answers to the "
    "questions that were wrongly answered.'\n\n"
    "PRACTICAL ASK: challenge every teacher to pick ONE end-of-term paper this year and give it "
    "back for group correction instead of just recording the marks. One paper. That is the "
    "commitment for the day.")

# ---------------------------------------------------------------- 4. DIAGNOSTIC
slide_section(3, "Before You Teach",
              "Finding out what they already bring to your classroom.",
              img="04_pantry_market.jpg",
              notes_text="PART 3 (4 min)\n\n"
              "Diagnostic assessment is the one everybody skips, because it happens before the "
              "topic when you are already behind schedule. And it is probably the one that saves "
              "the most time.")

slide_image_text(
    "Diagnostic Assessment — Checking the Pantry",
    "What do they already have before you start cooking?",
    "04_pantry_market.jpg",
    [("Plain English",
      "Checking what learners already know and can do BEFORE you start teaching — so you know "
      "where to begin."),
     ("The classroom picture",
      "Checking your pantry before you cook. Have you got rice? Oil? Is the yam already boiled? "
      "Only then do you decide what to cook. If you start cooking and discover there is no salt, "
      "you have wasted the gas."),
     ("In your classroom",
      "A five-minute pre-test or entry ticket at the start of a new topic. Ask one question on "
      "Monday about the topic you teach on Wednesday — and find out that half the class never "
      "understood last term's foundation topic."),
     ("Popular tools",
      "Pre-tests, entry tickets, K-W-L charts, concept maps, short interviews, a simple show of hands.")],
    img_side="right",
    accent=TERRA,
    img_caption="Check the pantry before you light the fire.",
    note_text="DIAGNOSTIC (4 min)\n\n"
    "Ask: 'Who has ever started a new topic, taught it beautifully, set a test — and discovered "
    "the whole class had the wrong foundation from last term?' Every hand goes up. That is the "
    "cost of skipping diagnostic assessment.\n\n"
    "KEY LINE: 'You cannot build on a foundation you have not checked.'\n\n"
    "Emphasise the timing — the manual says to use it 'prior or at the beginning of the teaching "
    "and learning process' and even 'before they come to class'.\n\n"
    "Reassure on workload: this does not need to be a formal test. An entry ticket with one "
    "question takes three minutes and tells you where to pitch the lesson.\n\n"
    "Mention that diagnostic assessment also identifies learners with diverse needs, which feeds "
    "directly into differentiated assessment later.")

slide_concept(
    "Differentiated Assessment — Same Destination, Different Route",
    "The phrase teachers most often misunderstand",
    "Using different ways of assessing so that every learner, whatever their level or need, has "
    "a fair chance to show what they can do. You adapt the ROUTE. You never lower the DESTINATION.",
    "You have a very strong learner and a learner who is struggling. Both must show they "
    "understand photosynthesis — that is the learning outcome and it does not change. One writes "
    "the essay. The other draws and labels the diagram and explains it orally. Both have proven "
    "the same outcome. That is differentiation — not easier work for the weaker child, but a "
    "different door into the same room.",
    extra=[("The manual's six levers:  ", GREEN, True, False),
           ("varied formats  ·  flexible deadlines  ·  varying task difficulty  ·  accommodations "
            "·  individualised feedback  ·  learner involvement.", GREY, False, False)],
    accent=GREEN,
    notes_text="DIFFERENTIATED ASSESSMENT (4 min)\n\n"
    "This is where teachers most often go wrong, and where learners get short-changed. Many "
    "teachers believe differentiation means giving the 'weak ones' easier work — which quietly "
    "lowers the standard and widens the gap.\n\n"
    "DRIVE THIS HOME: 'Differentiate the route, not the destination.' If the learning outcome "
    "says the learner must analyse data, then eventually they must analyse data — no matter how "
    "we help them get there.\n\n"
    "Walk through the six levers on the slide. Ask teachers which one they already do naturally, "
    "and which one they have never tried.\n\n"
    "Connect back: the manual places this alongside GESI, SEL and SEN. Differentiation is not "
    "an add-on for special cases; it is the normal professional response to a real classroom, "
    "where thirty learners are never at the same point.")

# ---------------------------------------------------------------- 5. SEVEN PRINCIPLES
slide_section(4, "The Seven\nPillars",
              "The principles that hold up every good assessment practice.",
              img="05_pillars.jpg",
              notes_text="PART 4 (10 min)\n\n"
              "Section 2 of the manual lists seven principles of effective assessment. This is "
              "the part people skim — and it is the part that explains why some assessments work "
              "and others do not.\n\n"
              "Frame it as seven pillars holding up a roof. If one is missing, the building leans. "
              "You do not need all seven perfect; you need all seven present.")

slide_grid_cards(
    "Seven Pillars — Hold Up the Roof",
    "Everything that follows rests on these",
    "05_pillars.jpg",
    [("1 \u00b7 Validity", "Am I measuring the RIGHT thing?"),
     ("2 \u00b7 Reliability", "Do I get the SAME result every time, whoever marks it?"),
     ("3 \u00b7 Fairness and ethics", "Does every learner get an EQUITABLE chance?"),
     ("4 \u00b7 Transparency", "Do learners know WHAT they are judged on and HOW?"),
     ("5 \u00b7 Inclusivity", "Is it accessible to ALL learners, without exception?"),
     ("6 \u00b7 Practicability", "Can I ACTUALLY do this with my time and my class size?"),
     ("7 \u00b7 Assessment utility", "Is it USEFUL — does anybody learn anything from it?")],
    accent=GREEN,
    img_caption="Seven pillars. Remove one, and the roof leans.",
    note_text="THE SEVEN PILLARS (2 min overview)\n\n"
    "Read all seven quickly to give the shape of the section — do not dwell yet, each one gets "
    "its own slide next.\n\n"
    "Ask the room to guess which pillar is most often broken in real practice. Common answers: "
    "transparency and practicability. Both are true — we keep the criteria secret and then design "
    "assessments nobody has time to mark.\n\n"
    "Promise: 'By the end of the next ten minutes you will recognise which of these you already "
    "do well, and which one is quietly costing your learners marks.'")

slide_concept(
    "1 \u00b7 Validity — Measuring the Right Thing",
    "The pillar most often broken without anyone noticing",
    "Are you actually measuring the thing you say you are measuring? A valid assessment measures "
    "the important learning outcomes of the curriculum — not just whatever was easiest to set.",
    "A trader uses a weighing scale to sell yam. If she used that scale to measure the LENGTH of "
    "the yam, the reading would be worthless — not because the scale is broken, but because it "
    "is measuring the wrong thing. Now think of your last class test. Your learning outcome said "
    "learners should ANALYSE data. But every question you set asked them to RECALL a definition. "
    "Your test was neat. It was well marked. It was simply measuring the wrong thing. That is "
    "an invalid assessment — and the learners who studied hard still lost marks for it.",
    extra=[("The teacher's quick validity check:  ", GREEN, True, False),
           ("\u201cDoes this question actually test the skill I said I was teaching?\u201d  "
            "If the answer is no, the marks it produces mean very little.", GREY, False, False)],
    notes_text="VALIDITY (3 min)\n\n"
    "THE TRAP: validity fails silently. An invalid test still produces marks, still fills the "
    "mark book, still looks professional. Nobody complains. But the numbers do not mean what "
    "everyone thinks they mean.\n\n"
    "THE CLEAREST TEST: compare your learning outcome to your questions. If the outcome says "
    "'analyse' and your questions say 'state', you have an invalid assessment. Ask teachers to "
    "silently do this check on the last test they set. Most will find at least one problem.\n\n"
    "THE FIX — and this is the highest-value practical tip in this section: build a TABLE OF "
    "SPECIFICATION. A simple grid mapping each question to a topic and a thinking level. The "
    "manual asks teachers to 'create a learning and assessment plan (i.e. table of test "
    "specification)'. Without it, we all default to recall questions, because recall questions "
    "are the easiest to write. The grid forces balance.\n\n"
    "Also mention the other validity requirements from the manual: clearly state the purpose of "
    "the assessment, define the performance criteria up front, and score against those criteria "
    "to avoid bias and stereotyping.")

slide_image_text(
    "2 \u00b7 Reliability — The Same Result Every Time",
    "Reliability is about trust in the numbers",
    "06_scales.jpg",
    [("Plain English",
      "Do you get the same result every time, whoever is doing the marking? Reliable results are "
      "dependable enough to make decisions with."),
     ("The classroom picture",
      "A market scale that gives the same weight for the same pile of tomatoes — on Monday or "
      "Saturday, whoever is standing at the stall. Now imagine the scale was generous on Monday "
      "and harsh on Saturday. Nobody would trust that market again."),
     ("Where we break it",
      "If this term's 30 marks in your subject are generously given and next term's are harsh, "
      "learners are being ranked against a ruler that keeps changing length. That is unreliable."),
     ("How the manual says to fix it",
      "Clear learning outcomes · a colleague reviews your marking · more than one assessment "
      "method for the same outcome · proper marking schemes with weightings · rubrics given "
      "OUT before the task · a calm, suitable room.")],
    img_side="right",
    accent=GREEN,
    img_caption="Same pile, same reading — every time.",
    note_text="RELIABILITY (3 min)\n\n"
    "OPEN WITH THE MANUAL'S OWN WARNING — it is blunt and it stings: the 'connoisseur' approach "
    "to assessment — that is, 'I know it when I see it but I cannot put it into words' — is NOT "
    "acceptable. Read that line out loud. It is a direct challenge to the way many of us mark.\n\n"
    "WHY IT MATTERS PRACTICALLY: unreliable marking is unfair in a way learners can feel but not "
    "prove. A learner in your class and a learner in the class next door are compared as if they "
    "were marked the same way. If they were not, the comparison is fiction.\n\n"
    "THE MOST POWERFUL HABIT on this slide: give your rubric or marking scheme OUT BEFORE the "
    "task, not after. The manual says this explicitly for performance and practical assessment.\n\n"
    "CHEAP RELIABILITY HACK to offer teachers: swap a set of scripts with a colleague in the same "
    "subject and mark each other's sample. The manual recommends having another teacher review "
    "the work. You will be shocked how differently two people read the same answer.")

slide_concept(
    "3 \u00b7 Fairness and Ethics — An Equitable Chance",
    "Fairness is not treating everyone identically — it is giving everyone a real chance",
    "Assessment strategies should give learners an equitable opportunity to demonstrate what they "
    "know and can do, taking into account their ability, learning styles, gender and Special "
    "Educational Needs. It also means never assessing what you have not taught.",
    "A test question asks learners to explain the rules of cricket. A learner from a village that "
    "has never played cricket — and never seen it — cannot answer, and fails. That question did "
    "not measure physics. It measured cricket. Now swap it for a question about the physics of a "
    "loaded trotro braking at a junction, and every learner in the room has something to think "
    "with. Same physics. Same difficulty. Fair.",
    extra=[("The manual's red flags:  ", TERRA, True, False),
           ("culturally biased content  ·  unfamiliar words and examples  ·  assessing content "
            "you never taught  ·  failing to give SEN learners extra time or alternative formats.",
            GREY, False, False)],
    accent=TERRA,
    notes_text="FAIRNESS AND ETHICS (3 min)\n\n"
    "START WITH THE CRICKET EXAMPLE. It is memorable and it is exactly the failure mode the "
    "manual warns against — 'culturally biased or discriminatory content, unfamiliar words, "
    "questioning or examples'.\n\n"
    "THEN ASK THE HARD QUESTION: 'How many of us have set questions using examples from books "
    "written far from here, about things our learners have never seen?' Let that sit. The fix is "
    "not easier questions — it is LOCAL, FAMILIAR CONTEXTS for the same rigour. That is what the "
    "trotro example shows.\n\n"
    "THE SECOND BIG IDEA on this slide: do not assess what you have not taught. The manual lists "
    "this under fairness, and it is the single most common fairness complaint learners make.\n\n"
    "Also cover the logistics side, which teachers forget counts as fairness: communicate the "
    "date, time, venue and format in advance, and give clear instructions about what is expected.")

slide_concept(
    "4 \u00b7 Transparency — No Secret Exams",
    "If learners have to guess what you want, you have already made it harder than it needs to be",
    "Making the assessment process and criteria clear and understandable to learners. They should "
    "know what is being judged, how it is judged, and what counts as a pass.",
    "Two teachers set the same project. The first teacher says \u201cdo the project, I will mark "
    "it\u201d — and the learners spend the week guessing. The second teacher hands out the rubric "
    "on day one: what is expected, the marks for each part, what a top answer looks like. Both "
    "teachers mark with the same rubric. The second class almost always scores higher — not "
    "because they worked harder, but because they knew where to aim. Assessment should not be a "
    "guessing game you are the referee of.",
    extra=[("Transparency also requires you to:  ", GREEN, True, False),
           ("share the learning outcomes being assessed  ·  tell learners the pass mark  ·  share "
            "results openly with learners and parents/guardians  ·  give learners a route to seek "
            "review and redress if they disagree with a score.", GREY, False, False)],
    accent=GREEN,
    notes_text="TRANSPARENCY (3 min)\n\n"
    "THE KILLER LINE: 'Assessment should not be a guessing game where the teacher is the only "
    "person who knows the answer.'\n\n"
    "THE MOST COMMON OFFENCE: hiding the rubric until after the marking, usually to 'stop them "
    "cheating'. Point out that a learner who aims perfectly at a published rubric is not cheating. "
    "They are learning. That is exactly what we want.\n\n"
    "THE REDRESS POINT is the one nobody expects. The manual requires you to 'provide an "
    "opportunity for learners to seek review and redress'. Ask: 'If a learner in this school "
    "disputes a mark, what actually happens? Is there a route?' In many schools there is none. "
    "That is a gap worth naming.\n\n"
    "LINK BACK: transparency and validity are partners. Learners cannot demonstrate a learning "
    "outcome they were never told about.")

slide_concept(
    "5 \u00b7 Inclusivity — Fair and Accessible to ALL",
    "Not an extra duty. It is a standing requirement of the framework",
    "Creating assessment practices that are fair and accessible to every learner — regardless of "
    "gender, disability, poverty, background or learning difference. The framework names three "
    "lenses you must apply: GESI, SEL and SEN.",
    "A learner with low vision in your class cannot read an ordinary small-print paper — but they "
    "understand the work perfectly. A learner who is hard of hearing misses the spoken "
    "instructions you gave once, quickly, at the front. A learner who is dyslexic knows the "
    "science but loses marks to their handwriting. In all three cases we did not measure their "
    "learning. We measured their eyes, their ears and their handwriting. Inclusivity means "
    "removing that barrier so the mark reflects the learning.",
    extra=[("GESI", GREEN, True, False), (" = Gender Equality and Social Inclusion.   ", GREY, False, False),
           ("SEL", GREEN, True, False), (" = Socio-Emotional Learning.   ", GREY, False, False),
           ("SEN", GREEN, True, False), (" = Special Educational Needs.", GREY, False, False)],
    accent=GREEN,
    notes_text="INCLUSIVITY (4 min)\n\n"
    "THE KEY REFRAME: 'When a learner with low vision fails your test, you have not discovered "
    "that they do not know the work. You have discovered that they cannot see your paper. Those "
    "are completely different findings, and only one of them is about learning.'\n\n"
    "DECODE THE THREE ACRONYMS — teachers see these letters constantly and rarely have them "
    "explained:\n"
    "  · GESI — Gender Equality and Social Inclusion. No learner disadvantaged by gender, "
    "disability, poverty or background.\n"
    "  · SEL — Socio-Emotional Learning. The learner's feelings, confidence and relationships "
    "matter and are assessed, not just their marks.\n"
    "  · SEN — Special Educational Needs. Learners needing specific support.\n\n"
    "THE MANUAL'S PRACTICAL TOOLS to name out loud: braille, oral translation, text-to-speech "
    "and other AI supports, sign language interpretation, assistive technology, extra time, "
    "alternative formats.\n\n"
    "A WARNING WORTH GIVING: build fairness into your RUBRIC. If you are assessing science "
    "understanding, do not dock marks for handwriting or grammar. The manual specifically says "
    "to develop rubrics that consider grammar, vocabulary, handwriting and presentation — "
    "decide deliberately, do not do it by accident.\n\n"
    "Mention the manual points teachers to the inclusivity section of the NPLAF (page 32) for "
    "further reading.")

slide_concept(
    "6 \u00b7 Practicability — Can You Actually Do It?",
    "The pillar that kills the most beautiful assessment plans",
    "The assessment must be feasible, convenient, efficient and successful with the real "
    "resources you have — your actual class size, your actual timetable, your actual materials.",
    "You design a wonderful one-hour one-on-one oral assessment for your Government class. It is "
    "rigorous. It is valid. It is inclusive. It is also completely impossible with 60 learners — "
    "and by Thursday it will have collapsed into chaos, you will be three weeks behind, and you "
    "will abandon it. The plan failed not because it was bad, but because it was not practicable. "
    "A good assessment is one you can still run in week eight, when you are tired and the term "
    "is crowded.",
    extra=[("The three practicability questions:  ", TERRA, True, False),
           ("Do I have the materials and the security?  ·  Does this format suit my class size "
            "and their level?  ·  Do I have the TIME to set it, mark it, AND give constructive "
            "feedback?", GREY, False, False)],
    accent=TERRA,
    notes_text="PRACTICABILITY (3 min)\n\n"
    "OPEN WITH THE HONEST OBSERVATION: 'We have all designed a beautiful assessment in January "
    "that we quietly stopped running by March.' This slide is about why that happens — and it is "
    "not a character flaw. It is a design flaw.\n\n"
    "THE THREE QUESTIONS on the slide are the whole test. Walk through them for a real example. "
    "A 20-item test: markable in one evening, feedback possible. A one-hour-per-learner oral: "
    "mathematically impossible.\n\n"
    "THE POINT TEACHERS MISS: set, mark AND give feedback. Many assessments are practicable to "
    "administer but NOT practicable to give feedback on. Since feedback is where the learning "
    "actually happens, an assessment you cannot give feedback on has already failed.\n\n"
    "VALIDATE THEIR REALITY: yes, class sizes are large and time is short. That is exactly why "
    "practicability is written into the national framework as a PRINCIPLE, not an excuse. The "
    "framework expects you to choose assessment methods that fit your actual conditions.\n\n"
    "PRACTICAL TIP: for large classes, use assessment strategies that scale — whole-class "
    "response, peer assessment with a rubric, exit cards, group projects with individual "
    "accountability.")

slide_concept(
    "7 \u00b7 Assessment Utility — Is It Actually Useful?",
    "The question that separates data collection from real assessment",
    "The assessment must be useful. It must give you and the learner information worth having, "
    "and that information must lead to action.",
    "You set a class test. You mark it. You enter the scores in the mark book. You file the "
    "papers. Term ends. Now ask honestly: did anybody learn anything from that test? If the "
    "test did not change what you taught next, and did not change what the learner did next, "
    "then all that happened was data collection. Utility is the difference between a test that "
    "teaches and a test that merely records.",
    extra=[("Utility in practice:  ", GREEN, True, False),
           ("state the intended use of the result up front  ·  produce credible information for "
            "learners and parents  ·  give constructive feedback  ·  and then ACT on it — re-teach, "
            "regroup, intervene, adapt.", GREY, False, False)],
    accent=GREEN,
    notes_text="ASSESSMENT UTILITY (3 min)\n\n"
    "THIS IS THE SLIDE THAT TIES THE WHOLE SEVEN TOGETHER. You can satisfy validity, reliability, "
    "fairness, transparency, inclusivity and practicability — and still produce an assessment "
    "that nobody uses. That is the most common failure of all, because it is invisible.\n\n"
    "THE TEST TO OFFER: 'Name one thing you changed last term because of an assessment result.' "
    "Give people a moment. Ask two or three to answer out loud. The answers show what utility "
    "looks like in practice.\n\n"
    "LINK BACK to the soup: a cook who tastes the soup and then does nothing about it has wasted "
    "the taste. Tasting without adjusting is not cooking. Assessing without acting is not "
    "assessment.\n\n"
    "CLOSE THE SECTION: 'Seven pillars. Validity, reliability, fairness, transparency, "
    "inclusivity, practicability, utility. You will not get all seven perfect every week. But "
    "when an assessment disappoints you, come back to this list — the answer is almost always "
    "hiding in one of the seven.'")

# ---------------------------------------------------------------- 6. JARGON BUSTER
slide_section(5, "The Jargon\nBuster",
              "Every technical term, in plain words, with a picture.",
              img="08_dok_ladder.jpg",
              notes_text="PART 5 (8 min)\n\n"
              "This is the section teachers came for. Slow down here. This is the reference part "
              "of the workshop — the part they will photograph or write down.\n\n"
              "Note: the full Jargon Buster booklet contains every term. On the slides we cover "
              "the ones that cause the most trouble and the ones used most often.")

slide_image_text(
    "Rubric — The Recipe Card",
    "The single most useful document in your assessment toolkit",
    "10_rubric_card.jpg",
    [("Plain English",
      "A table showing exactly what counts and how many marks each part is worth, at each level "
      "of quality."),
     ("The classroom picture",
      "It is a recipe card with measurements. Not \u201cadd enough salt\u201d — but \u201cone "
      "teaspoon\u201d. Not \u201cwrite a good essay\u201d — but \u201cthree clear arguments, each "
      "supported by one example, maximum two pages.\u201d"),
     ("Why it changes everything",
      "It removes the guesswork for the learner AND the bias from the marker. It is the fastest "
      "route to both transparency and reliability at the same time."),
     ("The rule nobody follows",
      "Give the rubric OUT BEFORE the task — never after. The manual says this explicitly for "
      "performance and practical assessment.")],
    img_side="left",
    accent=GOLD,
    img_caption="A recipe card with measurements, not guesses.",
    note_text="RUBRIC (3 min)\n\n"
    "THE ANALOGY TO STRESS: 'Add enough salt' is not a recipe. 'One teaspoon' is a recipe. Most "
    "of our instructions to learners are the first kind — and then we are surprised when thirty "
    "learners produce thirty different things.\n\n"
    "THE DOUBLE BENEFIT: a rubric improves BOTH transparency and reliability simultaneously. It "
    "tells the learner where to aim, and it stops two markers giving wildly different scores to "
    "the same script. It is the best-value document in this whole workshop.\n\n"
    "THE RULE TO REPEAT: handout BEFORE the task. A rubric revealed after marking is not "
    "transparency. It is a post-mortem.\n\n"
    "PRACTICAL OFFER: if teachers have never built a rubric, the manual contains exemplar rubrics "
    "in Appendix C — including for project-based assessment. There is no need to start from a "
    "blank page.")

slide_table(
    "Depth of Knowledge — The Staircase of Thinking",
    "How DEEP is the thinking? Not how difficult the words look.",
    ["Level", "Name", "Plain English", "SHS example — Social Studies"],
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
    intro="DoK stands for Depth of Knowledge. It is a ladder, not a difficulty rating — the words can stay simple while the thinking goes deep.",
    row_h=0.78, fsize=11,
    note_text="DEPTH OF KNOWLEDGE (4 min)\n\n"
    "THE BIGGEST MISCONCEPTION, and the reason this slide exists: teachers think a question is "
    "DoK 4 because it uses big words or long sentences. It does not. DoK measures the TYPE of "
    "thinking required.\n\n"
    "SAY THIS CLEARLY: 'A long question with big vocabulary that asks you to remember something "
    "is still DoK 1. A short, simple question that asks you to judge a real situation and defend "
    "your judgement can be DoK 3.'\n\n"
    "WALK UP THE STAIRCASE using the climate change example on the slide — it is drawn directly "
    "from the ministry's own curriculum resource, so teachers will meet it again.\n\n"
    "WHY IT MATTERS FOR VALIDITY: if your learning outcome is at DoK 3 ('analyse', 'evaluate') "
    "and your test questions are all at DoK 1 ('state', 'list', 'define'), your assessment is "
    "invalid — no matter how well marked. The manual uses Depth of Knowledge levels throughout "
    "the curriculum's key assessment tasks.\n\n"
    "PRACTICAL EXERCISE, if time allows: ask teachers to take one question from their last test "
    "and label it. Most will find their whole paper sits at DoK 1 and 2. That is a diagnosis, "
    "not a criticism — and it is easy to fix with a table of specification.")

slide_table(
    "Jargon Buster — The Everyday Terms",
    "The words we use constantly but rarely define",
    ["Term", "Plain English", "Classroom picture"],
    [["Diagnostic assessment", "Checking prior knowledge before teaching.", "Checking the pantry before you cook."],
     ["Formative assessment", "Checking during teaching so you can adjust.", "Tasting the soup while it is on the fire."],
     ["Summative assessment", "A final summary at the end, for certification.", "Serving the meal at the table."],
     ["Differentiated assessment", "Different routes to prove the same outcome.", "Some draw the diagram, some write the essay — same outcome."],
     ["Performance / authentic", "Learners actually DO it, in a real-life way.", "Not describing a titration — standing at the bench doing it."],
     ["Portfolio", "A curated collection showing progress over time.", "Not a folder of everything — chosen pieces, with reflection."],
     ["Learning analytics", "Spotting patterns in results to act early.", "18 out of 45 failed the same question — that is a whole-class gap."],
     ["Metacognition", "Thinking about your own thinking.", "\u201cI understand the theory but I keep making arithmetic slips.\u201d"],
     ["e-Assessment", "Using digital tools to assess.", "A class WhatsApp voice-note submission; a Google Form quiz."]],
    col_w=[2.45, 4.55, 5.09],
    row_h=0.46, fsize=11.5,
    note_text="JARGON BUSTER (4 min)\n\n"
    "Do NOT read this table line by line — it is a reference slide. Pick the three or four that "
    "you know your staff find hardest, and dwell on those.\n\n"
    "RECOMMENDED EMPHASES:\n"
    "  · 'Metacognition' — nobody ever explains it. It simply means thinking about your own "
    "thinking: knowing how you learn and where you get stuck. It is the heart of Assessment as "
    "Learning.\n"
    "  · 'Learning analytics' — sounds like software, but the example shows it is just careful "
    "attention to patterns in your mark book. You can do it with a pen and a class list.\n"
    "  · 'Portfolio' — the misconception is that a portfolio is a folder of everything. It is a "
    "curated selection WITH reflection. The reflection is where the learning is.\n"
    "  · 'e-Assessment' — teachers assume it means a computer lab. It does not. A WhatsApp "
    "voice note is e-assessment.\n\n"
    "POINT TO THE HANDOUT: 'Every term we cover is in the Jargon Buster booklet with the same two "
    "columns — a plain meaning and a classroom picture. Take it with you.'")

# ---------------------------------------------------------------- 7. FEEDBACK AND TOOLS
slide_section(6, "Feedback —\nWhere Learning Happens",
              "The most powerful assessment tool you already own.",
              img="07_feedback_loop.jpg",
              notes_text="PART 6 (6 min)\n\n"
              "Every assessment in this workshop exists to produce one thing: feedback that "
              "changes what the learner does next. This is the payoff section.\n\n"
              "It is also the section where teachers realise they are doing more work than they "
              "need to, for less effect than they want.")

slide_diagram_text(
    "Feedback — The Good Mechanic",
    "Why \u201cgood work, 7/10\u201d is not feedback",
    [("Assessment happens", "You gather evidence — a test, an exercise, a question in class."),
     ("You give SPECIFIC feedback", "Not \u201cgood work\u201d. Point at the exact part: \u201cthis step is wrong, because\u2026\u201d"),
     ("The learner acts on it", "This needs TIME in the next lesson. No time, no feedback."),
     ("Learning improves", "The gap closes — and you can see that it closed.")],
    [("The useless mechanic",
      "Says \u201cthe car is bad\u201d and takes your money. You leave knowing nothing, and the "
      "car is no better."),
     ("The good mechanic",
      "Opens the bonnet, points at ONE exact part and says: \u201cThis hose is cracked. Replace it "
      "and the overheating stops.\u201d You now know what to do next."),
     ("Plain English",
      "Feedback is information that helps a learner close the gap between where they are and "
      "where they need to be."),
     ("What it is NOT",
      "It is not a score. It is not praise. \u201cGood work, 7/10\u201d tells a learner nothing "
      "about what to do differently. A mark is a verdict; feedback is a map."),
     ("The rule teachers break",
      "Feedback without TIME TO ACT ON IT is decoration. If you hand back corrections and "
      "immediately move to the next topic, the feedback dies in the exercise book.")],
    diagram_side="right",
    accent=TERRA,
    caption="Feedback only works if the learner has time to act on it.",
    note_text="FEEDBACK (4 min)\n\n"
    "THE MECHANIC ANALOGY is the heart of this slide. Everyone has met both mechanics. Ask the "
    "room which one they would return to, and why. Then make the turn: 'Which mechanic are we to "
    "our learners?'\n\n"
    "THE HARD TRUTH: most of what we call feedback is a number plus a word. 'Good work, 7/10' "
    "gives the learner a verdict, not a next step. The learner knows they got 7. They do not know "
    "what to do to get 9.\n\n"
    "THE MOST ACTIONABLE RULE — and the one teachers find genuinely relieving: feedback needs "
    "TIME. Ten minutes at the start of the next lesson for learners to read and act on your "
    "corrections is worth more than two hours of extra marking. Without that time, all your "
    "careful marking is decoration.\n\n"
    "THE MANUAL'S TERMS to use: 'constructive feedback' — it points to the next step, not just "
    "the fault. And note the manual's instruction that learners must be 'guided to develop how "
    "to use or respond to feedback to improve learning'. Using feedback is itself a skill we have "
    "to teach.\n\n"
    "SELF-CHECK TO OFFER: 'When you hand back a marked script, does the learner do anything with "
    "it other than look at the score?'")

slide_diagram_text(
    "Three-Minute Tools You Can Use on Monday",
    "No budget. No printing. No extra marking load.",
    [("Pick ONE tool", "Just one. Not all five. Choose the one that suits your subject."),
     ("Give it 3 minutes", "Exit card at the door, entry ticket at the start, or a quick Think-Pair-Share."),
     ("Read what you collect", "Ten slips tell you where the whole class is stuck."),
     ("Start tomorrow from there", "That is the whole loop — assessment feeding teaching.")],
    [("Exit card",
      "Last three minutes of the lesson: on a slip, \u201cwrite one thing you learned today and "
      "one thing you are still confused about.\u201d Collect at the door. You now know exactly "
      "where to start next lesson."),
     ("Entry ticket",
      "First three minutes: one question that shows what they already bring. Free diagnostic "
      "assessment, every single day."),
     ("Think-Pair-Share",
      "Think alone, then discuss with a partner, then share with the class. Three minutes of real "
      "thinking time means the quiet learners in the back have something to say when you ask."),
     ("K-W-L chart",
      "Three columns: what I Know, what I Want to know, what I Learned. Makes thinking visible "
      "and makes the learner reflective."),
     ("Self- and peer-assessment",
      "Learners check their own or each other's work AGAINST THE RUBRIC. Only works if they have "
      "the criteria in hand — otherwise it becomes a popularity contest.")],
    diagram_side="left",
    accent=GOLD,
    caption="Three minutes. No budget. No extra marking load.",
    note_text="THREE-MINUTE TOOLS (4 min)\n\n"
    "THIS IS THE 'TAKE IT HOME' SLIDE. Its job is to prove that formative assessment is not extra "
    "work. Every tool here costs three minutes and no money.\n\n"
    "PICK ONE AND COMMIT THE ROOM. Do not ask them to adopt all five. Ask each teacher to choose "
    "ONE tool to try in their very next lesson and, if possible, tell the person next to them "
    "which one they chose. Public commitment increases follow-through dramatically.\n\n"
    "WHY EXIT CARDS ARE THE BEST STARTING POINT: they take three minutes, they need only scrap "
    "paper, they work with 60 learners, and they tell you precisely where to begin tomorrow. If "
    "you only ever adopt one thing from today, adopt the exit card.\n\n"
    "THE PEER-ASSESSMENT WARNING is worth saying: peer assessment fails when learners do not have "
    "the rubric. The manual is clear that learners should assess 'against specified criteria'. "
    "Give them the criteria, and train them to be specific and kind.\n\n"
    "NOTE ON THE MANUAL: it lists Think-Pair-Share, exit cards, K-W-L and peer assessment among "
    "the curriculum's formative strategies — these are not extra inventions.")

# ---------------------------------------------------------------- 8. PLANNING
slide_section(7, "Planning It\nProperly",
              "Turning all of this into a plan you can actually follow.",
              img="09_plan_journey.jpg",
              notes_text="PART 7 (5 min)\n\n"
              "Everything so far has been about principles and individual strategies. This section "
              "is about putting them together into a plan for a term — so assessment stops being "
              "improvised and starts being designed.\n\n"
              "This is the difference between a teacher who assesses and a teacher who plans "
              "assessment.")

slide_image_text(
    "The Six-Phase Assessment Plan",
    "The manual's model, in plain words",
    "09_plan_journey.jpg",
    [("Phase 1 \u00b7 Learning outcome",
      "What should the learner be able to DO? Everything else flows from this one sentence. If "
      "you cannot state it, you cannot assess it."),
     ("Phase 2 \u00b7 Assessment strategies",
      "HOW will I find out? Not just a test — observation, project, oral, practical, portfolio. "
      "Choose the method that matches the outcome."),
     ("Phase 3 \u00b7 Tasks or questions",
      "The actual questions and tasks learners will meet. This is where validity is won or lost."),
     ("Phase 4 \u00b7 Grading criteria or rubrics",
      "What 'good' looks like, at each level, decided BEFORE the learners start — not while you "
      "are marking at midnight."),
     ("Phase 5 \u00b7 Timeline and sequencing",
      "When will each assessment happen, and how does it fit the term? Spread the load so week "
      "eight is not a disaster."),
     ("Phase 6 \u00b7 Feedback and reporting",
      "How and when will learners act on the feedback — and how does it reach parents?"),
    ],
    img_side="left",
    accent=GREEN,
    img_caption="A plan is a journey with milestones, not a surprise at the end.",
    note_text="THE SIX-PHASE PLAN (4 min)\n\n"
    "This is the manual's own assessment plan model, translated into plain words. Walk through "
    "the phases in order — they are sequential for a reason.\n\n"
    "THE CRITICAL INSIGHT: Phase 4 (rubrics) comes BEFORE the learners start. Many teachers "
    "write the rubric while marking, which means the standard drifts as they go — this is the "
    "single biggest cause of unreliable marking. Decide the standard first, then mark against it.\n\n"
    "PHASE 5 IS THE ONE THAT PROTECTS YOU. Notice how often teachers accidentally schedule "
    "three projects, two practicals and a test in the same fortnight. Sequencing assessment "
    "across the term is a professional skill, not a luxury. Show the term calendar and ask: "
    "'Where are our pressure points?'\n\n"
    "PHASE 6 IS THE ONE WE SKIP. If there is no planned time for feedback, there will be no "
    "feedback — only marks.\n\n"
    "PRACTICAL ASK: 'Before you leave today, take ONE topic you will teach in the next month and "
    "write its Phase 1 learning outcome on a piece of paper. That is the beginning of a real "
    "assessment plan.'\n\n"
    "A note of realism: the manual also says to be ready to adapt. A plan is a map, not a "
    "straitjacket.")

slide_concept(
    "The 70 / 30 Split — Your Marks Now Have Weight",
    "Why school-based assessment is no longer soft data",
    "Thirty per cent of a learner's final grade comes from work done in your school and marked by "
    "you — portfolios, performance and project work, plus end-of-term examinations. Seventy per "
    "cent comes from WAEC's final examination.",
    "For years, internal marks were treated as a formality at the back of the mark book. That has "
    "changed. The internal 30% now sits on a school-based transcript, with far greater "
    "transparency and quality assurance. Practically: if your school's 30% is not credible, "
    "transparent and defensible, the learner is the one who suffers when their transcript is "
    "scrutinised — and the school is the one that is exposed. Reliable rubrics are no longer "
    "good practice. They are protection.",
    extra=[("Why reliability suddenly matters so much:  ", TERRA, True, False),
           ("the same rubric, applied the same way, across every class in the school, is what "
            "makes 30% of a young person's future defensible.", GREY, False, False)],
    accent=TERRA,
    notes_text="THE 70/30 SPLIT (3 min)\n\n"
    "THIS IS THE SLIDE THAT CHANGES BEHAVIOUR. Principles are persuasive; accountability is "
    "compelling.\n\n"
    "THE HISTORY, briefly: for years, continuing assessment marks were often treated casually — "
    "generous scoring, lost record books, reconstructed marks at the end of term. The reform "
    "means the 30% is now recorded on a school-based transcript with greater transparency and "
    "quality assurance.\n\n"
    "THE CONSEQUENCE, stated plainly: 'If a learner's transcript is questioned and your 30% "
    "cannot be defended, it is the learner who loses a place — and your school whose credibility "
    "is damaged.'\n\n"
    "THE SOLUTION is not more paperwork. It is the simple disciplines from today: a rubric "
    "written BEFORE the task, applied consistently, kept on record, and the same across all "
    "classes teaching the same subject.\n\n"
    "ASK THE SCHOOL-LEVEL QUESTION: 'Do all of us teaching the same subject at the same level "
    "use the same criteria? Could we produce the evidence if NaSIA asked?' This is the moment to "
    "raise consistency as a departmental issue, not just an individual one.")

# ---------------------------------------------------------------- 9. E-ASSESSMENT
slide_section(8, "Assessment and\nTechnology",
              "e-Assessment — including the versions that work without a computer lab.",
              img="01_hero_classroom.jpg",
              notes_text="PART 8 (3 min)\n\n"
              "Keep this section short and practical. The single goal is to break the assumption "
              "that e-assessment requires a computer lab, reliable power and school-wide Wi-Fi.\n\n"
              "If your school has those things, wonderful. If it does not, you can still do "
              "e-assessment tomorrow.")

slide_diagram_text(
    "e-Assessment — You Probably Already Do It",
    "Technology for assessing, not just teaching",
    [("Learner submits", "A 60-second WhatsApp voice note, a photo of their practical work, or an online quiz."),
     ("The tool marks \u2014 or you listen", "A Google Form marks itself; you listen to voice notes on your way home."),
     ("Feedback goes back fast", "The learner sees their score or hears your comment in minutes, not weeks."),
     ("You see the pattern", "Everyone failed the same question \u2014 a whole-class gap you can now fix.")],
    [("Plain English",
      "Using information technology and digital tools to assess learners — to design, deliver, "
      "score, analyse, report and manage assessment."),
     ("The myth to kill",
      "e-Assessment does NOT require a computer lab. If your school has one phone per group and "
      "a WhatsApp class group, you can do e-assessment this week."),
     ("Low-tech, high-value examples",
      "A class WhatsApp group where learners submit a 60-second voice note explaining a concept  "
      "\u00b7  a Google Form quiz marked automatically  \u00b7  a Kahoot round at the end of a "
      "lesson  \u00b7  photographs of practical work submitted for feedback."),
     ("Where it genuinely excels",
      "Immediate feedback to learners  \u00b7  automatic marking that saves your evenings  \u00b7  "
      "digital portfolios that show progress over a whole programme  \u00b7  analytics that reveal "
      "whole-class gaps."),
     ("Two cautions from the manual",
      "Make provision for learners with SEN who need extra support with the platform, and ensure "
      "the integrity and security of the platform to prevent cheating and cyberbullying.")],
    diagram_side="right",
    accent=GREEN,
    caption="One phone per group is enough to start.",
    note_text="e-ASSESSMENT (3 min)\n\n"
    "LEAD WITH THE MYTH-BUSTER: most teachers hear 'e-assessment' and switch off, thinking of "
    "computer labs and unreliable internet. The manual's definition is simply 'the use of "
    "information technology and digital tools to assess learners' achievement'. A phone is "
    "technology.\n\n"
    "THE MOST PRACTICAL IDEA ON THIS SLIDE is the WhatsApp voice note. A learner records sixty "
    "seconds explaining a concept. You listen while walking home. You hear their actual reasoning "
    "— which a written answer often hides — and you can respond with a voice note in thirty "
    "seconds. That is genuinely powerful formative assessment with tools every learner already "
    "has.\n\n"
    "THE QUALITY-OF-LIFE POINT: automatic marking. A Google Form quiz marks itself and can email "
    "the learner their score instantly. For large classes, this returns hours of your week.\n\n"
    "THE MANUAL'S EXAMPLES to name if asked: online quizzes and tests, digital portfolios, "
    "simulations and virtual labs, discussion boards and blogs, e-open-book assessments, "
    "multimedia projects, learning analytics, digital rubrics.\n\n"
    "DO NOT SKIP THE CAUTIONS. The manual is explicit that teachers must make provision for "
    "learners with SEN who may need extra support with the platform, and must ensure integrity, "
    "security and ethical use — to prevent cheating and cyberbullying. Say both out loud; they "
    "are the questions a thoughtful staff will ask.")

# ---------------------------------------------------------------- 10. CLOSE
slide_concept(
    "The Whole Workshop on One Page",
    "If you remember nothing else, remember this",
    "Assessment is not the exam at the end. It is the nervous system of your teaching — it tells "
    "you what is working while you can still do something about it.",
    "BEFORE you teach, you check what they bring — that is DIAGNOSTIC. WHILE you teach, you check "
    "and adjust — that is FORMATIVE, and you taste the soup. AT THE END you summarise — that is "
    "SUMMATIVE, and you serve the meal. ALL THE WAY THROUGH, you make it valid, reliable, fair, "
    "transparent, inclusive, practicable and useful.",
    extra=[("And the two sentences that carry it all:  ", GREEN, True, False),
           ("the learner must know what \u201cgood\u201d looks like — and must be given time to "
            "get there.", INK, True, False)],
    accent=GREEN,
    notes_text="THE CLOSE (3 min)\n\n"
    "This is the emotional and intellectual landing point. Slow right down.\n\n"
    "DELIVER THE THREE BEATS CLEARLY:\n"
    "  BEFORE — diagnostic — you check what they bring.\n"
    "  WHILE — formative — you taste the soup.\n"
    "  AT THE END — summative — you serve the meal.\n"
    "  THROUGHOUT — the seven pillars hold up the roof.\n\n"
    "THEN THE TWO SENTENCES. These are the whole workshop compressed:\n"
    "  1. The learner must know what 'good' looks like. (transparency — the rubric before the task)\n"
    "  2. The learner must have time to get there. (feedback with time to act on it)\n\n"
    "If a teacher leaves today remembering only those two sentences, the workshop has succeeded.")



def slide_closing():
    s = new()
    picture_cover(s, "01_hero_classroom.jpg", 0, 0, SW, SH)
    from pptx.oxml.ns import qn
    ov = rect(s, 0, 0, SW, SH, GREEN_DK)
    solid_el = ov.fill._xPr.find(qn('a:solidFill'))
    clr = solid_el.find(qn('a:srgbClr'))
    clr.append(clr.makeelement(qn('a:alpha'), {'val': '88000'}))
    rect(s, 0, 0, SW, 0.13, GOLD)
    tf = textbox(s, M + 0.15, 1.55, CW - 0.3, 4.4)
    para(tf, "Three Things to Take Home", 34, WHITE, bold=True, font=HEAD_FONT,
         first=True, space_after=22)
    para(tf, "1.   Pick ONE topic you will teach this month. Write its learning outcome in one "
             "sentence. That is Phase 1 of a real assessment plan.", 16,
         RGBColor(0xE4, 0xF0, 0xEB), space_after=14, line_spacing=1.12)
    para(tf, "2.   Try ONE three-minute tool — an exit card, an entry ticket, a quick "
             "Think-Pair-Share. Just one, in your next lesson.", 16,
         RGBColor(0xE4, 0xF0, 0xEB), space_after=14, line_spacing=1.12)
    para(tf, "3.   Take ONE end-of-term paper this year and give it back for group correction "
             "instead of only recording the marks.", 16,
         RGBColor(0xE4, 0xF0, 0xEB), space_after=24, line_spacing=1.12)
    para(tf, "Thank you.  \u2014  Now let us hear from you:  what will you try first?",
         19, GOLD, bold=True, font=HEAD_FONT, space_after=6)
    footer(s, dark=True)
    notes(s, "CLOSING AND DISCUSSION (3-5 min)\n\n"
             "Do NOT end by summarising the whole workshop again — you have just done that. End "
             "with a COMMITMENT and a CONVERSATION.\n\n"
             "Read the three take-homes slowly. Then ask each teacher to turn to the person beside "
             "them and say which ONE of the three they will actually do. Ninety seconds only. "
             "Spoken commitment to a peer dramatically increases follow-through compared with a "
             "general exhortation from the front.\n\n"
             "THEN OPEN THE ROOM. Good questions to invite:\n"
             "  · Which of the seven pillars do you think our school breaks most often?\n"
             "  · Do all of us teaching the same subject use the same criteria? Should we?\n"
             "  · What would it take for us to share rubrics as a department?\n"
             "  · Is there a learner in your class who is being assessed on their eyes, their "
             "ears or their handwriting rather than their learning?\n\n"
             "CLOSE WITH ENCOURAGEMENT, NOT GUILT. The message is: 'You are already assessing "
             "every day. Today was about doing it deliberately, so it costs you less and gives "
             "your learners more.'\n\n"
             "Point to the two handouts — the Jargon Buster booklet and the PDF of these slides — "
             "and remind them the exemplars and rubrics live in Appendix C of the NaCCA Teacher "
             "Assessment Manual and Toolkit, which is available on the NaCCA curriculum resources "
             "portal.")
    return s


slide_closing()

prs.save(OUT)
print("saved:", OUT)
print("slides:", len(prs.slides._sldIdLst))
