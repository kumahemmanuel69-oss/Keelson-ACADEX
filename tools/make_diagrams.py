#!/usr/bin/env python3
"""Draw the teaching diagrams for the KATON deck.

These are charts with real labels, so they are drawn with PIL rather than
generated as AI images — an AI illustration would garble the numbers and the
words, and these diagrams have to be exactly right on a projector.

SIZING RULE — read before editing
---------------------------------
Each diagram is displayed about 11.2 in wide on the slide. On a 2400 px canvas
that is 2400 / (11.2 * 72) = roughly 3 px per point. So:

    smallest font that is still readable on a projector = 33 px
    body text                                            = 34-38 px
    headings                                             = 44-52 px

Text below about 30 px becomes unreadable from the back of a hall. Do not add
detail by shrinking type — cut content instead.

Output: assets/diagrams/*.png  (2400 x 1000)
"""
import os
import math
from PIL import Image, ImageDraw, ImageFont

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "..", "assets", "diagrams")

GREEN    = (11, 110, 79)
GREEN_DK = (7, 74, 53)
GREEN_LT = (227, 241, 234)
GREEN_MD = (198, 226, 212)
GOLD     = (244, 180, 0)
GOLD_DK  = (146, 102, 0)
GOLD_LT  = (253, 243, 216)
TERRA    = (193, 68, 14)
TERRA_LT = (251, 233, 224)
INK      = (20, 38, 46)
GREY     = (90, 107, 114)
CREAM    = (253, 248, 240)
WHITE    = (255, 255, 255)

REG = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
BLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"

W, H = 2400, 1000


def F(sz, bold=False):
    return ImageFont.truetype(BLD if bold else REG, sz)


def wrap(draw, text, fnt, maxw):
    words, lines, cur = text.split(), [], ""
    for w in words:
        t = (cur + " " + w).strip()
        if draw.textlength(t, font=fnt) <= maxw:
            cur = t
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def para(d, x, y, text, fnt, fill, maxw, lh=None):
    lh = lh or int(fnt.size * 1.24)
    for ln in wrap(d, text, fnt, maxw):
        d.text((x, y), ln, font=fnt, fill=fill)
        y += lh
    return y


def centre(d, cx, y, text, fnt, fill):
    w = d.textlength(text, font=fnt)
    d.text((cx - w / 2, y), text, font=fnt, fill=fill)
    return y + fnt.size


def rrect(d, box, r, fill=None, outline=None, width=3):
    d.rounded_rectangle(box, radius=r, fill=fill, outline=outline, width=width)


# --------------------------------------------------------------------------- 1
def dartboard():
    """Validity x reliability, the classic four-panel target."""
    im = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(im)

    cw, ch = W // 2, H // 2
    panels = [
        (0, 0, "VALID AND RELIABLE", GREEN,
         "On the centre, every time.",
         "The right thing, measured well.", "centre"),
        (1, 0, "RELIABLE, NOT VALID", TERRA,
         "Tidy \u2014 and off target.",
         "Consistently wrong. Looks competent.", "offcentre"),
        (0, 1, "VALID, NOT RELIABLE", GOLD_DK,
         "Right target, no repeatability.",
         "The average is fine. One mark is not.", "spread"),
        (1, 1, "NEITHER", GREY,
         "Scattered everywhere.",
         "No use for marks, feedback or planning.", "wild"),
    ]

    for col, row, title, accent, head, body, kind in panels:
        x0, y0 = col * cw, row * ch
        cx, cy = x0 + 232, y0 + ch // 2
        R = 178

        d.ellipse([cx - R, cy - R, cx + R, cy + R], fill=WHITE, outline=GREY, width=4)
        d.ellipse([cx - R * 0.70, cy - R * 0.70, cx + R * 0.70, cy + R * 0.70],
                  fill=CREAM, outline=GREY, width=3)
        d.ellipse([cx - R * 0.40, cy - R * 0.40, cx + R * 0.40, cy + R * 0.40],
                  fill=WHITE, outline=GREY, width=3)
        d.ellipse([cx - 10, cy - 10, cx + 10, cy + 10],
                  fill=GREEN if kind == "centre" else GOLD)

        if kind == "centre":
            off = [(-14, -10), (12, -16), (-8, 14), (16, 8), (0, -2), (6, -26)]
            bx, by = cx, cy
        elif kind == "offcentre":
            off = [(74, -70), (108, -88), (88, -46), (122, -58), (66, -94), (100, -28)]
            bx, by = cx + 14, cy - 14
        elif kind == "spread":
            off = [(0, -100), (-100, -22), (96, -18), (0, 92), (-72, 54),
                   (78, 46), (-26, -50), (30, 20)]
            bx, by = cx, cy
        else:
            off = [(-146, -98), (136, -110), (-106, 120), (132, 98),
                   (18, -136), (-32, 6), (88, -36), (-146, 62), (150, 6)]
            bx, by = cx, cy

        for dx, dy in off:
            hx, hy = bx + dx, by + dy
            d.ellipse([hx - 12, hy - 12, hx + 12, hy + 12], fill=INK)
            d.ellipse([hx - 4, hy - 4, hx + 4, hy + 4], fill=WHITE)

        tx = x0 + 452
        tw = cw - 482
        yy = y0 + 118
        d.text((tx, yy), title, font=F(38, True), fill=accent)
        yy += 56
        yy = para(d, tx, yy, head, F(33, True), INK, tw)
        yy += 6
        para(d, tx, yy, body, F(32), GREY, tw, lh=42)

    d.line([0, ch, W, ch], fill=(228, 233, 231), width=3)
    d.line([cw, 0, cw, H], fill=(228, 233, 231), width=3)

    im.save(os.path.join(OUT, "dartboard.png"))
    print("  dartboard.png")


# --------------------------------------------------------------------------- 2
def dok_staircase():
    """Depth of Knowledge as four rising steps."""
    im = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(im)

    base = 894
    steps = [
        ("DoK 1", "Recall", "Give back a fact.",
         "\u201cState 3 causes of climate change.\u201d", GREEN_LT, GREEN_DK, 330),
        ("DoK 2", "Skills", "Use it in a routine task.",
         "\u201cExplain the greenhouse effect.\u201d", GREEN_MD, GREEN_DK, 470),
        ("DoK 3", "Strategic", "Reason, justify, plan.",
         "\u201cWhy is the farmer\u2019s yield falling?\u201d", GOLD_LT, GOLD_DK, 610),
        ("DoK 4", "Extended", "Investigate, create, combine.",
         "\u201cSurvey waste disposal for a week.\u201d", TERRA_LT, TERRA, 750),
    ]

    sw = W // 4
    for i, (lvl, name, plain, ex, fill, fg, top) in enumerate(steps):
        x0 = i * sw + 20
        x1 = (i + 1) * sw - 20
        rrect(d, [x0, base - top, x1, base], 18, fill=fill)
        rrect(d, [x0, base - top, x1, base - top + 13], 7, fill=fg)

        cy = base - top + 40
        d.text((x0 + 24, cy), lvl, font=F(52, True), fill=fg)
        cy += 68
        d.text((x0 + 24, cy), name, font=F(36, True), fill=INK)
        cy += 52
        cy = para(d, x0 + 24, cy, plain, F(32), GREY, sw - 54, lh=42)
        cy += 10
        para(d, x0 + 24, cy, ex, F(31), fg, sw - 54, lh=41)

    d.line([70, 58, W - 150, 58], fill=GOLD, width=9)
    d.polygon([(W - 150, 36), (W - 150, 80), (W - 96, 58)], fill=GOLD)
    d.text((70, 4), "DEEPER THINKING \u2014 not harder words",
           font=F(34, True), fill=GOLD_DK)

    d.line([0, base + 10, W, base + 10], fill=GREY, width=6)
    d.text((40, base + 32), "DoK 1\u20132: the floor of every paper",
           font=F(33, True), fill=GREY)
    t = "DoK 3\u20134: where WASSCE rewards you"
    d.text((W - d.textlength(t, font=F(33, True)) - 40, base + 32),
           t, font=F(33, True), fill=TERRA)

    im.save(os.path.join(OUT, "dok_staircase.png"))
    print("  dok_staircase.png")


# --------------------------------------------------------------------------- 3
def assessment_cycle():
    """The five-step loop, drawn as a closed circle."""
    im = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(im)

    cx, cy, R = 690, 500, 322
    nodes = ["PLAN", "TEACH", "ASSESS", "ANALYSE", "ACT"]
    n = len(nodes)

    d.ellipse([cx - R, cy - R, cx + R, cy + R], outline=GREEN_LT, width=28)

    pts = []
    for i in range(n):
        a = -math.pi / 2 + i * 2 * math.pi / n
        pts.append((cx + R * math.cos(a), cy + R * math.sin(a)))

    for i in range(n):
        a1 = -math.pi / 2 + i * 2 * math.pi / n
        a2 = -math.pi / 2 + (i + 1) * 2 * math.pi / n
        am = (a1 + a2) / 2
        ax, ay = cx + R * math.cos(am), cy + R * math.sin(am)
        tangent = am + math.pi / 2
        L, Wd = 48, 32
        tipx, tipy = ax + L * math.cos(tangent), ay + L * math.sin(tangent)
        d.polygon([
            (tipx, tipy),
            (tipx - Wd * math.cos(tangent - 0.5), tipy - Wd * math.sin(tangent - 0.5)),
            (tipx - Wd * math.cos(tangent + 0.5), tipy - Wd * math.sin(tangent + 0.5)),
        ], fill=GOLD)

    for i, label in enumerate(nodes):
        px, py = pts[i]
        r = 90
        d.ellipse([px - r, py - r, px + r, py + r], fill=GREEN, outline=WHITE, width=7)
        f = F(30 if len(label) > 6 else 34, True)
        w = d.textlength(label, font=f)
        d.text((px - w / 2, py - 18), label, font=f, fill=WHITE)

    centre(d, cx, cy - 116, "Assessment", F(44, True), INK)
    centre(d, cx, cy - 62, "is a CYCLE", F(44, True), GREEN)
    centre(d, cx, cy + 6, "not an event", F(34), GREY)

    x = 1180
    rrect(d, [x, 96, W - 60, 356], 18, fill=GREEN_LT)
    rrect(d, [x, 96, x + 14, 356], 6, fill=GREEN)
    d.text((x + 40, 128), "Most schools stop too early",
           font=F(36, True), fill=GREEN_DK)
    para(d, x + 40, 190, "We PLAN, TEACH and ASSESS \u2014 then the marks go in "
                         "the book and nothing changes.",
         F(33), INK, W - x - 130, lh=44)
    para(d, x + 40, 286, "That is record-keeping, not assessment.",
         F(33, True), TERRA, W - x - 130, lh=44)

    rrect(d, [x, 396, W - 60, 656], 18, fill=GOLD_LT)
    rrect(d, [x, 396, x + 14, 656], 6, fill=GOLD)
    d.text((x + 40, 428), "The two steps we skip", font=F(36, True), fill=GOLD_DK)
    para(d, x + 40, 490, "ANALYSE \u2014 look for the pattern, not the person.",
         F(33), INK, W - x - 130, lh=44)
    para(d, x + 40, 552, "ACT \u2014 re-teach, regroup, or move on.",
         F(33), INK, W - x - 130, lh=44)

    rrect(d, [x, 696, W - 60, 906], 18, fill=WHITE, outline=TERRA, width=3)
    d.text((x + 40, 726), "The rule on timing", font=F(36, True), fill=TERRA)
    para(d, x + 40, 788, "The gap between ASSESS and ACT should be DAYS.",
         F(33, True), INK, W - x - 130, lh=44)

    im.save(os.path.join(OUT, "cycle.png"))
    print("  cycle.png")


# --------------------------------------------------------------------------- 4
def split_70_30():
    """The donut."""
    im = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(im)

    cx, cy, ro, ri = 560, 500, 372, 210
    box = [cx - ro, cy - ro, cx + ro, cy + ro]
    end = -90 + 360 * 0.30
    d.arc(box, -90, end, fill=GREEN, width=ro - ri)
    d.arc(box, end, 270, fill=GOLD, width=ro - ri)

    centre(d, cx, cy - 92, "100%", F(74, True), INK)
    centre(d, cx, cy + 4, "the learner's", F(34), GREY)
    centre(d, cx, cy + 48, "final grade", F(34), GREY)

    mid_r = (ro + ri) / 2

    def band_label(angle_deg, big, small, fg):
        a = math.radians(angle_deg)
        lx, ly = cx + mid_r * math.cos(a), cy + mid_r * math.sin(a)
        w1 = d.textlength(big, font=F(40, True))
        w2 = d.textlength(small, font=F(24))
        d.text((lx - w1 / 2, ly - 30), big, font=F(40, True), fill=fg)
        d.text((lx - w2 / 2, ly + 16), small, font=F(24), fill=fg)

    band_label(-90 + 360 * 0.30 / 2, "30%", "internal", WHITE)
    band_label(-90 + 360 * 0.30 + 360 * 0.70 / 2, "70%", "external", GOLD_DK)

    x = 1020
    rrect(d, [x, 100, W - 50, 350], 18, fill=GREEN_LT)
    rrect(d, [x, 100, x + 14, 350], 6, fill=GREEN)
    d.text((x + 40, 132), "30%  \u2014  MARKED BY YOUR SCHOOL",
           font=F(37, True), fill=GREEN_DK)
    para(d, x + 40, 200, "Portfolios, project and performance work, plus "
                         "end-of-term examinations.",
         F(33), INK, W - x - 120, lh=44)
    para(d, x + 40, 306, "Recorded on a school-based transcript.",
         F(33, True), GREEN_DK, W - x - 120, lh=44)

    rrect(d, [x, 390, W - 50, 640], 18, fill=GOLD_LT)
    rrect(d, [x, 390, x + 14, 640], 6, fill=GOLD)
    d.text((x + 40, 422), "70%  \u2014  SET AND MARKED BY WAEC",
           font=F(37, True), fill=GOLD_DK)
    para(d, x + 40, 490, "The final WASSCE paper. Every learner in the country "
                         "sits it on the same day.",
         F(33), INK, W - x - 120, lh=44)

    rrect(d, [x, 680, W - 50, 930], 18, fill=CREAM, outline=TERRA, width=3)
    d.text((x + 40, 710), "WHY THIS CHANGES YOUR MARKING",
           font=F(36, True), fill=TERRA)
    para(d, x + 40, 774, "The 70 is fixed. The 30 is the part you control \u2014 "
                         "and it now follows the learner.",
         F(33), INK, W - x - 120, lh=44)

    im.save(os.path.join(OUT, "split_70_30.png"))
    print("  split_70_30.png")


# --------------------------------------------------------------------------- 5
def transcript():
    """A mock school-based transcript, deliberately kept sparse so it reads."""
    im = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(im)

    cols = [("SUBJECT", 760), ("INTERNAL\n(30)", 400), ("WASSCE\n(70)", 400),
            ("FINAL\n(100)", 400), ("GRADE", 400)]
    rows = [
        ("Core Mathematics", "24", "52", "76", "B2"),
        ("English Language", "25", "48", "73", "B3"),
        ("Integrated Science", "20", "41", "61", "C4"),
        ("Social Studies", "22", "45", "67", "C5"),
        ("Elective Mathematics", "26", "55", "81", "A1"),
    ]

    x0, y0 = 40, 128
    row_h, hdr_h = 96, 132
    total_w = sum(c[1] for c in cols)

    d.text((x0, 20), "SCHOOL-BASED TRANSCRIPT \u2014 a worked example",
           font=F(38, True), fill=GREEN_DK)

    rrect(d, [x0, y0, x0 + total_w, y0 + hdr_h], 0, fill=GREEN_DK)
    x = x0
    for name, wid in cols:
        for i, ln in enumerate(name.split("\n")):
            f = F(31, True)
            tw = d.textlength(ln, font=f)
            d.text((x + wid / 2 - tw / 2, y0 + 34 + i * 38), ln, font=f, fill=WHITE)
        x += wid

    for r, row in enumerate(rows):
        yy = y0 + hdr_h + r * row_h
        bg = WHITE if r % 2 == 0 else CREAM
        d.rectangle([x0, yy, x0 + total_w, yy + row_h], fill=bg)
        x = x0
        for c, val in enumerate(row):
            bold = c in (0, 1, 3, 4)
            f = F(36, bold)
            col = INK
            if c == 1:
                col = GREEN
            if c == 3:
                col = GREEN_DK
            if c == 4:
                col = GREEN if val.startswith("A") else INK
            if c == 0:
                d.text((x + 20, yy + 30), val, font=f, fill=col)
            else:
                tw = d.textlength(val, font=f)
                d.text((x + cols[c][1] / 2 - tw / 2, yy + 30), val, font=f, fill=col)
            x += cols[c][1]
        d.line([x0, yy, x0 + total_w, yy], fill=(226, 231, 229), width=2)

    yy = y0 + hdr_h + len(rows) * row_h
    d.line([x0, yy, x0 + total_w, yy], fill=GREEN_DK, width=5)

    d.text((x0, yy + 26), "The internal 30 is built from portfolio + project/practical + end-of-term exam.",
           font=F(31), fill=GREY)
    d.text((x0, yy + 68), "The question an inspector will ask:  \u201cShow me the work behind this 24.\u201d",
           font=F(33, True), fill=GREEN_DK)

    im.save(os.path.join(OUT, "transcript.png"))
    print("  transcript.png")


def main():
    os.makedirs(OUT, exist_ok=True)
    print("drawing diagrams into %s" % os.path.normpath(OUT))
    dartboard()
    dok_staircase()
    assessment_cycle()
    split_70_30()
    transcript()
    print("done \u2014 smallest type is 24 px (band labels only), body 32 px+")
    print("body text renders at about %.1f pt on the slide" % (32 * 11.18 / 2400 * 72))


if __name__ == "__main__":
    main()
