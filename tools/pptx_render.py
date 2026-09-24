#!/usr/bin/env python3
"""
A small PPTX -> PNG renderer, good enough to proof the deck we generate ourselves.

Handles: solid-fill autoshapes (rect / rounded rect / oval), textboxes with runs
(size, bold, italic, colour, alignment), pictures with crop, and z-order.
Not a general-purpose renderer, but accurate for this build.
"""
import os, sys, zipfile
from lxml import etree
from PIL import Image, ImageDraw, ImageFont

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
EMU_IN = 914400
SCALE = 140                      # px per inch in the proof
# DejaVu is wider than Calibri/Trebuchet; compress so line-wrapping matches Office
WIDTH_FUDGE = 0.86

FONTS = {
    (False, False): '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
    (True,  False): '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf',
    (False, True):  '/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf',
    (True,  True):  '/usr/share/fonts/truetype/dejavu/DejaVuSans-BoldOblique.ttf',
}
_fc = {}


def font(size_pt, bold, italic):
    key = (round(size_pt, 1), bold, italic)
    if key not in _fc:
        path = FONTS.get((bold, italic)) or FONTS[(False, False)]
        if not os.path.exists(path):
            path = FONTS[(False, False)]
        _fc[key] = ImageFont.truetype(path, max(6, int(size_pt * WIDTH_FUDGE * SCALE / 72.0)))
    return _fc[key]


def emu(v):
    return int(v) / EMU_IN * SCALE


def rgb_of(el):
    c = el.find('.//a:srgbClr', NS)
    if c is not None:
        h = c.get('val')
        return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))
    return None


def alpha_of(el):
    a = el.find('.//a:alpha', NS)
    return int(a.get('val')) / 100000.0 if a is not None else 1.0


def wrap(draw, text, fnt, max_w):
    """Greedy word wrap using real pixel widths."""
    words, lines, cur = text.split(), [], ''
    for w in words:
        t = (cur + ' ' + w).strip()
        if draw.textlength(t, font=fnt) <= max_w or not cur:
            cur = t
        else:
            lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def text_width(draw, text, fnt):
    return draw.textlength(text, font=fnt)


class Renderer:
    def __init__(self, pptx):
        self.z = zipfile.ZipFile(pptx)
        import re as _re
        self.slides = sorted(
            (n for n in self.z.namelist()
             if n.startswith('ppt/slides/slide') and n.endswith('.xml')),
            key=lambda n: int(_re.search(r'slide(\d+)\.xml$', n).group(1))
        )
        self.slide_no = {n: int(_re.search(r'slide(\d+)\.xml$', n).group(1))
                         for n in self.slides}
        rels = etree.fromstring(self.z.read('ppt/_rels/presentation.xml.rels'))
        self.relmap = {r.get('Id'): r.get('Target') for r in rels}
        pres = etree.fromstring(self.z.read('ppt/presentation.xml'))
        sz = pres.find('.//p:sldSz', NS)
        self.W = int(emu(sz.get("cx")))
        self.H = int(emu(sz.get("cy")))

    def slide_rels(self, idx):
        """media path -> target for slide idx (1-based)."""
        rp = f'ppt/slides/_rels/slide{idx}.xml.rels'
        out = {}
        if rp in self.z.namelist():
            r = etree.fromstring(self.z.read(rp))
            for rel in r:
                out[rel.get('Id')] = rel.get('Target')
        return out

    def render(self, idx, out_png):
        img = Image.new('RGB', (self.W, self.H), (255, 255, 255))
        d = ImageDraw.Draw(img)
        name = self.slides[idx - 1]
        root = etree.fromstring(self.z.read(name))
        rels = self.slide_rels(self.slide_no[name])
        tree = root.find('.//p:cSld/p:spTree', NS)
        for sp in tree:
            tag = etree.QName(sp).localname
            if tag == 'sp':
                self.draw_shape(img, d, sp)
            elif tag == 'pic':
                self.draw_pic(img, sp, rels)
        img.save(out_png)
        return out_png

    # ----------------------------------------------------------- shapes
    def xfrm(self, sp):
        x = sp.find('.//a:xfrm', NS)
        if x is None:
            return None
        off, ext = x.find('a:off', NS), x.find('a:ext', NS)
        if off is None or ext is None:
            return None
        return (emu(off.get('x')), emu(off.get('y')),
                emu(ext.get('cx')), emu(ext.get('cy')))

    def draw_shape(self, img, d, sp):
        box = self.xfrm(sp)
        if not box:
            return
        x, y, w, h = box
        geom = sp.find('.//a:prstGeom', NS)
        kind = geom.get('prst') if geom is not None else 'rect'
        spPr = sp.find('.//p:spPr', NS)
        fill = None
        if spPr is not None:
            sf = spPr.find('a:solidFill', NS)
            if sf is not None:
                fill = rgb_of(sf)
        # geometry
        if fill:
            if kind == 'ellipse':
                d.ellipse([x, y, x + w, y + h], fill=fill)
            elif kind == 'roundRect':
                r = min(w, h) * 0.10
                d.rounded_rectangle([x, y, x + w, y + h], radius=r, fill=fill)
            else:
                d.rectangle([x, y, x + w, y + h], fill=fill)
        self.draw_text(d, sp, (x, y, w, h))

    def draw_text(self, d, sp, box):
        tx = sp.find('.//p:txBody', NS)
        if tx is None:
            return
        x, y, w, h = box
        # margins
        ml = mb = mr = mt = 0.0
        bp = tx.find('a:bodyPr', NS)
        if bp is not None:
            ml = emu(bp.get('lIns') or 91440)
            mr = emu(bp.get('rIns') or 91440)
            mt = emu(bp.get('tIns') or 45720)
            mb = emu(bp.get('bIns') or 45720)
        # vertical anchor
        anch = bp.get('anchor') if bp is not None else None
        avail_w = max(10, (w - ml - mr))
        paras = tx.findall('a:p', NS)
        # pre-measure for middle-anchoring
        blocks = []
        for p in paras:
            runs = p.findall('a:r', NS)
            if not runs:
                blocks.append(([], 1.0))
                continue
            algn = p.find('a:pPr', NS)
            algn = algn.get('algn') if algn is not None else 'l'
            blocks.append((runs, algn))
        rendered = []
        total_h = 0
        for runs, algn in blocks:
            items = []
            for r in runs:
                t = r.find('a:t', NS)
                if t is None or t.text is None:
                    continue
                rPr = r.find('a:rPr', NS)
                sz = float(rPr.get('sz') or 1800) / 100.0 if rPr is not None else 18.0
                bold = rPr is not None and rPr.get('b') == '1'
                ital = rPr is not None and rPr.get('i') == '1'
                col = rgb_of(rPr) if rPr is not None else None
                items.append((t.text, sz, bold, ital, col or (0, 0, 0)))
            sa = p.find('a:pPr', NS)
            spc_aft = 0.0
            if sa is not None:
                aft = sa.find('a:spcAft', NS)
                if aft is not None:
                    pts = aft.find('a:spcPts', NS)
                    if pts is not None and pts.get('val'):
                        spc_aft = float(pts.get('val')) / 100.0
            ln_spc = 1.0
            if sa is not None:
                ls = sa.find('a:lnSpc', NS)
                if ls is not None and ls.find('a:spcPct', NS) is not None:
                    ln_spc = float(ls.find('a:spcPct', NS).get('val')) / 100000.0
            # flow all runs of this paragraph inline, then wrap to the box width
            toks = []                      # (word, font, colour, size)
            for text, sz, bold, ital, col in items:
                fnt = font(sz, bold, ital)
                for j, piece in enumerate(text.split('\n')):
                    if j:
                        toks.append(('\n', fnt, col, sz))
                    for wd in piece.split():
                        toks.append((wd, fnt, col, sz))
            lines = []
            cur, cur_w = [], 0.0
            base = None
            for wd, fnt, col, sz in toks:
                if wd == '\n':
                    lines.append((cur, cur_w, base)); cur, cur_w, base = [], 0.0, None
                    continue
                if base is None:
                    base = sz
                wpx = d.textlength((wd + ' '), font=fnt)
                if cur and cur_w + wpx > avail_w:
                    lines.append((cur, cur_w, base))
                    cur, cur_w = [(wd, fnt, col, sz)], d.textlength(wd, font=fnt)
                    base = sz
                else:
                    cur.append((wd, fnt, col, sz))
                    cur_w += wpx
            if cur:
                lines.append((cur, cur_w, base))
            # convert to render-ready lines
            flat = []
            for words, _, bsz in lines:
                bsz = bsz or 12
                lh = bsz * SCALE / 72.0 * 1.18 * ln_spc
                flat.append((words, lh, bsz, algn))
            lines = flat
            rendered.append((lines, spc_aft))
            total_h += sum(l[1] for l in lines) + spc_aft * SCALE / 72.0
        if os.environ.get('PPTX_QA'):
            if total_h > 1.5 and total_h > (h - mt - mb) + 1.5:
                t = ' | '.join(
                    ''.join(rr.find('a:t', NS).text or '' for rr in rs)[:46]
                    for rs, _ in blocks if rs)
                print(f'  OVERFLOW slide? box_h={h:.0f} text_h={total_h:.0f} :: {t[:110]}')
        cy = y + mt
        if anch == 'ctr':
            cy = y + (h - total_h) / 2.0
        elif anch == 'b':
            cy = y + h - mb - total_h
        for lines, spc_aft in rendered:
            for words, lh, bsz, algn in lines:
                if not words:
                    cy += lh
                    continue
                sw = sum(d.textlength(w + ' ', font=f) for w, f, _, _ in words)
                sw -= d.textlength(' ', font=words[-1][1])
                if algn == 'ctr':
                    sx = x + ml + (avail_w - sw) / 2
                elif algn == 'r':
                    sx = x + ml + avail_w - sw
                else:
                    sx = x + ml
                yy = cy + (lh - bsz * SCALE / 72.0 * 1.18) / 2
                for w, f, col, _ in words:
                    d.text((sx, yy), w, font=f, fill=col)
                    sx += d.textlength(w + ' ', font=f)
                cy += lh
            cy += spc_aft * SCALE / 72.0

    # ----------------------------------------------------------- pictures
    def draw_pic(self, img, sp, rels):
        box = self.xfrm(sp)
        if not box:
            return
        x, y, w, h = box
        blip = sp.find('.//a:blip', NS)
        if blip is None:
            return
        rid = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
        target = rels.get(rid)
        if not target:
            return
        path = 'ppt/' + target.replace('../', '')
        try:
            im = Image.open(self.z.open(path)).convert('RGB')
        except Exception as e:
            print('  !! image failed', path, e)
            return
        # apply crop
        src = sp.find('.//a:srcRect', NS)
        cl = ct = cr = cb = 0.0
        if src is not None:
            cl = int(src.get('l') or 0) / 100000.0
            ct = int(src.get('t') or 0) / 100000.0
            cr = int(src.get('r') or 0) / 100000.0
            cb = int(src.get('b') or 0) / 100000.0
        iw, ih = im.size
        l = int(cl * iw); t = int(ct * ih)
        r = int((1 - cr) * iw); b = int((1 - cb) * ih)
        im = im.crop((l, t, r, b))
        img.paste(im.resize((int(w), int(h)), Image.LANCZOS), (int(x), int(y)))


if __name__ == '__main__':
    pptx = sys.argv[1]
    outdir = sys.argv[2]
    pages = [int(a) for a in sys.argv[3:]] if len(sys.argv) > 3 else None
    os.makedirs(outdir, exist_ok=True)
    r = Renderer(pptx)
    n = len(r.slides)
    for i in range(1, n + 1):
        if pages and i not in pages:
            continue
        p = r.render(i, os.path.join(outdir, f'slide-{i:02d}.png'))
        print(p)
    print('total slides:', n)
