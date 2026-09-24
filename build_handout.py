#!/usr/bin/env python3
"""
Build the companion PDF handout for the KATON 2026 assessment workshop.
Designed as the document colleagues keep and re-read after the session.
"""
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph,
                                Spacer, Table, TableStyle, Image, PageBreak,
                                KeepTogether, NextPageTemplate)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

ROOT = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(ROOT, "assets", "opt")
OUT = os.path.join(ROOT, "deliverables", "Assessment-Handout-KATON-2026.pdf")

# ------------------------------------------------------------------ fonts
FD = "/usr/share/fonts/truetype/dejavu"
pdfmetrics.registerFont(TTFont("Body", f"{FD}/DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont("Body-B", f"{FD}/DejaVuSans-Bold.ttf"))
# no oblique DejaVu on this box; Bitstream Vera is DejaVu's parent face, so it matches
RAL = "/usr/local/lib/python3.11/dist-packages/reportlab/fonts"
pdfmetrics.registerFont(TTFont("Body-I", f"{RAL}/VeraIt.ttf"))
pdfmetrics.registerFont(TTFont("Body-BI", f"{RAL}/VeraBI.ttf"))
pdfmetrics.registerFontFamily("Body", normal="Body", bold="Body-B",
                              italic="Body-I", boldItalic="Body-BI")

# ------------------------------------------------------------------ palette
GREEN = colors.HexColor("#0B6E4F")
GREEN_DK = colors.HexColor("#074A35")
GREEN_LT = colors.HexColor("#E3F1EA")
GOLD = colors.HexColor("#F4B400")
GOLD_LT = colors.HexColor("#FDF3D8")
TERRA = colors.HexColor("#C1440E")
TERRA_LT = colors.HexColor("#FBE9E0")
INK = colors.HexColor("#14262E")
GREY = colors.HexColor("#5A6B72")
PAPER = colors.HexColor("#F7F9F8")
CREAM = colors.HexColor("#FDF8F0")

PW, PH = A4
MARG = 18 * mm
CW = PW - 2 * MARG

# ------------------------------------------------------------------ styles
S = {}
S['h1'] = ParagraphStyle('h1', fontName="Body-B", fontSize=19, leading=24,
                         textColor=GREEN_DK, spaceBefore=2, spaceAfter=3)
S['kicker'] = ParagraphStyle('kicker', fontName="Body-B", fontSize=8.5, leading=11,
                             textColor=TERRA, spaceAfter=2)
S['h2'] = ParagraphStyle('h2', fontName="Body-B", fontSize=13.5, leading=17,
                         textColor=GREEN_DK, spaceBefore=12, spaceAfter=5)
S['h3'] = ParagraphStyle('h3', fontName="Body-B", fontSize=11, leading=14,
                         textColor=TERRA, spaceBefore=9, spaceAfter=3)
S['body'] = ParagraphStyle('body', fontName="Body", fontSize=9.6, leading=13.6,
                           textColor=INK, spaceAfter=5, alignment=TA_LEFT)
S['small'] = ParagraphStyle('small', fontName="Body", fontSize=8.4, leading=11.6,
                            textColor=GREY, spaceAfter=4)
S['bullet'] = ParagraphStyle('bullet', parent=S['body'], leftIndent=11,
                             bulletIndent=2, spaceAfter=3.5)
S['cell'] = ParagraphStyle('cell', fontName="Body", fontSize=8.6, leading=11.4,
                           textColor=INK)
S['cellb'] = ParagraphStyle('cellb', fontName="Body-B", fontSize=8.6, leading=11.4,
                            textColor=GREEN_DK)
S['cellh'] = ParagraphStyle('cellh', fontName="Body-B", fontSize=8.8, leading=11.6,
                            textColor=colors.white)
S['quote'] = ParagraphStyle('quote', fontName="Body-I", fontSize=10.2, leading=14.4,
                            textColor=GREEN_DK, leftIndent=9, spaceBefore=3,
                            spaceAfter=6)
S['cover_t'] = ParagraphStyle('ct', fontName="Body-B", fontSize=32, leading=38,
                              textColor=colors.white)
S['cover_s'] = ParagraphStyle('cs', fontName="Body-B", fontSize=15, leading=21,
                              textColor=GOLD)
S['cover_b'] = ParagraphStyle('cb', fontName="Body", fontSize=10.5, leading=15,
                              textColor=colors.HexColor("#D8E8E1"))
S['toc'] = ParagraphStyle('toc', fontName="Body", fontSize=10, leading=15.5,
                          textColor=INK)


def P(t, st='body'):
    return Paragraph(t, S[st])


def bullets(items, style='bullet'):
    return [Paragraph(t, S[style], bulletText='\u2022') for t in items]


def rule(color=GOLD, w=CW, h=1.6, space=6):
    t = Table([['']], colWidths=[w], rowHeights=[h])
    t.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, -1), color),
                           ('LEFTPADDING', (0, 0), (-1, -1), 0),
                           ('RIGHTPADDING', (0, 0), (-1, -1), 0),
                           ('TOPPADDING', (0, 0), (-1, -1), 0),
                           ('BOTTOMPADDING', (0, 0), (-1, -1), 0)]))
    return [t, Spacer(1, space)]


def band(title, sub=None, color=GREEN):
    """Coloured heading band."""
    flow = [Paragraph(title, ParagraphStyle('bt', fontName="Body-B", fontSize=12,
                                            leading=15, textColor=colors.white))]
    if sub:
        flow.append(Paragraph(sub, ParagraphStyle('bs', fontName="Body", fontSize=8.6,
                                                  leading=11,
                                                  textColor=colors.HexColor("#D8E8E1"))))
    t = Table([[flow]], colWidths=[CW])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), color),
        ('LEFTPADDING', (0, 0), (-1, -1), 9), ('RIGHTPADDING', (0, 0), (-1, -1), 9),
        ('TOPPADDING', (0, 0), (-1, -1), 6), ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    return [t, Spacer(1, 8)]


def datatable(headers, rows, widths=None, fd_size=8.6):
    n = len(headers)
    widths = widths or [CW / n] * n
    data = [[Paragraph(h, S['cellh']) for h in headers]]
    for r in rows:
        data.append([Paragraph(c, S['cellb'] if i == 0 else S['cell'])
                     for i, c in enumerate(r)])
    t = Table(data, colWidths=widths, repeatRows=1)
    st = [('BACKGROUND', (0, 0), (-1, 0), GREEN_DK),
          ('VALIGN', (0, 0), (-1, -1), 'TOP'),
          ('LEFTPADDING', (0, 0), (-1, -1), 7),
          ('RIGHTPADDING', (0, 0), (-1, -1), 7),
          ('TOPPADDING', (0, 0), (-1, -1), 5.5),
          ('BOTTOMPADDING', (0, 0), (-1, -1), 5.5),
          ('LINEBELOW', (0, 0), (-1, -2), 0.4, colors.HexColor("#DFE7E4")),
          ('BOX', (0, 0), (-1, -1), 0.6, colors.HexColor("#CBD8D3"))]
    for i in range(1, len(data)):
        if i % 2 == 0:
            st.append(('BACKGROUND', (0, i), (-1, i), PAPER))
    t.setStyle(TableStyle(st))
    return t


def two_col(items, lw, rw, gap=8):
    """items = [(left_flow, right_flow)] -> aligned two-column rows."""
    data = [[a, b] for a, b in items]
    t = Table(data, colWidths=[lw, rw])
    t.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))
    return t


def concept(title, meaning, picture, extra=None):
    """The signature block: term + plain meaning + classroom picture."""
    flow = [Paragraph(title, S['h3'])]
    box_rows = [[Paragraph("<b>PLAIN ENGLISH</b>", S['small'])],
                [Paragraph(meaning, S['body'])]]
    bt = Table(box_rows, colWidths=[CW])
    bt.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), GREEN_LT),
        ('LINEBEFORE', (0, 0), (0, -1), 3, GREEN),
        ('LEFTPADDING', (0, 0), (-1, -1), 9), ('RIGHTPADDING', (0, 0), (-1, -1), 9),
        ('TOPPADDING', (0, 0), (0, 0), 6), ('BOTTOMPADDING', (0, 0), (0, 0), 0),
        ('TOPPADDING', (0, 1), (0, 1), 0), ('BOTTOMPADDING', (0, 1), (0, 1), 6),
    ]))
    flow.append(bt)
    flow.append(Spacer(1, 3))
    if picture:
        ct = Table([[Paragraph("<b>CLASSROOM PICTURE</b>", S['small'])],
                    [Paragraph(picture, S['body'])]], colWidths=[CW])
        ct.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), GOLD_LT),
            ('LINEBEFORE', (0, 0), (0, -1), 3, GOLD),
            ('LEFTPADDING', (0, 0), (-1, -1), 9),
            ('RIGHTPADDING', (0, 0), (-1, -1), 9),
            ('TOPPADDING', (0, 0), (0, 0), 6), ('BOTTOMPADDING', (0, 0), (0, 0), 0),
            ('TOPPADDING', (0, 1), (0, 1), 0), ('BOTTOMPADDING', (0, 1), (0, 1), 6),
        ]))
        flow.append(ct)
    if extra:
        flow.append(Spacer(1, 3))
        flow.append(Paragraph(extra, S['small']))
    flow.append(Spacer(1, 7))
    return KeepTogether(flow)


def img(path, w):
    p = os.path.join(ASSETS, path)
    from PIL import Image as PILImage
    iw, ih = PILImage.open(p).size
    h = w * ih / iw
    return Image(p, width=w, height=h, hAlign='CENTER')


# ------------------------------------------------------------------ doc
class Doc(BaseDocTemplate):
    def __init__(self, fn):
        BaseDocTemplate.__init__(self, fn, pagesize=A4,
                                 leftMargin=MARG, rightMargin=MARG,
                                 topMargin=17 * mm, bottomMargin=15 * mm,
                                 title="Effective Assessment Practices — Workshop Handout",
                                 author="KATON 2026")
        frame = Frame(MARG, self.bottomMargin, CW,
                      PH - self.topMargin - self.bottomMargin, id='n',
                      leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
        cover = Frame(MARG, self.bottomMargin, CW,
                      PH - self.topMargin - self.bottomMargin, id='c',
                      leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
        self.addPageTemplates([
            PageTemplate(id='cover', frames=[cover], onPage=self.cover_bg),
            PageTemplate(id='main', frames=[frame], onPage=self.deco),
        ])
        self.page_no = 0

    def cover_bg(self, c, d):
        c.saveState()
        c.setFillColor(GREEN_DK)
        c.rect(0, 0, PW, PH, stroke=0, fill=1)
        c.setFillColor(GOLD)
        c.rect(0, PH - 9, PW, 9, stroke=0, fill=1)
        c.restoreState()

    def deco(self, c, d):
        c.saveState()
        c.setFillColor(GREEN)
        c.rect(0, PH - 7, PW, 7, stroke=0, fill=1)
        c.setFillColor(GREY)
        c.setFont("Body", 7.4)
        c.drawString(MARG, 9.5 * mm,
                     "Effective Assessment Practices  \u00b7  KATON 2026  \u00b7  "
                     "after the NaCCA Teacher Assessment Manual and Toolkit")
        c.drawRightString(PW - MARG, 9.5 * mm, str(c.getPageNumber()))
        c.setStrokeColor(colors.HexColor("#DCE5E1"))
        c.setLineWidth(0.5)
        c.line(MARG, 12.5 * mm, PW - MARG, 12.5 * mm)
        c.restoreState()


story = []

# ================================================================ COVER
story.append(Spacer(1, 34 * mm))
story.append(Paragraph("Effective<br/>Assessment<br/>Practices", S['cover_t']))
story.append(Spacer(1, 7 * mm))
story.append(Paragraph("Plain words. Everyday classrooms. Real Ghanaian schools.",
                       S['cover_s']))
story.append(Spacer(1, 6 * mm))
story.append(Paragraph(
    "A workshop handbook for Senior High School teachers. Every technical term from the "
    "<b>Teacher Assessment Manual and Toolkit</b> explained twice: once in plain English, "
    "and once as a picture from a classroom like yours.", S['cover_b']))
story.append(Spacer(1, 12 * mm))
story.append(Paragraph("KATON 2026", ParagraphStyle(
    'k', fontName="Body-B", fontSize=14, textColor=colors.white)))
story.append(Paragraph(
    "Built on the Teacher Assessment Manual and Toolkit \u2014 National Council for "
    "Curriculum and Assessment (NaCCA), Ministry of Education, Republic of Ghana.",
    ParagraphStyle('k2', fontName="Body", fontSize=8.8, leading=12.4,
                   textColor=colors.HexColor("#B9CFC7"))))
story.append(NextPageTemplate('main'))
story.append(PageBreak())

# ================================================================ CONTENTS
story += band("What is in this handbook")
story.append(Paragraph(
    "This handbook follows the order of the workshop. It is yours to keep, write on and "
    "share. When a term confuses you later in the term, find it here.", S['body']))
story.append(Spacer(1, 4))

toc = [
    ("1", "The Soup and the Meal", "The one picture that explains formative and summative assessment"),
    ("2", "AfL, AaL and AoL", "The three acronyms, untangled"),
    ("3", "The Three Main Purposes", "Before, during and after teaching"),
    ("4", "The Seven Pillars", "What holds a good assessment up"),
    ("5", "Depth of Knowledge", "The staircase of thinking"),
    ("6", "The Jargon Buster", "Every technical term, in plain words"),
    ("7", "Strategies in Action", "Portfolio, debate and practical assessment, brought to earth"),
    ("8", "Phrases We Get Wrong", "Common misreadings, corrected"),
    ("9", "The One-Page Summary", "Pin this above your desk"),
]
rows = [[Paragraph(f"<b>{n}</b>", S['toc']),
         Paragraph(f"<b>{t}</b>", S['toc']),
         Paragraph(s, S['small'])] for n, t, s in toc]
t = Table(rows, colWidths=[10 * mm, 58 * mm, CW - 68 * mm])
t.setStyle(TableStyle([
    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ('LEFTPADDING', (0, 0), (-1, -1), 0), ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ('TOPPADDING', (0, 0), (-1, -1), 4), ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ('LINEBELOW', (0, 0), (-1, -2), 0.4, colors.HexColor("#E2EAE7")),
]))
story.append(t)
story.append(Spacer(1, 10))

# ================================================================ 1. SOUP
story += band("1 \u00b7 The Soup and the Meal",
              "If you remember one picture from this workshop, remember this one")
story.append(Paragraph(
    "The whole idea of assessment lives in one everyday Ghanaian scene. Think of a woman "
    "cooking soup for her family.", S['body']))
story.append(Spacer(1, 4))

story.append(KeepTogether([
    img("02_soup_tasting.jpg", 66 * mm), Spacer(1, 6),
    Paragraph("<b>Tasting the soup \u2014 formative assessment</b>",
              ParagraphStyle('x', fontName="Body-B", fontSize=10.6,
                             textColor=GREEN_DK, spaceAfter=3)),
    Paragraph(
        "She tastes while the soup is still on the fire. No salt? She adds some. Too much "
        "pepper? She adds water. She has not failed the soup \u2014 she has improved it "
        "<i>before</i> anyone sits down to eat. That is assessment <b>for</b> learning, and it "
        "saves the meal.", S['body']),
]))
story.append(Spacer(1, 10))
story.append(KeepTogether([
    img("03_meal_served.jpg", 66 * mm), Spacer(1, 6),
    Paragraph("<b>Serving the meal \u2014 summative assessment</b>",
              ParagraphStyle('x2', fontName="Body-B", fontSize=10.6,
                             textColor=TERRA, spaceAfter=3)),
    Paragraph(
        "The cooking is finished. This is the meal as it is. No amount of tasting will help "
        "now. You serve it, the family eats, and the meal is judged as served. That is "
        "assessment <b>of</b> learning \u2014 end of term, WASSCE, certification.", S['body']),
]))
story.append(Spacer(1, 6))
story.append(Paragraph(
    "<b>The question to ask yourself:</b> does anybody here wait until the food is on the "
    "table before they first taste it? Of course not. So why do we wait until the end-of-term "
    "exam to find out what our learners misunderstood \u2014 when it is too late to fix it?",
    S['quote']))
story.append(Spacer(1, 4))
story.append(Paragraph(
    "One more thing worth noticing: the cook who tastes is not being nosy or wasting time. "
    "Tasting is not an interruption to cooking \u2014 tasting <b>is</b> cooking.",
    S['body']))

story.append(PageBreak())

# ================================================================ 2. THREE SISTERS
story += band("2 \u00b7 AfL, AaL and AoL \u2014 The Three Acronyms, Untangled",
              "These three cause the most confusion. Here they are side by side.")
story.append(datatable(
    ["Acronym", "Full name", "Who is doing the work", "In one plain sentence"],
    [["AfL", "Assessment <b>for</b> Learning", "The TEACHER",
      "I check as we go, so I can teach better."],
     ["AaL", "Assessment <b>as</b> Learning", "The LEARNER",
      "The student checks themselves, so they learn better."],
     ["AoL", "Assessment <b>of</b> Learning", "The SYSTEM and teacher",
      "A final summary of what has been achieved."]],
    widths=[20 * mm, 42 * mm, 40 * mm, CW - 102 * mm]))
story.append(Spacer(1, 6))
story.append(Paragraph(
    "<b>AfL and AaL together make up FORMATIVE assessment. AoL is SUMMATIVE assessment.</b>",
    S['body']))
story.append(Paragraph(
    "<b>A simple test to use in your own practice:</b> ask, \u201cwho is going to change "
    "what they do as a result of this information?\u201d If the answer is the <b>teacher</b>, "
    "it is AfL. If it is the <b>learner</b>, it is AaL. If nobody changes anything and it "
    "simply gets recorded, it is AoL.", S['body']))
story.append(Paragraph(
    "<b>AaL is the one we forget.</b> We tell learners their marks, but we rarely teach them "
    "to check themselves. That skill is what carries them into university and into work \u2014 "
    "and it is the heart of Assessment as Learning.", S['body']))

story.append(Spacer(1, 6))
story += band("Formative vs Summative \u2014 side by side", color=TERRA)
story.append(datatable(
    ["", "FORMATIVE (the tasting)", "SUMMATIVE (the serving)"],
    [["When", "During teaching \u2014 before, during and after the lesson",
      "After teaching \u2014 end of unit, term, year or programme"],
     ["Purpose", "To IMPROVE learning while there is still time",
      "To PROVE and record overall achievement"],
     ["Who uses it", "Teacher and learner", "School, WAEC and other external bodies"],
     ["What it sounds like", "\u201cI see where you got stuck \u2014 let us try it this way\u201d",
      "\u201cThis is your final grade for the term\u201d"],
     ["Examples", "Questions, exit cards, class exercises, observation, peer assessment, drafts",
      "End-of-term exam, WASSCE, projects, portfolios, practicals"]],
    widths=[28 * mm, (CW - 28 * mm) / 2, (CW - 28 * mm) / 2]))
story.append(Spacer(1, 8))

story.append(Paragraph("The golden rule most teachers miss", S['h2']))
story.append(Paragraph(
    "Summative results can be used <b>formatively</b>, and formative work can count "
    "<b>summatively</b>. The label describes the <i>purpose</i> the information is put to \u2014 "
    "not the paper itself.", S['body']))
story += bullets([
    "<b>Using summative for formative:</b> mark the end-of-term paper, but instead of only "
    "recording scores, hand the scripts back, put learners in groups to find the answers they "
    "missed, and re-teach the two topics the whole class failed. That exam just became a "
    "teaching tool.",
    "<b>Using formative for summative:</b> the class exercises and the group project you grade "
    "against a clear rubric become part of the 30% school-based assessment. Your everyday work "
    "is already counting.",
])
story.append(Paragraph(
    "Why this matters to you: the same piece of work can do two jobs. You are not being asked "
    "to double your marking \u2014 you are being asked to use the information twice.", S['small']))

story.append(PageBreak())

# ================================================================ 3. PURPOSES
story += band("3 \u00b7 The Three Main Purposes",
              "Before, during and after teaching \u2014 and what each one is for")

story.append(Paragraph("BEFORE you teach \u2014 Diagnostic assessment", S['h2']))
story.append(Paragraph(
    "Checking what learners already know and can do <b>before</b> you start teaching, so you "
    "know where to begin. It is checking your pantry before you cook: have you got the rice? "
    "The oil? Is the yam already boiled? Only then do you decide what to cook.", S['body']))
story.append(Paragraph(
    "In practice: a five-minute pre-test or entry ticket at the start of a new topic. Ask one "
    "question on Monday about the topic you teach on Wednesday \u2014 and discover that half "
    "the class never understood last term's foundation topic. Tools: pre-tests, entry tickets, "
    "K-W-L charts, concept maps, short interviews, a simple show of hands.", S['small']))

story.append(Paragraph("WHILE you teach \u2014 Formative assessment", S['h2']))
story.append(Paragraph(
    "Checking learning while teaching is still going on, so you can fix things before it is "
    "too late. Used formally and informally to gather information about learners and their "
    "learning needs, so that teachers can modify or improve teaching and learning. "
    "The key word is <b>improve</b> \u2014 not <b>prove</b>.", S['body']))
story.append(Paragraph(
    "In practice: questions, observation, class exercises, draft work, peer and self "
    "assessment, exit cards, Think-Pair-Share, concept maps, projects, reflection journals, "
    "impromptu quizzes, checklists and rubrics.", S['small']))

story.append(Paragraph("AT THE END \u2014 Summative assessment", S['h2']))
story.append(Paragraph(
    "A formal, planned check at the end to sum up overall achievement \u2014 for selection, "
    "certification and placement. The cooking is finished; this is the meal as it is. Summative "
    "assessment is not the enemy. It has a real and necessary job. The point is only that "
    "summative <i>alone</i> is a poor diet: if you only ever serve meals and never taste, you "
    "will keep serving the same mistakes.", S['body']))
story.append(Paragraph(
    "Examples: class tests, end-of-term and end-of-year examinations, Test of Practical, term "
    "papers, projects, portfolios, performance and practical assessment, research reports.",
    S['small']))

story.append(Spacer(1, 4))
story.append(Paragraph("Two more forms worth knowing", S['h2']))
story.append(concept(
    "Differentiated assessment",
    "Using different ways of assessing so that every learner, whatever their level or need, "
    "has a fair chance to show what they can do. <b>You adapt the route, never the "
    "destination.</b>",
    "You have a very strong learner and a learner who is struggling. Both must show they "
    "understand photosynthesis \u2014 that is the learning outcome and it does not change. One "
    "writes the essay, the other draws and labels the diagram and explains it orally. Both have "
    "proven the same outcome. Not easier work for the weaker child, but a different door into "
    "the same room.",
    "<b>The six levers:</b> varied formats \u00b7 flexible deadlines \u00b7 varying task "
    "difficulty \u00b7 accommodations \u00b7 individualised feedback \u00b7 learner involvement."))
story.append(concept(
    "Performance-based, practical and authentic assessment",
    "Learners <b>actually do</b> something real \u2014 not just write about it. "
    "\u201cAuthentic\u201d simply means: it looks like real life.",
    "You do not ask them to describe a titration on paper \u2014 they stand at the bench and "
    "titrate. You do not ask them to list the parts of a machine \u2014 they dismantle and "
    "reassemble it. Examples: experiments, exhibitions and STEM fairs, fieldwork reports, "
    "projects, drama performances, debate, simulation, oral and aural assessment."))
story.append(concept(
    "e-Assessment",
    "Using information technology and digital tools to assess learning. It can be used for both "
    "formative and summative purposes.",
    "<b>It does not need a computer lab.</b> A class WhatsApp group where learners submit a "
    "60-second voice note explaining a concept; a Google Form quiz that marks itself; a Kahoot "
    "round at the end of a lesson; photographs of practical work sent for feedback. Where "
    "internet is a challenge, WhatsApp, SMS and offline forms still count."))

story.append(PageBreak())

# ================================================================ 4. SEVEN PILLARS
story += band("4 \u00b7 The Seven Pillars of Effective Assessment",
              "Section 2 of the manual. Think of them as seven pillars holding up a roof \u2014 "
              "remove one and the building leans.")
story.append(Paragraph(
    "You will not get all seven perfect every week. But when an assessment disappoints you, "
    "come back to this list \u2014 the answer is almost always hiding in one of the seven.",
    S['body']))
story.append(Spacer(1, 4))

story.append(concept(
    "1 \u00b7 Validity \u2014 measuring the right thing",
    "Are you actually measuring the thing you say you are measuring? A valid assessment "
    "measures the important learning outcomes of the curriculum \u2014 not just whatever was "
    "easiest to set.",
    "A trader uses a weighing scale to sell yam. If she used that scale to measure the "
    "<b>length</b> of the yam, the reading would be worthless \u2014 not because the scale is "
    "broken, but because it is measuring the wrong thing. Now think of your last class test. "
    "Your learning outcome said learners should <b>analyse</b> data. But every question asked "
    "them to <b>recall</b> a definition. Your test was neat. It was well marked. It was simply "
    "measuring the wrong thing.",
    "<b>The quick check:</b> \u201cdoes this question actually test the skill I said I was "
    "teaching?\u201d <b>The fix:</b> build a table of specification \u2014 a grid mapping each "
    "question to a topic and a thinking level."))

story.append(concept(
    "2 \u00b7 Reliability \u2014 the same result every time",
    "Do you get the same result every time, whoever is doing the marking? Reliable results are "
    "dependable enough to make decisions with.",
    "A market scale that gives the same weight for the same pile of tomatoes \u2014 on Monday "
    "or Saturday, whoever is standing at the stall. Now imagine the scale was generous on "
    "Monday and harsh on Saturday. Nobody would trust that market again. If this term's 30 "
    "marks in your subject are generously given and next term's are harsh, learners are being "
    "ranked against a ruler that keeps changing length.",
    "<b>The manual is blunt about this:</b> the \u201cconnoisseur\u201d approach \u2014 "
    "\u201cI know it when I see it but I cannot put it into words\u201d \u2014 is not "
    "acceptable. <b>The fix:</b> clear outcomes, a colleague reviewing your marking, more than "
    "one method for the same outcome, marking schemes with weightings, and rubrics given out "
    "<i>before</i> the task."))

story.append(concept(
    "3 \u00b7 Fairness and ethics \u2014 an equitable chance",
    "Giving learners an equitable opportunity to demonstrate what they know and can do, taking "
    "into account their ability, learning styles, gender and Special Educational Needs. It also "
    "means never assessing what you have not taught.",
    "A test question asks learners to explain the rules of cricket. A learner from a village "
    "that has never played \u2014 or seen \u2014 cricket cannot answer, and fails. That question "
    "did not measure physics. It measured cricket. Now swap it for the physics of a loaded "
    "trotro braking at a junction, and every learner in the room has something to think with. "
    "Same physics. Same difficulty. Fair.",
    "<b>Red flags:</b> culturally biased content \u00b7 unfamiliar words and examples \u00b7 "
    "assessing content you never taught \u00b7 failing to give SEN learners extra time or "
    "alternative formats."))

story.append(concept(
    "4 \u00b7 Transparency \u2014 no secret exams",
    "Making the assessment process and criteria clear and understandable to learners. They "
    "should know what is being judged, how it is judged, and what counts as a pass.",
    "Two teachers set the same project. The first says \u201cdo the project, I will mark "
    "it\u201d \u2014 and the learners spend the week guessing. The second hands out the rubric "
    "on day one: what is expected, the marks for each part, what a top answer looks like. Both "
    "teachers mark with the same rubric. The second class almost always scores higher \u2014 "
    "not because they worked harder, but because they knew where to aim.",
    "<b>Transparency also requires:</b> sharing the learning outcomes being assessed \u00b7 "
    "telling learners the pass mark \u00b7 sharing results openly with learners and parents or "
    "guardians \u00b7 giving learners a route to seek review and redress."))

story.append(concept(
    "5 \u00b7 Inclusivity \u2014 fair and accessible to ALL",
    "Creating assessment practices that are fair and accessible to every learner, regardless of "
    "gender, disability, poverty, background or learning difference. The framework names three "
    "lenses you must apply: <b>GESI</b> (Gender Equality and Social Inclusion), <b>SEL</b> "
    "(Socio-Emotional Learning) and <b>SEN</b> (Special Educational Needs).",
    "A learner with low vision cannot read an ordinary small-print paper \u2014 but they "
    "understand the work perfectly. A learner who is hard of hearing misses the spoken "
    "instructions you gave once, quickly, at the front. A learner who is dyslexic knows the "
    "science but loses marks to their handwriting. In all three cases we did not measure their "
    "learning. We measured their eyes, their ears and their handwriting.",
    "<b>Practical tools:</b> braille \u00b7 oral translation \u00b7 text-to-speech and other "
    "assistive technology \u00b7 sign language interpretation \u00b7 extra time \u00b7 "
    "alternative formats. <b>A warning worth giving:</b> if you are assessing science "
    "understanding, decide deliberately whether handwriting and grammar should cost marks."))

story.append(concept(
    "6 \u00b7 Practicability \u2014 can you actually do it?",
    "The assessment must be feasible, convenient, efficient and successful with the real "
    "resources you have \u2014 your actual class size, your actual timetable, your actual "
    "materials.",
    "You design a wonderful one-hour one-on-one oral assessment for your Government class. It is "
    "rigorous. It is valid. It is inclusive. It is also completely impossible with 60 learners "
    "\u2014 and by Thursday it will have collapsed into chaos, you will be three weeks behind, "
    "and you will abandon it. The plan failed not because it was bad, but because it was not "
    "practicable. A good assessment is one you can still run in week eight.",
    "<b>The three questions:</b> do I have the materials and the security? \u00b7 does this "
    "format suit my class size and their level? \u00b7 do I have the time to set it, mark it "
    "<i>and</i> give constructive feedback?"))

story.append(concept(
    "7 \u00b7 Assessment utility \u2014 is it actually useful?",
    "The assessment must give you and the learner information worth having \u2014 and that "
    "information must lead to action.",
    "You set a class test. You mark it. You enter the scores in the mark book. You file the "
    "papers. Term ends. Now ask honestly: did anybody learn anything from that test? If it did "
    "not change what you taught next, and did not change what the learner did next, then all "
    "that happened was data collection. Utility is the difference between a test that teaches "
    "and a test that merely records.",
    "<b>A test to offer yourself:</b> name one thing you changed last term because of an "
    "assessment result. A cook who tastes the soup and then does nothing about it has wasted the "
    "taste. Assessing without acting is not assessment."))

story.append(PageBreak())

# ================================================================ 5. DoK
story += band("5 \u00b7 Depth of Knowledge \u2014 The Staircase of Thinking",
              "How DEEP is the thinking? Not how difficult the words look.")
story.append(Paragraph(
    "<b>The biggest misconception:</b> teachers think a question is DoK 4 because it uses big "
    "words or long sentences. It does not. A long question with big vocabulary that asks you to "
    "remember something is still DoK 1. A short, simple question that asks you to judge a real "
    "situation and defend your judgement can be DoK 3. <b>DoK is about the thinking, not the "
    "vocabulary.</b>", S['body']))
story.append(Spacer(1, 3))
story.append(datatable(
    ["Level", "Name", "Plain English", "SHS example \u2014 Social Studies"],
    [["DoK 1", "Recall / Reproduction", "Give back a fact you were taught.",
      "\u201cState three causes of climate change.\u201d"],
     ["DoK 2", "Skills and Concepts", "Use your knowledge in a familiar routine.",
      "\u201cExplain how the greenhouse effect works, using a diagram.\u201d"],
     ["DoK 3", "Strategic Thinking",
      "Reason, justify, plan. There is more than one route to the answer.",
      "\u201cA cocoa farmer's yield is falling. Using the data given, diagnose the likely cause "
      "and justify your answer.\u201d"],
     ["DoK 4", "Extended Thinking",
      "Investigate over time, combine ideas, produce something original.",
      "\u201cConduct a week-long survey in your community on waste disposal and present "
      "recommendations to the school assembly.\u201d"]],
    widths=[15 * mm, 32 * mm, 52 * mm, CW - 99 * mm]))
story.append(Spacer(1, 6))
story.append(Paragraph(
    "<b>Why it matters for validity:</b> if your learning outcome is at DoK 3 "
    "(\u201canalyse\u201d, \u201cevaluate\u201d) and your test questions are all at DoK 1 "
    "(\u201cstate\u201d, \u201clist\u201d, \u201cdefine\u201d), your assessment is invalid "
    "\u2014 no matter how well it is marked.", S['body']))
story.append(Paragraph(
    "<b>A five-minute exercise:</b> take one question from your last test and label it. Most "
    "teachers find their whole paper sits at DoK 1 and 2. That is a diagnosis, not a criticism "
    "\u2014 and it is easy to fix with a table of specification.", S['body']))

story.append(Spacer(1, 6))
story += band("The Six-Phase Assessment Plan",
              "Turning all of this into a plan you can actually follow", color=TERRA)
story.append(datatable(
    ["Phase", "What it is", "Why it matters"],
    [["1. Learning outcome", "What should the learner be able to DO?",
      "Everything else flows from this one sentence. If you cannot state it, you cannot assess it."],
     ["2. Assessment strategies", "HOW will I find out?",
      "Choose the method that matches the outcome \u2014 not always a test."],
     ["3. Tasks or questions", "The actual questions learners will meet.",
      "This is where validity is won or lost."],
     ["4. Grading criteria or rubrics", "What \u201cgood\u201d looks like, at each level.",
      "Decide this BEFORE learners start \u2014 not while marking at midnight. This is the "
      "biggest cause of unreliable marking."],
     ["5. Timeline and sequencing", "When each assessment happens.",
      "Spread the load so week eight is not a disaster."],
     ["6. Feedback and reporting", "How learners act on the feedback, and how parents hear.",
      "If there is no planned time for feedback, there will be no feedback \u2014 only marks."]],
    widths=[38 * mm, 52 * mm, CW - 90 * mm]))

story.append(PageBreak())

# ================================================================ 6. JARGON BUSTER
story += band("6 \u00b7 The Jargon Buster",
              "Every technical term, in plain words, with a picture")

JB = [
    ("Assessment",
     "Finding out what a learner knows and can do, and how far they have come towards the goal "
     "you set.",
     "It is not only the end-of-term paper. It is the question you asked mid-lesson and the "
     "blank faces that answered it. It is the exercise book you marked at 9pm. Every one of "
     "those is assessment \u2014 and every one changed what you did next."),
    ("Learning outcome",
     "A clear sentence saying what a learner should be able to <b>do</b> by the end of a lesson "
     "or unit.",
     "\u201cBy the end of this lesson, the learner should be able to balance a simple chemical "
     "equation.\u201d Not \u201cteach chapter 4.\u201d"),
    ("Diagnostic assessment",
     "Checking what learners already know and can do <b>before</b> you start teaching.",
     "Checking the pantry before you cook. A five-minute pre-test or entry ticket at the start "
     "of a new topic."),
    ("Formative assessment",
     "Checking learning <b>while</b> teaching is still going on, so you can fix things before "
     "it is too late.",
     "Tasting the soup while it is still on the fire. Halfway through the mole concept you see "
     "blank faces and stop to re-teach with a different example."),
    ("Summative assessment",
     "A formal, planned check at the end, to sum up overall achievement for certification.",
     "Serving the meal at the table. The cooking is finished; this is the meal as it is."),
    ("Differentiated assessment",
     "Different ways of assessing so every learner has a fair chance to prove the same outcome. "
     "<b>Change the route, not the destination.</b>",
     "Your top learner writes the essay; your struggling learner draws, labels and explains "
     "orally. Same outcome proven twice."),
    ("Performance / authentic assessment",
     "Learners actually <b>do</b> it, in a way that looks like real life.",
     "Not describing a titration on paper \u2014 standing at the bench doing it."),
    ("Portfolio",
     "A <b>curated</b> collection of a learner's work across time, showing progress.",
     "Not a folder of everything. Chosen pieces, with a reflection where the learner explains "
     "why each was chosen and what they improved."),
    ("Rubric",
     "A table showing exactly what counts and how many marks each part is worth, at each level.",
     "A recipe card with measurements. Not \u201cadd enough salt\u201d but \u201cone "
     "teaspoon.\u201d Give it out <b>before</b> the task, never after."),
    ("Marking scheme",
     "The model answer sheet showing where each mark falls.",
     "Your colleague should be able to mark your paper and produce almost the same scores you did."),
    ("Table of specification",
     "A planning grid making sure your test covers the content and the thinking levels it is "
     "supposed to.",
     "Instead of 20 recall questions because recall is easy to set, the grid forces balance "
     "across topics and DoK levels. Your best defence of validity."),
    ("Depth of Knowledge (DoK)",
     "How <b>deep</b> the thinking is \u2014 not how hard the question looks. Four levels.",
     "Four steps up a staircase: recall, skills, strategic thinking, extended thinking."),
    ("Project-based assessment",
     "Learners work on a meaningful, extended task and produce something real.",
     "Not a weekend homework. A term-long task \u2014 research, plan, produce, present, and be "
     "judged on all of it."),
    ("Case study",
     "A real or realistic scenario learners must analyse and decide on.",
     "\u201cHere are the accounts of a small market business that is failing. As a consultant, "
     "what would you advise, and why?\u201d"),
    ("Test of Practical Knowledge (TPK)",
     "A written paper testing your knowledge of <b>how to do</b> a practical \u2014 the theory "
     "behind it.",
     "Why you add acid to the burette the way you do; what would happen if you did not; how you "
     "would correct a reading."),
    ("Questioning",
     "Using questions deliberately to find out what learners understand and to push thinking "
     "deeper.",
     "Not only \u201cwho can tell me\u2026?\u201d to the eager boy in the front row. Deliberate "
     "wait time, questions directed by name, and follow-ups like \u201cwhy do you say so?\u201d"),
    ("Think-Pair-Share",
     "Think alone, then discuss with a partner, then share with the whole class.",
     "Three minutes of genuine thinking time, so the quiet learners at the back have something "
     "to say by the time you ask."),
    ("Think-aloud",
     "Learners say out loud what is going on in their head as they solve a problem.",
     "\u201cI am reading the question\u2026 first I will find the mass\u2026 now I am not sure "
     "which formula to use.\u201d It exposes the reasoning a written answer hides."),
    ("Exit card / exit ticket",
     "A short question at the very end of the lesson, answered on a slip before learners leave.",
     "\u201cWrite one thing you learned today and one thing you are still confused about.\u201d "
     "Three minutes, scrap paper, works with 60 learners."),
    ("Entry ticket",
     "A short question at the start of a lesson showing what learners already bring.",
     "Free diagnostic assessment, every single day, at no cost."),
    ("K-W-L chart",
     "Three columns: what I <b>K</b>now, what I <b>W</b>ant to know, what I <b>L</b>earned.",
     "Before the topic they fill K and W; after the topic they fill L. It makes learning visible "
     "and the learner reflective."),
    ("Concept map",
     "A diagram showing how ideas are linked together.",
     "Not a mind-map of neat bubbles only \u2014 the <b>arrows must be labelled</b> "
     "(\u201ccauses\u201d, \u201cis part of\u201d, \u201cdepends on\u201d). The labels are where "
     "the understanding is."),
    ("Graphic organiser",
     "Any visual frame that helps learners arrange their thoughts \u2014 Venn diagram, "
     "flowchart, table.",
     "Comparing two economic systems in a Venn diagram, where the overlap is the interesting "
     "part."),
    ("Reflection journal",
     "A learner's own written record of their thinking and growth, reviewed over time.",
     "Five minutes at the end of a double period: \u201cwhat did I find hard today? What will I "
     "do differently next time?\u201d"),
    ("Peer assessment",
     "Learners assess each other's work against given criteria.",
     "Only works if they have the <b>rubric in hand</b> and have been trained to be kind and "
     "specific. Otherwise it becomes a popularity contest."),
    ("Self-assessment",
     "Learners judge their own work against the criteria.",
     "Learner marks their own draft against the rubric, then you mark it. Where the two differ "
     "is a goldmine of teaching information."),
    ("Metacognition",
     "Thinking about your own thinking. Knowing <b>how</b> you learn and <b>where</b> you are "
     "stuck.",
     "\u201cI understand the theory but I keep making arithmetic slips\u201d shows "
     "metacognition. \u201cI just do not get maths\u201d does not \u2014 yet."),
    ("Feedback",
     "Information that helps a learner close the gap between where they are and where they need "
     "to be.",
     "<b>The mechanic.</b> A useless mechanic says \u201cthe car is bad.\u201d A good mechanic "
     "points at one exact part: \u201cthis hose is cracked; replace it and the overheating "
     "stops.\u201d Specific, targeted, actionable. A score is a verdict; feedback is a map."),
    ("Constructive feedback",
     "Feedback that points to the next step, not just the fault.",
     "\u201cYour introduction is clear, but your conclusion does not answer the question. "
     "Rewrite just the last paragraph, and here is how.\u201d"),
    ("Gamification",
     "Using game elements \u2014 points, levels, badges, competition \u2014 to drive learning "
     "and assessment.",
     "A house-based quiz with a points board; a \u201clevel up\u201d chart on the wall where "
     "learners move up as they master each outcome."),
    ("Simulation",
     "Creating a safe model of a real situation for learners to work through.",
     "A Model United Nations debate; a courtroom role-play in a Government class; a mock "
     "business negotiation."),
    ("Dramatisation / dramatic monologue",
     "Learning and being assessed through acting it out.",
     "A Literature learner performs a monologue as a character from <i>Things Fall Apart</i>; "
     "a History class dramatises a historical negotiation."),
    ("Displays and exhibitions",
     "Learners present their work publicly, for an audience.",
     "The STEM fair where Robotics learners exhibit robots they dismantled and reassembled, and "
     "the school community views them."),
    ("Checklist",
     "A simple yes/no list of what should be present.",
     "\u201cDoes the practical report have: aim, apparatus, method, results, conclusion?\u201d "
     "Tick, tick, tick. Fast and reliable for skill sequences."),
    ("Critiquing",
     "Learners give reasoned judgement on a piece of work, a performance or an argument, with "
     "justification.",
     "\u201cThe layout works because the eye moves top-to-bottom, but the key is missing.\u201d"),
    ("Fieldwork / field report",
     "Learning out in the real environment, then writing up what was observed and analysed.",
     "A Biology class visiting a local river or farm, recording species and observations, then "
     "writing a structured report."),
    ("Learning analytics",
     "Looking at patterns in assessment data to spot trends and act early.",
     "Your mark book shows 18 out of 45 learners failed the same mole calculation question "
     "\u2014 one whole-class gap you can now re-teach, instead of 18 private problems."),
]

for term, meaning, picture in JB:
    story.append(concept(term, meaning, picture))

story.append(PageBreak())

# ================================================================ 7. SYSTEM TERMS
story += band("6b \u00b7 The System: Acronyms and Structures",
              "The letters you see on circulars and in the curriculum")
story.append(datatable(
    ["Acronym", "Full name", "Plain English", "Why teachers should care"],
    [["<b>NaCCA</b>", "National Council for Curriculum and Assessment",
      "Sets the curriculum and assessment guidance", "This manual comes from them"],
     ["<b>MoE</b>", "Ministry of Education", "The ministry running the system",
      "Funds and directs policy"],
     ["<b>GES</b>", "Ghana Education Service", "Runs public pre-tertiary schools",
      "Your employer"],
     ["<b>WAEC</b>", "West African Examinations Council", "Runs WASSCE, the external exam",
      "Their 70% sits on top of your 30%"],
     ["<b>NaSIA</b>", "National School Inspectorate Authority", "Inspects schools",
      "They will look at your assessment records"],
     ["<b>NTC</b>", "National Teaching Council", "Regulates the teaching profession",
      "Licensure and continuous professional development"],
     ["<b>NPLAF</b>", "National Pre-tertiary Learning and Assessment Framework",
      "The national master document on assessment",
      "Read its inclusivity section (page 32)"],
     ["<b>SEAF</b>", "Secondary Education Assessment Framework",
      "How SHS learners are assessed across grade levels", "The direct rulebook for SHS"],
     ["<b>SBA</b>", "School-Based Assessment", "Assessment done in school, by you, and counted",
      "It is the <b>30%</b>"],
     ["<b>PLC</b>", "Professional Learning Community",
      "Teachers who meet regularly to learn together", "Like today, but ongoing"],
     ["<b>CPD</b>", "Continuous Professional Development", "Ongoing learning for teachers",
      "Your licence renewal depends on it"],
     ["<b>SEN</b>", "Special Educational Needs", "Learners who need specific support",
      "Inclusivity is not optional"],
     ["<b>GESI</b>", "Gender Equality and Social Inclusion",
      "No learner disadvantaged by gender, disability or background",
      "A cross-cutting requirement"],
     ["<b>SEL</b>", "Socio-Emotional Learning", "The learner's feelings, attitudes and social skills",
      "It is assessed, not just academic work"],
     ["<b>DoK</b>", "Depth of Knowledge", "The four levels of thinking depth",
      "Used in the curriculum's assessment tasks"],
     ["<b>21st-century skills</b>", "\u2014",
      "Critical thinking, creativity, collaboration, communication",
      "WAEC now tests these"]],
    widths=[24 * mm, 44 * mm, 52 * mm, CW - 120 * mm]))

story.append(Spacer(1, 8))
story.append(Paragraph("The 70/30 split \u2014 your marks now have weight", S['h2']))
story.append(Paragraph(
    "<b>Thirty per cent</b> of a learner's final grade comes from work done in your school and "
    "marked by you \u2014 portfolios, performance and project work, plus end-of-term "
    "examinations. <b>Seventy per cent</b> comes from WAEC's final examination.", S['body']))
story.append(Paragraph(
    "For years, internal marks were treated as a formality at the back of the mark book. That "
    "has changed. The internal 30% now sits on a <b>school-based transcript</b>, with far "
    "greater transparency and quality assurance. Practically: if your school's 30% is not "
    "credible, transparent and defensible, the learner is the one who suffers when their "
    "transcript is scrutinised \u2014 and the school is the one that is exposed. Reliable "
    "rubrics are no longer good practice. They are protection.", S['body']))
story.append(Paragraph(
    "<b>The school-level question worth asking:</b> do all of us teaching the same subject at "
    "the same level use the same criteria? Could we produce the evidence if NaSIA asked?", S['body']))

story.append(PageBreak())

# ================================================================ 7. STRATEGIES
story += band("7 \u00b7 Strategies in Action",
              "Three of the manual's strategies, brought down to earth")
story.append(Paragraph(
    "The manual lists around thirty assessment strategies in Section 7. Here are three that teachers "
    "most often get wrong, or avoid because they look like too much work. Notice the same pattern in "
    "all three: <b>the strategy is not the hard part. Deciding the criteria before you start is.</b>",
    S['body']))
story.append(Spacer(1, 4))

story.append(concept(
    "Portfolio \u2014 the learner's own story",
    "A carefully chosen collection of a learner's work across time, showing progress towards the "
    "learning outcomes. Usable formatively and, with a clear rubric, summatively.",
    "Two things separate a real portfolio from a pile of papers. First it is <b>curated</b> \u2014 the "
    "learner chooses what goes in, and learns by choosing. Second it carries <b>reflection</b> \u2014 "
    "beside each piece, a short note: \u201cI chose this because\u2026\u201d and \u201cnext time I "
    "would\u2026\u201d. Without reflection you have a folder. With it, you have assessment as learning.",
    "<b>Where the manual helps:</b> Appendix C.1 gives worked portfolio exemplars for Science and "
    "Mathematics, including the rubric. <b>Time-saver:</b> you do not mark every piece \u2014 you mark "
    "the reflections and the final selection, which is where the thinking is visible."))

story.append(concept(
    "Debate \u2014 thinking out loud",
    "Learners take a position on an issue, argue it with evidence, and respond to the opposing case. "
    "You assess the quality of the thinking, not the conclusion.",
    "A Government class debates: \u201cShould Ghana raise the minimum wage?\u201d Half the room argues "
    "for it, half against. What you are assessing: clarity of the argument \u00b7 use of evidence, not "
    "just opinion \u00b7 quality of rebuttal \u2014 do they actually answer the other side? \u00b7 "
    "listening \u00b7 respectful disagreement.",
    "<b>The rule that makes it work:</b> give the rubric <i>before</i> the debate and give learners "
    "research time. A debate sprung on learners with no preparation assesses <b>confidence, not "
    "competence</b> \u2014 and it rewards the loudest voices. Offer a written or small-group "
    "alternative for learners who cannot perform before a large audience."))

story.append(concept(
    "Practical and performance \u2014 doing it for real",
    "Learners demonstrate a skill in a real or simulated setting, and you judge the <b>doing</b> "
    "\u2014 not just the writing about it. The manual calls this performance-based, practical or "
    "authentic assessment.",
    "The Chemistry learner stands at the bench and titrates. The Technical Skills learner dismantles "
    "and reassembles an engine. The Visual Arts learner throws a pot. The Music learner performs. You "
    "are watching the skill, not reading about it. And if you have no laboratory, a simulated or "
    "improvised setting still counts \u2014 the manual explicitly allows \u201creal or simulated\u201d "
    "contexts.",
    "<b>The teacher's five jobs, from the manual:</b> design a task requiring real-life application "
    "\u00b7 provide resources, guidance and support \u00b7 evaluate against <i>predetermined</i> "
    "criteria \u00b7 model the skill \u00b7 serve as coach or mentor throughout. <b>Why the rubric "
    "matters most here:</b> a written script can be re-read; a practical is judged live, once. Without "
    "criteria fixed in advance you will mark the confident performer higher than the careful one "
    "\u2014 and that is a reliability failure, not an opinion."))

story.append(PageBreak())

# ================================================================ 8. MISREADINGS
story += band("8 \u00b7 Phrases We Get Wrong",
              "Common misreadings, corrected")
story.append(datatable(
    ["People say\u2026", "But the manual means\u2026"],
    [["\u201cFormative assessment is just class exercises.\u201d",
      "Formative assessment is a <b>purpose</b> \u2014 any assessment used to <i>improve</i> "
      "learning, including a homework, a question or a conversation. The tool matters less than "
      "the use."],
     ["\u201cContinuous assessment means plenty of tests.\u201d",
      "It means <b>a variety of evidence over time</b> \u2014 projects, portfolios, "
      "observations, performance \u2014 not just a bigger pile of papers."],
     ["\u201cI differentiate, so I give the weak ones easier work.\u201d",
      "Differentiation <b>changes the route, not the destination.</b> The learning outcome "
      "stays; the support and format change."],
     ["\u201cThe exam is over, so I just record the marks.\u201d",
      "The manual explicitly says to <b>use summative results formatively</b> \u2014 hand the "
      "paper back, let them discuss the errors, and re-teach the gaps."],
     ["\u201cI marked it and gave it back.\u201d",
      "Feedback without <b>time to act on it</b> is decoration. Leave space in the next lesson "
      "for learners to use it."],
     ["\u201cFeedback means saying \u2018good work\u2019.\u201d",
      "It means telling the learner <b>exactly where the gap is and what the next step is.</b>"],
     ["\u201cSetting multiple-choice questions is easy.\u201d",
      "It is the most difficult item type to write well. Distractors must reflect <b>real "
      "misconceptions</b>, not random wrong answers."],
     ["\u201cAssessment is extra work on top of teaching.\u201d",
      "Assessment is the thing that tells you whether your teaching landed. Done well, it "
      "<b>reduces</b> your work \u2014 because you stop guessing what to re-teach."]],
    widths=[62 * mm, CW - 62 * mm]))

story.append(Spacer(1, 8))
story.append(Paragraph("Three things to try on Monday", S['h2']))
story += bullets([
    "<b>An exit card.</b> Last three minutes: \u201cwrite one thing you learned today and one "
    "thing you are still confused about.\u201d Collect at the door. You now know exactly where "
    "to start next lesson. Three minutes, scrap paper, works with 60 learners.",
    "<b>An entry ticket.</b> First three minutes: one question that shows what they already "
    "bring. Free diagnostic assessment, every single day.",
    "<b>A Think-Pair-Share.</b> Think alone, discuss with a partner, then share with the class. "
    "Three minutes of genuine thinking time means the quiet learners at the back have something "
    "to say when you ask.",
])
story.append(Paragraph(
    "<b>Pick one. Not all three.</b> Spoken commitment increases follow-through \u2014 tell the "
    "person next to you which one you chose.", S['small']))

story.append(PageBreak())

# ================================================================ 9. SUMMARY
story += band("9 \u00b7 The One-Page Summary", "Pin this above your desk")
story.append(Spacer(1, 6))
CEN = ParagraphStyle('cen', fontName="Body", fontSize=11.5, leading=18,
                      textColor=INK, alignment=TA_CENTER, spaceAfter=9)
story.append(Paragraph(
    '<font size="16" color="#074A35"><b>Assessment is not the exam at the end.</b></font><br/>'
    '<font size="12.5" color="#0B6E4F">It is the nervous system of your teaching.</font>',
    ParagraphStyle('c0', parent=CEN, spaceAfter=14)))
story.append(Paragraph(
    "<b>Assessment FOR learning</b> \u2014 the teacher listens and adjusts.<br/>"
    "<b>Assessment AS learning</b> \u2014 the learner reflects and adjusts.<br/>"
    "<b>Assessment OF learning</b> \u2014 the system records and certifies.", CEN))
story.append(Spacer(1, 4))
story.append(Paragraph(
    "<b>BEFORE</b> you teach, you check what they bring &nbsp;"
    "<font color='#5A6B72'>(diagnostic)</font><br/>"
    "<b>WHILE</b> you teach, you check and adjust &nbsp;"
    "<font color='#5A6B72'>(formative)</font> &nbsp;\u2014 <i>you taste the soup</i><br/>"
    "<b>AT THE END</b>, you summarise &nbsp;"
    "<font color='#5A6B72'>(summative)</font> &nbsp;\u2014 <i>you serve the meal</i><br/>"
    "<b>ALL THE WAY THROUGH</b>, you make it valid, reliable, fair, transparent,<br/>"
    "inclusive, practicable and useful.", CEN))
story.append(Spacer(1, 8))

quote = Table([[Paragraph(
    "<font size=\"13\" color=\"#074A35\"><b>And the two sentences that carry it all:</b></font>"
    "<br/><br/><font size=\"12\">The learner must know what \u201cgood\u201d looks like "
    "\u2014 and must be given time to get there.</font>",
    ParagraphStyle('q', fontName="Body", fontSize=12, leading=19, textColor=INK))]],
    colWidths=[CW])
quote.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, -1), GOLD_LT),
    ('LINEBEFORE', (0, 0), (0, -1), 4, GOLD),
    ('LEFTPADDING', (0, 0), (-1, -1), 12), ('RIGHTPADDING', (0, 0), (-1, -1), 12),
    ('TOPPADDING', (0, 0), (-1, -1), 12), ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
]))
story.append(quote)
story.append(Spacer(1, 12))

story.append(Paragraph("Where to go for more", S['h2']))
story += bullets([
    "<b>The Teacher Assessment Manual and Toolkit \u2014 Handbook for Teachers</b>, NaCCA, "
    "Ministry of Education, Republic of Ghana. Free on the NaCCA curriculum resources portal "
    "(curriculumresources.edu.gh). <b>Appendix C</b> contains ready-made exemplars and rubrics "
    "for almost every strategy named in this handbook \u2014 portfolios, projects, debates, "
    "questioning, practicals, gamification and more.",
    "<b>Appendix A</b> lists the assessment strategies embedded in each subject's curriculum, "
    "including strategies for learners with Special Educational Needs and the Depth of Knowledge "
    "levels.",
    "<b>Appendix B</b> is a toolkit covering learners with SEN, identifying gifted and talented "
    "learners, internal assessment practices and critical thinking skills.",
])
story.append(Spacer(1, 8))
story.append(Paragraph(
    "Compiled for the KATON 2026 workshop on Effective Assessment Practices. Definitions and "
    "frameworks follow the NaCCA Teacher Assessment Manual and Toolkit; the plain-English "
    "explanations, classroom pictures and matter-of-fact advice were added for teachers, by "
    "teachers. Share it freely.", S['small']))

Doc(OUT).build(story)
print("saved:", OUT)
