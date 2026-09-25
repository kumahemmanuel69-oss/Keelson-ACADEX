# Keelson-ACADEX

An intelligent Research agent.

---

## KATON 2026 — Effective Assessment Practices Workshop

Presentation materials built from the **Teacher Assessment Manual and Toolkit**
(NaCCA / Ministry of Education, Republic of Ghana), rewritten for an SHS staff
audience: every technical term gets a plain-English meaning and an everyday
Ghanaian classroom picture.

### Deliverables (`deliverables/`)

| File | What it is |
|---|---|
| `Effective-Assessment-Practices-KATON-2026.pptx` | 51-slide editable deck, 16:9, speaker notes on every slide |
| `Assessment-Handout-KATON-2026.pdf` | 28-page A4 companion handbook for colleagues |
| `Presenters-Run-Sheet.md` | Timings, core-slide markers, facilitation cues |
| `Assessment-Jargon-Buster.md` | The glossary as plain markdown |

### Deck follows the workshop agenda (2 h 30 m)

The deck maps 1:1 onto the KATON running order, with the agenda timing printed as
a badge on every segment.

| Clock | Segment | Slides |
|---|---|---|
| 00:00 | Welcome & Workshop Objectives | 1–3 |
| 00:05 | Icebreaker — One Word for Assessment | 4 |
| 00:10 | Session 1 — Understanding Assessment: Purpose & Types | 5–13 |
| 00:30 | Session 2 — Ghana's Assessment Policy Landscape | 14–19 |
| 00:45 | Break | 20 |
| 01:00 | Session 3 — Principles of Effective Assessment | 21–29 |
| 01:15 | Session 4 — Assessment Strategies Across Levels | 30–38 |
| 01:35 | Practical Workshop — Design a Task & Rubric | 39–43 |
| 02:00 | Gallery Walk & Group Share-Outs | 44 |
| 02:10 | Using Assessment Data: Feedback & Records | 45–48 |
| 02:20 | Common Challenges & Practical Solutions | 49 |
| 02:25 | Action Planning & Closing | 50–51 |

### Workshop objectives (slide 3)

1. **Distinguish & Apply** — distinguish formative and summative assessment, and use each purposefully at the right point in a lesson or term.
2. **Navigate Policy** — explain Ghana's continuous assessment requirements, and locate your grade level within the Standards-Based Curriculum.
3. **Design Fair Tools** — create valid, fair and inclusive assessment tasks and simple rubrics suited to your learners' level.
4. **Use Data for Learning** — turn assessment results into targeted feedback, better teaching decisions, and clear records.

### Content themes

1. **What assessment really means** — the plain definition, and why every mark you make is assessment
2. **Why we assess** — improve learning, diagnose, inform teaching, certify and select, motivate
3. **The soup and the serving** — formative vs summative, and AfL / AaL / AoL side by side
4. **Ghana's policy landscape** — who does what, NPLAF and SEAF, the 70/30 split, CA requirements, the school-based transcript
5. **The Seven Pillars** — validity, reliability, fairness, transparency, inclusivity, practicability, utility
6. **Strategies that work** — Depth of Knowledge, portfolio, debate, practical/performance, rubrics, matching strategy to level
7. **Practical workshop** — groups design a real task and rubric, then a gallery walk to share
8. **Assessment data** — feedback that moves a learner, records that hold up, and turning marks into action
9. **Common challenges** — the five things that get in the way, and what to do about them

### Source builders

| File | Purpose |
|---|---|
| `build_deck.py` | Generates the PowerPoint from `assets/` |
| `build_handout.py` | Generates the PDF handbook (ReportLab) |
| `tools/pptx_render.py` | Minimal PPTX→PNG proofing renderer with text-overflow QA |

### Rebuilding

```bash
pip install --break-system-packages python-pptx pillow reportlab pymupdf lxml
python3 build_deck.py
python3 build_handout.py

# proof the deck: renders PNGs and reports any text overflow
PPTX_QA=1 python3 tools/pptx_render.py deliverables/Effective-Assessment-Practices-KATON-2026.pptx build/proof
```

`assets/opt/` holds width-optimised JPEGs derived from the full-size PNGs in
`assets/`, and is gitignored. It does not need to be committed: `build_deck.py`
calls `ensure_assets()` on every run and regenerates any missing JPEG from its
source PNG automatically.

### Sanity checks after a rebuild

A missing `assets/opt/` used to fail silently — the build still succeeded, but
every illustration became an `[ illustration pending ]` box and the deck shrank
from ~3.9 MB to ~230 KB. After rebuilding, confirm:

- deck size is **≈3.9 MB** (not ~230 KB)
- **zero** `[ illustration pending ]` text boxes
- **51 slides** and **51 slides with speaker notes**
- `PPTX_QA=1` reports **0 overflow**

### Content sources

- **Teacher Assessment Manual and Toolkit — Handbook for Teachers**, NaCCA, Ministry of
  Education, Ghana (2024). Free at `curriculumresources.edu.gh`. Sections 1–7 and
  Appendices A–C.
- The Secondary Education Assessment Framework (SEAF) and National Pre-tertiary
  Learning and Assessment Framework (NPLAF) referenced within it.

Classroom scenarios, plain-language explanations and facilitation guidance are
original to this deck; all definitions and framework requirements follow NaCCA.
