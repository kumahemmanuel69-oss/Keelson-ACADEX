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
| `Effective-Assessment-Practices-KATON-2026.pptx` | 40-slide editable deck, 16:9, speaker notes on every slide |
| `Assessment-Handout-KATON-2026.pdf` | 22-page A4 companion handbook for colleagues |
| `Presenters-Run-Sheet.md` | Timings, core-slide markers, facilitation cues |
| `Assessment-Jargon-Buster.md` | The glossary as plain markdown |

### Deck contents

1. **What assessment really means** — the definition, and AfL / AaL / AoL untangled
2. **The Soup and the Meal** — formative vs summative, via the soup-tasting metaphor
3. **Before you teach** — diagnostic assessment, differentiated assessment
4. **The Seven Pillars** — validity, reliability, fairness, transparency, inclusivity, practicability, utility
5. **The Jargon Buster** — rubrics, Depth of Knowledge, everyday terms
6. **Strategies that work** — portfolio, debate, practical/authentic assessment
7. **Feedback** — the good mechanic, and three-minute tools for Monday
8. **Planning** — the six-phase assessment plan, and the 70/30 split
9. **Technology** — e-assessment without a computer lab

### Source builders

| File | Purpose |
|---|---|
| `build_deck.py` | Generates the PowerPoint from `assets/opt/` |
| `build_handout.py` | Generates the PDF handbook (ReportLab) |
| `tools/pptx_render.py` | Minimal PPTX→PNG proofing renderer with text-overflow QA |

### Rebuilding

```bash
pip install python-pptx pillow reportlab pymupdf lxml
python3 build_deck.py
python3 build_handout.py

# proof the deck: renders PNGs and reports any text overflow
PPTX_QA=1 python3 tools/pptx_render.py deliverables/Effective-Assessment-Practices-KATON-2026.pptx build/proof
```

Note: `assets/opt/` holds width-optimised JPEGs derived from the full-size PNGs in
`assets/`. Regenerate with the snippet in `build_deck.py`'s header if it is missing.

### Content sources

- **Teacher Assessment Manual and Toolkit — Handbook for Teachers**, NaCCA, Ministry of
  Education, Ghana (2024). Free at `curriculumresources.edu.gh`. Sections 1–7 and
  Appendices A–C.
- The Secondary Education Assessment Framework (SEAF) and National Pre-tertiary
  Learning and Assessment Framework (NPLAF) referenced within it.

Classroom scenarios, plain-language explanations and facilitation guidance are
original to this deck; all definitions and framework requirements follow NaCCA.
