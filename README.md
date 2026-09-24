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
| `Effective-Assessment-Practices-KATON-2026.pptx` | 36-slide editable deck, 16:9, with full speaker notes on every slide |
| `Assessment-Handout-KATON-2026.pdf` | 21-page A4 companion handbook for colleagues |
| `Presenters-Run-Sheet.md` | Timings, core-slide markers and facilitation cues |
| `Assessment-Jargon-Buster.md` | The jargon-buster section as plain markdown |

### Source builders

| File | Purpose |
|---|---|
| `build_deck.py` | Generates the PowerPoint from `assets/opt/` |
| `build_handout.py` | Generates the PDF handbook (ReportLab) |
| `tools/pptx_render.py` | Minimal PPTX→PNG proofing renderer + text-overflow QA |

### Rebuilding

```bash
pip install python-pptx pillow reportlab pymupdf lxml
python3 build_deck.py
python3 build_handout.py

# proof the deck (optional): renders PNGs and reports text overflow
PPTX_QA=1 python3 tools/pptx_render.py deliverables/Effective-Assessment-Practices-KATON-2026.pptx build/proof
```

### Content sources

- **Teacher Assessment Manual and Toolkit — Handbook for Teachers**, NaCCA, Ministry of
  Education, Ghana (2024). Free at `curriculumresources.edu.gh`. Sections 1–7 and
  Appendices A–C.
- The Secondary Education Assessment Framework (SEAF) and National Pre-tertiary
  Learning and Assessment Framework (NPLAF) referenced within it.

Classroom scenarios, plain-language explanations and facilitation guidance are
original to this deck; all definitions and framework requirements follow NaCCA.
