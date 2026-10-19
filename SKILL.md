---
name: "research-rolodex"
description: "Add research papers to Research Rolodex.pptx — one card per paper, organized by theme, sorted by date. Use when the user drops PDFs to add, says 'add to the Rolodex', 'file this paper', or invokes /research-rolodex."
---

# Research Rolodex Skill

## Purpose
Append one slide per paper to a growing `.pptx` file — one visual index card per paper, organized by section/theme.  No synthesis, no narrative. Just a clean, searchable reference card for every paper you read.

---

## First-Run Setup

If the deck doesn't exist yet:
```
python3 setup.py
```
This creates the folder structure, a blank deck with a title slide, and an example `config.json`.

---

## Dependencies
```
pip install python-pptx pillow pymupdf send2trash
```

---

## Workflow (Every Session)

1. **User drops a PDF** (and optionally screenshot files) into the Rolodex folder or chat.
2. **Claude runs a single extraction pass** — title, authors, journal, date, DOI, body content, FreeText annotations, highlighted text, and auto-extracts study details (type, sample size, evidence level). Study details are auto-filled and never asked of the user.
3. **Claude asks in one batched message**: theme/section name, value rating (1–10), and any personal notes or takeaways — figures auto-detected, no confirmation needed.
4. **Claude builds slides** and appends to the deck:
   - 0 screenshots → `nofig` layout (text only, full width)
   - 1 screenshot → right or wide layout depending on aspect ratio
   - 2+ screenshots → text slide + figures companion slide (consecutive)
5. Screenshot files are moved to system Trash after embedding.
6. Source PDFs moved to their theme subfolders.
7. Claude appends one row per paper to `Research Rolodex Index.csv` and regenerates `Research Rolodex.bib`.
8. Final slide count verified once.

---

## File Organization

```
Research Rolodex/
├── setup.py
├── append-slide.py
├── example-config.json
├── sections.txt              ← section names in color-palette order
├── Research Rolodex.pptx
├── Research Rolodex Index.csv
├── Research Rolodex.bib
├── Machine Learning/
│   └── [paper PDFs]
├── Clinical Studies/
│   └── [paper PDFs]
└── ...
```

---

## Paper Slide Layout

### Accent Bar
Thin color bar at the very top (full width, ~0.10" tall) in the section's palette color.

### Value Rating Badge
Small filled square (left of the title). Background: section color. Bold white number (1–10), provided by user at import time.

### Title Area
- **Line 1:** Paper title — Arial Narrow Bold 14pt, section color
- **Line 2:** Authors — Arial Narrow Regular 10pt, dark ink
- **Line 3:** Journal/citation — Consolas 9pt, gray. Format: `Journal YYYY;Vol(Issue):Pages` or `Mon YYYY`.

### Body Content
Arial Narrow 11pt, bold section-color labels:
```
Problem:    [1 sentence]
Methods:    [1 sentence]
Results:    [2 sentences, key numbers]
Conclusion: [1 sentence]
```
Adapt labels to the paper type (Scope, Framework, Argument, Study, etc.)

### Figure Area
- Prefer user-provided screenshots matching the PDF filename with a suffix (e.g. `2022 Annals Smith GoNoGo [R] - Fig2.png`)
- Fallback: extract from PDF using PyMuPDF + PIL
- **Never hallucinate figures**
- If no suitable figure exists, leave blank
- Layout adapts to aspect ratio (right column or below body)

### DOI Footnote
`DOI: [doi]` at bottom, Consolas 8pt, gray. Omit if no DOI.

---

## Speaker Notes Format

```
Upload Date: Mon YYYY

STUDY: [type], n=[N], Level [1-5]   ← auto-extracted; omit blank parts

My Notes:
• [FreeText annotation from PDF, verbatim — or blank bullet if none]

Highlights:
• [verbatim highlight 1]
• [verbatim highlight 2]
```

For papers with structured takeaways, `My Notes:` becomes `Takeaways (Mon YYYY):` instead.

**Rules:**
- Upload Date is always the first line
- Blank line between Upload Date / STUDY line and My Notes
- Blank line between My Notes section and Highlights
- My Notes always has at least one bullet (empty if no annotation text)
- Highlights are **verbatim** — word-for-word, no paraphrasing

---

## Auto-Filled Study Details

Extracted in the same pass. **Never prompt the user for these.**

| Field | Values |
|---|---|
| Study type | RCT, prospective cohort, retrospective cohort, case-control, case series, cross-sectional, systematic review/meta-analysis, model development/validation, narrative review, guideline, other |
| Sample size | `n=` as reported; blank if not stated |
| Evidence level | Oxford CEBM 1–5 inferred from study type; blank if unclear |

---

## Section Color Palette

Six colors wrap cyclically for 7+ sections — section 7 reuses color 1, section 8 reuses color 2, etc.

| Slot | Hex | Example use |
|---|---|---|
| 1 | #0085CA | Blue |
| 2 | #006747 | Green |
| 3 | #5D2584 | Purple |
| 4 | #C8102E | Red |
| 5 | #007B8F | Teal |
| 6 | #002D62 | Navy |

Section order is controlled by `THEME_ORDER` in `append-slide.py` (defaults to empty, meaning first-appearance order) or by editing `sections.txt`.

---

## Companion Index (CSV + BibTeX)

### `Research Rolodex Index.csv`
- One row per paper, appended each session (never rewrite or reorder existing rows)
- Columns: `Title, Authors, Journal, Date, DOI, PMID, Theme, Rating, Study Type, Sample Size, Evidence Level, Takeaways, Added`
- De-duplicate by DOI (or title if no DOI)

### `Research Rolodex.bib`
- BibTeX entries rebuilt from CSV after each import
- Cite key: `AuthorLastNameYYYYFirstTitleWord`

---

## Append-Only Approach

`append-slide.py` opens the existing `.pptx` and adds new slides without modifying existing ones. You can freely:
- Edit speaker notes in PowerPoint
- Rearrange or delete slides
- Adjust formatting or colors

None of these will be overwritten on the next import.

---

## Building a slide programmatically

```bash
python3 append-slide.py my-config.json
```

See `example-config.json` for the full config schema. The `deck_path` field is required; all others follow the paper schema above.
