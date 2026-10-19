# Research Rolodex

**A tool from the [Program for Surgical Intelligence](https://surgicalintelligence.us)**

A growing PowerPoint reference deck — one visual index card per paper, organized by section, sorted by date within each section. No synthesis, no narrative. Just a clean, searchable card for every paper you read.

Built to be driven by Claude (or any LLM that can read PDFs and write JSON), but works fine manually too.

---

## What it looks like

Each paper gets one slide:

- **PSI phase-recognition bar** at the very top (the signature 6-segment motif)
- **Value rating badge** (1–10) to the left of the title
- **Title · Authors · Journal/citation** in the header
- **Structured body** — Problem / Methods / Results / Conclusion (or whatever labels fit the paper)
- **Figure** if you have a screenshot (right column or full-width below, depending on aspect ratio)
- **DOI footnote** at the bottom
- **PSI footer** (logo mark + "Program for Surgical Intelligence") at bottom-left
- **Speaker notes** with Upload Date, study details, your personal notes, and verbatim highlights — hidden from the slide face, searchable in PowerPoint

Section divider slides separate each topic. The color palette wraps automatically if you have more than 6 sections (section 7 reuses color 1, etc.).

---

## Setup

### 1. Install dependencies

```bash
pip install python-pptx pillow pymupdf send2trash
```

### 2. Run first-time setup

```bash
python3 setup.py
```

This will ask you where to keep your Rolodex, what to call your sections, and create the folder structure and a blank deck with a title slide.

### 3. Add your first paper

Edit the `example-config.json` that setup.py creates, then:

```bash
python3 append-slide.py example-config.json
```

---

## Config schema

```json
{
  "deck_path":   "/path/to/Research Rolodex/Research Rolodex.pptx",
  "theme":       "Machine Learning",
  "layout":      "nofig",
  "footer_text": "",

  "paper": {
    "title":    "Paper title",
    "date":     "Jan 2024",
    "authors":  "Smith J, Jones A, et al.",
    "doi":      "10.1234/example",
    "citation": "Journal Name 2024;1(1):1–10",
    "rating":   8,

    "body": [
      {"label": "Problem",    "text": "One sentence."},
      {"label": "Methods",    "text": "One sentence."},
      {"label": "Results",    "text": "Two sentences with key numbers."},
      {"label": "Conclusion", "text": "One sentence."}
    ],

    "notes": "Upload Date: Jan 2024\n\nMy Notes:\n• \n\nHighlights:\n• Key finding."
  }
}
```

**Layouts:** `nofig` · `wide` (figure below body) · `right` (figure right column) · `figs_only` (figures companion slide)

**Multi-figure:** For 2+ screenshots, create two configs — one `nofig` and one `figs_only` — or use the `second_slide` key (see `example-config.json`).

---

## Speaker notes format

```
Upload Date: Mon YYYY

STUDY: [type], n=[N], Level [1-5]

My Notes:
• [your annotation or blank]

Highlights:
• [verbatim highlight from the paper]
```

`STUDY:` is auto-filled by Claude from the abstract; never prompted. For papers with structured personal takeaways, `My Notes:` becomes `Takeaways (Mon YYYY):`.

---

## Claude integration

Drop `SKILL.md` into your Claude project or paste it into your system prompt. Claude will:

1. Read the PDF (title, authors, journal, DOI, body content, annotations, highlights, study details)
2. Ask you in one message: which section, value rating (1–10), and any personal notes
3. Build and run the config — including a second slide if you have 2+ figures
4. Move screenshot files to Trash after embedding
5. Move the PDF into its section subfolder
6. Append a row to `Research Rolodex Index.csv` and regenerate `Research Rolodex.bib`

---

## Section colors

Six colors, wrapping for 7+ sections:

| Slot | Hex | Default |
|------|-----|---------|
| 1 | `#0085CA` | Blue |
| 2 | `#006747` | Green |
| 3 | `#5D2584` | Purple |
| 4 | `#C8102E` | Red |
| 5 | `#007B8F` | Teal |
| 6 | `#002D62` | Navy |

Section color assignment follows `THEME_ORDER` in `append-slide.py` (defaults to first-appearance order). Edit that list or `sections.txt` to lock in a specific order.

---

## Append-only

`append-slide.py` opens the existing `.pptx` and adds new slides without touching any existing ones. Edit notes, rearrange slides, adjust formatting — none of it will be overwritten on the next append.

---

## File layout

```
Research Rolodex/
├── setup.py
├── append-slide.py
├── example-config.json
├── sections.txt
├── Research Rolodex.pptx
├── Research Rolodex Index.csv
├── Research Rolodex.bib
├── Machine Learning/
├── Clinical Studies/
└── ...
```

---

## License

MIT
