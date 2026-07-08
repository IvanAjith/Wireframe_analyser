# Wireframe Analyzer

Turns an HTML or image wireframe into a **professional Word design
specification document** — organized Block → Visual (the background card,
then every individual thing sitting on it), with exact position, size,
font, and color for each, ready to hand to a Power BI/Canva build.

## Quick start

```bash
git clone <your-repo-url>
cd wireframe-analyzer
pip install -r requirements.txt
playwright install chromium
```

macOS also needs Tesseract for OCR (image wireframes only):
`brew install tesseract` · Linux: `sudo apt install tesseract-ocr` ·
Windows: install from https://github.com/UB-Mannheim/tesseract/wiki and
make sure it's on PATH.

**Every time you want to analyze a wireframe:**

1. Drop the file into `source/` (any `.html`, `.png`, `.jpg`).
2. Run:
   ```bash
   python3 run.py your_wireframe.html
   ```
3. Open the result in `output/your_wireframe_design_spec.docx`.

That's it — same two steps on any machine, once the one-time setup above
is done.

Running it again on a file you've already documented **never overwrites**
the previous result — it writes `_v2`, `_v3`, etc. instead. Delete old
versions from `output/` yourself whenever you want to tidy up.

```bash
# a few more examples
python3 run.py dashboard.png
python3 run.py report.html --canvas 1920x1080   # override auto-detected size
python3 run.py report.html --xlsx                # also get the old Excel sheet
python3 run.py report.html --no-images            # skip screenshots, faster/smaller
```

Any flag you add after the filename is passed straight through to
`wireframe_analyzer.py` — see below for the full list.

## What it does

Give it an **HTML file** or an **image** (`.png` / `.jpg`) of a wireframe.
It automatically:

1. Detects the **Navigation** sidebar (if there is one) and the **Header**
   bar (if there is one).
2. Splits the remaining canvas into **Section 1, Section 2, ...** based on
   the whitespace gaps between columns of content — matching how the
   wireframe is actually laid out.
3. Within each section, identifies each individual **Block** (a card or
   panel) and every **Visual** sitting on it (a chart, table, KPI number,
   nav item, filter, icon, etc.), records its exact position, size,
   background color, font family, font size, and font color, and grabs a
   cropped screenshot of it.
4. **Extracts the wireframe's own color theme** (its brand/accent color)
   and reuses it throughout the document — the cover page, headings, table
   shading — so the document itself feels like it belongs to that design
   rather than a generic template.
5. Writes it all into `<filename>_design_spec.docx`:
   - A **full-bleed cover page** in the wireframe's own accent color, with
     a thumbnail of the full wireframe and a quick-stats strip.
   - A **Contents + Color Palette + Power BI Setup Guide** page: every
     section's dominant color, a swatch grid of every distinct color found
     (hex codes included, ready to paste into a Power BI custom theme),
     and a starter settings table (page size, background, accent, font).
   - Then, section by section, content organized **Block → Visual**,
     matching how you'd actually build it in Canva or Power BI: build the
     Block (the card) first, then place each Visual on top of it. When
     several visuals share one card, they're all nested under that same
     Block — you only draw the card once.

If any visual would have exceeded the canvas (causing clipping/scrolling),
every position/size in the document is automatically scaled down to fit —
noted on the cover page.

## HTML vs. Image — accuracy difference

| | HTML input | Image input |
|---|---|---|
| Section/visual boundaries | Exact (real DOM structure) | Estimated (whitespace-gap detection) |
| Position/size | Exact | Estimated |
| Background color | Exact | Estimated (~close, may be off by a shade) |
| Font size/color | Exact | Estimated (OCR text height/pixel color) |
| Font family | Exact | **Not detectable** — flagged for manual check |
| Visual titles | Exact (from source text) | OCR-read — small dashboard fonts can be misread |

For image input specifically: on a dense, real-world screenshot, a few
adjacent list rows can get grouped into one entry instead of split apart
perfectly — an inherent limit of inferring structure from pixels rather
than real markup. Always skim the document against the source before
finalizing.

## Project layout

```
wireframe-analyzer/
├── run.py                  <- the command you actually type
├── wireframe_analyzer.py   <- the analysis engine (called by run.py)
├── requirements.txt
├── source/                 <- drop wireframes here (gitignored, stays local)
└── output/                 <- generated .docx files land here (gitignored)
```

`source/` and `output/` are gitignored on purpose — your wireframes and
generated documents are local working files, not repo content. Only the
folders themselves (via `.gitkeep`) are tracked, so `git clone` always
gives you the right structure ready to use.

## Full flag reference (wireframe_analyzer.py, via run.py)

| Flag | Meaning |
|---|---|
| `--canvas WxH` | Override the auto-detected canvas size, e.g. `--canvas 1920x1080` |
| `--out PATH` | Exact output path (run.py sets this for you automatically) |
| `--outdir DIR` | Write into a specific folder with the auto-generated name (used if you call `wireframe_analyzer.py` directly instead of via `run.py`) |
| `--xlsx` | Also generate the old-style Excel spec sheet |
| `--no-images` | Skip embedded screenshots (faster, smaller file) |

You can also call `wireframe_analyzer.py` directly (bypassing `run.py`'s
`source/`/`output/`/versioning convention) if you want full control:
```bash
python3 wireframe_analyzer.py /any/path/wireframe.html --outdir /any/path/output
```

## Known limitations (image mode)

- Font family can never be reliably read from a flat image — always
  flagged `NOT DETECTABLE FROM IMAGE`. Most corporate wireframes default
  to Segoe UI (Windows) — treat that as your starting guess.
- OCR can misread stylized or very small mockup text — always sanity-check
  each visual's title against the source.
- Section/visual boundaries are inferred from whitespace gaps; an unusually
  tight or unconventional layout may need a manual tweak here and there.
