Convert a Power BI dashboard export into an executive-ready presentation.

## Usage

```
/convert-to-deck "C:\path\to\dashboard.pdf"
/convert-to-deck "C:\path\to\dashboard.pptx"
/convert-to-deck "C:\path\to\report.pbip"
```

The argument `$ARGUMENTS` is the path to the source file.

---

## What to do

### Step 1 — Prepare

Run the extractor to pull out dashboard images and metadata:

```bash
python convert_dashboard.py $ARGUMENTS --prepare
```

If `--prepare` fails because dependencies are missing, install them first:

```bash
pip install -r requirements.txt
python convert_dashboard.py $ARGUMENTS --prepare
```

### Step 2 — Read the analysis request

```
Read temp/analysis_request.json
```

This tells you how many slides were extracted and where the images are saved.

If the source is a `.pbip` file and `temp/pbip_context.json` exists, read that too:

```
Read temp/pbip_context.json
```

### Step 3 — Analyze each dashboard image

For every slide listed in `analysis_request.json`, read the image:

```
Read temp/slide_1.png
Read temp/slide_2.png
... (all slides)
```

Act as a senior analyst advising an IT decision maker. For each slide:
- Extract every visible number, percentage, and label with exact units (13K not 13,000)
- Identify which platforms, apps, teams, or features are shown — only mention what is visible
- Generate 3 insights following the `"Bold punchy line || Supporting evidence with data"` format
- Write a headline that answers "so what?" (memorable, no data dump)

For slides with charts or tables, generate a `"chart"` spec so the builder renders clean SVG — never embed raw screenshots. See CLAUDE.md for chart spec formats.

Also generate across all slides:
- `deck_title` — compelling 5–10 word story-driven title (e.g. "Copilot Impact Confirmed: $14.7M in Value")
- `deck_subtitle` — scope line (platforms · org · date range)
- `executive_summary` — 5 synthesized findings, highest business impact first
- `recommendations` — 3–5 specific, data-grounded next steps

### Step 4 — Save insights

Write all insights to `temp/insights.json` following the schema in CLAUDE.md.

### Step 4b — Verify (MANDATORY before build)

Run the verification step to catch missing charts, vanilla headlines, and other issues:

```bash
python convert_dashboard.py --verify
```

If warnings appear:
- **Missing chart spec**: Go back and add a `"chart"` key to at least one insight on that slide
- **Generic headline**: Rewrite the headline to pass the "Would a VP forward this?" test
- **Slide count mismatch**: Check if you missed analysing any dashboard pages

Fix issues in `temp/insights.json` and re-run `--verify` until clean.

### Step 5 — Build

```bash
python convert_dashboard.py --build
```

The output file is saved as `dashboard_executive.pptx` in the same directory as the source file (or use `--output` to specify a custom path).

### Step 6 — Confirm

Tell the user:
- The output file path
- How many slides were generated
- The deck title used
