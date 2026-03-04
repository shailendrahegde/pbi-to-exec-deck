# Copilot Project Instructions

This file provides persistent instructions for Copilot Chat when working in this project.

## Project Overview

This project converts Power BI dashboard exports (.pptx or .pdf) into executive-ready analytics presentations with **compelling, analyst-grade insights** using Copilot's vision and analytical capabilities.

**Supported Input Formats:**
- PowerPoint (.pptx) - Exported from Power BI
- PDF (.pdf) - Exported from Power BI
- Power BI Project (.pbip) - Live model access via powerbi-modeling MCP
- **Note:** All formats produce identical output quality (16:9 PPTX)

## Core Philosophy

**Deterministic for Technical Tasks, Intelligent for Insights**

- ✅ **Deterministic Path:** Extract images, parse structure, render slides (fast, reliable)
- ✅ **Intelligent Path:** Analyze dashboards, identify gaps, generate strategic insights (high quality)
- ✅ **No API Key Required:** Uses current Copilot Chat session
- ✅ **No OCR Dependencies:** Works with any Power BI dashboard export

## Standard Workflow: Copilot-Powered Conversion

When a user requests dashboard conversion, you orchestrate **three steps**:

### Step 1: Extract

```bash
python convert_dashboard.py --source "<path>" --prepare --assistant copilot
```

This extracts slide images to `temp/`, parses titles and structure, and creates `temp/analysis_request.json` with slide metadata and text-layer data.

### Step 2: Analyze (You!)

**This is where you generate the insights.** No external API is needed — you ARE the LLM.

1. Read `temp/analysis_request.json`
2. For each slide, read the dashboard image (`temp/slide_N.png`)
3. Use the `text_layer` and `text_metrics` fields for extractable text/numbers
4. Generate analyst-grade insights following the formula below
5. Write results to `temp/insights.json`

### Step 3: Build

```bash
python convert_dashboard.py --build --output "<output>.pptx"
```

This loads your insights, creates professional slides (16:9 widescreen), embeds dashboard images with insights, and validates against the Constitution.

**Total time:** ~30 seconds including your analysis

---

## Your Task: Generate Analyst-Grade Insights

When the extraction completes, you should:

### 1. Read the analysis request
```
Read temp/analysis_request.json
```

### 2. For each slide, read the dashboard image
```
Read temp/slide_2.png
Read temp/slide_3.png
... (for all slides listed)
```

### 3. Act as senior analyst advisor to IT decision maker
- Identify key numbers, trends, patterns
- Spot gaps and opportunities
- Generate strategic insights with "so what"
- Provide actionable recommendations

### 4. Write insights to JSON

**IMPORTANT: After analyzing all slides, also generate:**

**Executive Summary (5 bullets):**
- Synthesize the most compelling insights across ALL dashboard pages
- Prioritize by business impact (highest value first)
- Each bullet must reference specific numbers from the data
- Answer: "What are the 5 things an executive MUST know?"
- Format: "[Specific finding] → [Business implication]"

**Next Steps (3-5 recommendations):**
- Actionable recommendations grounded in the data
- Prioritize by implementation value and feasibility
- Each must trace back to specific insights from the dashboard
- Format: "[Action verb] + [what to do] + [expected outcome]"

**Deck Title & Subtitle:**
Generate a compelling `deck_title` that captures the core story of the deck — the single most important narrative an executive should remember. It should be positively framed and attention-grabbing.

Format: `"[Evocative phrase]: [Core finding or story]"`

Examples of great deck titles:
- "Copilot Impact Confirmed: $14.7M in Value with 84% Active Adoption"
- "From Reach to Routine: Building AI Habits Across the Organization"
- "Momentum Confirmed: Three Platforms, One Direction, Zero Plateau"

The `deck_subtitle` describes scope: what platforms/products are covered, the company/org name if known, and the time period.

Format: `"[Platform 1] · [Platform 2] · [Platform 3] · [Month Year] – [Month Year]"`

Examples:
- "Agents · Unlicensed Chat · M365 Copilot · Mar – Jun 2025"
- "Microsoft 365 Copilot Impact Report · Apr–Oct 2025"

## SVG Chart Rule (CRITICAL — applies to ALL source types)

> **Always generate chart specs from extracted data. Never embed raw screenshots.**

The builder has two rendering modes. You MUST use the chart mode:

| Mode | Triggered when | Visual result |
|---|---|---|
| **Chart (SVG)** | Each `insight` object has a `"chart"` key | Polished rendered chart on the slide |
| **Screenshot fallback** | Insights are plain strings with no `"chart"` key | Raw PBI screenshot pasted as-is |

**Use chart mode for every slide where you can extract at least one data series.**
Only fall back to plain-string mode (screenshot) if the slide has no quantitative data at all (e.g. a text-only guidance page or navigation slide).

### How to extract chart data from a screenshot

For every slide with a visible chart or table:
1. Read the axis labels, legend items, and bar/line heights carefully
2. Transcribe each series as `{"label": "...", "value": N}` entries
3. Choose the appropriate chart type (see table below)
4. Set `"title"` to a short label matching the dashboard chart title

You do NOT need exact precision — reasonable visual estimates are fine. The goal is a clean rendered chart, not perfect numbers. If a bar reaches roughly 60% of the axis max, use that value.

### Chart types supported

| Dashboard visual | Use `type:` |
|---|---|
| Horizontal bars (manager leaderboard, ranked list) | `"bar"` |
| Vertical columns (time series, category comparison) | `"column"` |
| Line chart (trend over time) | `"line"` |
| Donut / pie | `"donut"` |
| KPI card (single big number) | `"kpi"` |
| Multiple KPI cards in a row | `"kpi_row"` |
| Data table / matrix | `"table"` |
| Treemap | `"treemap"` |
| Scatter / bubble | `"scatter"` |
| Funnel | `"funnel"` |
| Gauge | `"gauge"` |

### Per-insight chart spec format

Each insight bullet can optionally carry a chart. The first insight with a chart
is rendered as the main chart for that slide. A second chart (if provided) appears
in an expanded two-chart layout.

```json
{
  "text": "Bold punchy line || Supporting evidence with data",
  "chart": {
    "type": "bar",
    "title": "Weekly Actions by Manager",
    "data": [
      {"label": "Dana Bourque", "value": 48.95},
      {"label": "Matt Sheard",  "value": 38.09},
      {"label": "Saurabh Pant", "value": 33.62},
      {"label": "Jason Kim",    "value": 29.53}
    ]
  }
}
```

For a KPI card:
```json
{
  "text": "5.35 hours saved per person per week || ...",
  "chart": {"type": "kpi", "value": "5.35", "label": "Assisted Hours/Week/Person"}
}
```

For `kpi_row` (multiple KPI cards side-by-side):
```json
{
  "chart": {
    "type": "kpi_row",
    "items": [
      {"value": "256",   "label": "Active Users"},
      {"value": "33.77", "label": "Weekly Actions"},
      {"value": "52.6%", "label": "Power Users"},
      {"value": "58.5%", "label": "Growth"}
    ]
  }
}
```

For a table:
```json
{
  "chart": {
    "type": "table",
    "columns": ["Manager", "Users", "Weekly Actions"],
    "rows": [
      ["Dana Bourque", "4", "48.95"],
      ["Matt Sheard",  "52", "38.09"]
    ]
  }
}
```

### Full slide insight format with charts

```json
{
  "slide_number": 3,
  "title": "Usage Trend",
  "headline": "Dana Bourque team sets the intensity benchmark at 48.95 weekly actions",
  "insights": [
    {
      "text": "Dana Bourque leads at 48.95 actions/user || Small focused cohort of 4 proving the Power User ceiling is reachable",
      "chart": {
        "type": "bar",
        "title": "Weekly Actions by Manager",
        "data": [
          {"label": "Dana Bourque",  "value": 48.95},
          {"label": "Matt Sheard",   "value": 38.09},
          {"label": "Saurabh Pant",  "value": 33.62},
          {"label": "Jason Kim",     "value": 29.53}
        ]
      }
    },
    {
      "text": "30.76 avg weekly actions across the cohort || 3.44 active days/week and 2.54 apps signals multi-surface habit formation",
      "chart": null
    },
    {
      "text": "Jason Kim anchors 151 users at 29.53 || Largest group still above Power User threshold — strong org-wide floor",
      "chart": null
    }
  ],
  "numbers_used": ["48.95", "4", "38.09", "52", "33.62", "63", "29.53", "151"]
}
```

When using charts, the `insights` field is a **list of objects** (each with `"text"` and `"chart"`), not plain strings.

---

## Complete JSON Output Schema

```json
{
  "deck_title": "Compelling story-driven title",
  "deck_subtitle": "Agents · Chat · M365 Copilot · Month Year – Month Year",
  "executive_summary": [
    "Finding with specific number → business implication",
    "Finding with specific number → business implication",
    "Finding with specific number → business implication",
    "Finding with specific number → business implication",
    "Finding with specific number → business implication"
  ],
  "recommendations": [
    "Action: Specific recommendation with expected outcome",
    "Action: Specific recommendation with expected outcome",
    "Action: Specific recommendation with expected outcome"
  ],
  "slides": [
    {
      "slide_number": 1,
      "title": "Slide Title",
      "headline": "Insight-driven headline",
      "insights": [
        {
          "text": "Punchy bold line || Supporting evidence with specific data",
          "chart": {"type": "bar", "title": "Chart Title", "data": [{"label": "A", "value": 42}]}
        },
        {"text": "Second insight || Detail", "chart": null},
        {"text": "Third insight || Detail", "chart": null}
      ],
      "numbers_used": ["134", "1,275", "11%"]
    }
  ]
}
```

**Save to:** `temp/insights.json`

---

## Using Text Layer Data

When running with `--assistant copilot`, the extraction enriches each slide with text-layer data:

- `text_layer`: Raw text extracted from the PPTX/PDF (visual titles, labels, chart types)
- `text_metrics`: Extracted numeric values with context (percentages, counts, etc.)
- `text_key_phrases`: Key phrases found in the text

**How to use:**
- Cross-reference `text_layer` with the slide image for chart titles and axis labels
- Use `text_metrics` values when visible in the image — they provide confirmed numbers
- Use `text_key_phrases` for context about what the slide covers
- **Always verify text-layer data against the image** — the image is the source of truth

---

## PBIP Workflow (when source is a .pbip file)

When a user has a `.pbip` Power BI project open in Power BI Desktop, this
path queries the **live in-memory model** directly — no screenshots needed.

```bash
python convert_dashboard.py --source "MyReport.pbip" --prepare --assistant copilot
```

### What's different from the standard workflow

| | Standard (PDF/PPTX) | PBIP |
|---|---|---|
| Data source | Copilot reads images | Copilot queries live model via MCP |
| Accuracy | Visual estimation | Exact values from DAX |
| Depth | What's visible on screen | Can query any dimension/filter |
| Measure context | Unknown | Full DAX expression available |

### When `temp/pbip_context.json` exists alongside `analysis_request.json`:

**Step 1: Read the context file**
```
Read temp/pbip_context.json
```
This contains:
- `pages` — report page structure with visual types and field bindings
- `model.measures` — every measure with its full DAX expression
- `model.tables` — table and column structure
- `model.relationships` — table relationships
- `dax_queries` — pre-built DAX queries, one group per page

**Step 2: Execute DAX queries via the `powerbi-modeling` MCP**

For each page in `dax_queries`, run the queries using the MCP tool:
- The MCP connects to the running Power BI Desktop instance
- The returned table rows are your data source — use exact values
- State the measure name alongside the value for traceability

**Step 3: Drill deeper with custom DAX (optional)**

Modify pre-built queries to apply filters or explore dimensions:
```dax
EVALUATE
CALCULATETABLE(
    SUMMARIZECOLUMNS('Date'[Month], "Revenue", [Total Revenue]),
    DATESINPERIOD('Date'[Date], TODAY(), -30, DAY)
)
```

**Step 4: Use `measure_dax` to understand HOW each KPI is calculated**
- Measure uses `DATESYTD()` → note it's a year-to-date figure in insights
- Measure uses `DIVIDE()` → watch for zero-division handling
- Measure uses `CALCULATE()` with a filter → be explicit about what filter applies

**Step 5: Generate insights from queried values — NOT from estimation**
- Every number in insights must come from an executed DAX query result
- Follow the same insight formula as the standard workflow
- Reference measure names in parentheses: "47.3% completion rate ([Project Completion Rate])"

**Step 6: Write insights to `temp/insights.json`** (same format as always)

Then build:
```bash
python convert_dashboard.py --build --output "<output>.pptx"
```

### If DAX queries return no data

- Power BI Desktop may not be open, or the MCP may not be connected
- Fall back to any companion images if available (`temp/slide_N.png`)
- If no images exist either, mark slide as "Insufficient data for analysis"

---

## Deck Title Guidelines

**ALWAYS generate `deck_title` and `deck_subtitle`.** Never leave these blank — the default "Executive Insights" is generic and forgettable.

**`deck_title`** — The cover slide headline. Should be:
- Compelling and persuasive, not descriptive
- Positive framing: opportunities and wins, not gaps or problems
- Specific to the content: reference the org, product, or key finding
- 5–10 words max

**Good examples:**
- ✅ "Swetha Org: Strong Copilot Momentum with Clear Scaling Opportunities"
- ✅ "M365 Copilot: 40% Power Users and Growing"
- ✅ "Copilot Adoption Accelerating — Efficiency and Wellbeing Gains Confirmed"

**Bad examples:**
- ❌ "Executive Insights" (generic default)
- ❌ "Copilot Usage Report" (descriptive, not persuasive)
- ❌ "Low Adoption in Some Teams" (negative framing)

**`deck_subtitle`** — One line of context. Use the date range from the dashboard filters and/or the org/leader name.
- Example: "Swetha Org · Dec 2025 – Feb 2026"
- Example: "Super User Analysis · Q1 2026"

---

## Insight Generation Guidelines

**CRITICAL: Before analyzing any dashboard, read `DASHBOARD_READING_RULES.md` for data extraction accuracy rules (if available).**

When generating insights (Step 2), follow these principles:

### ✅ DO: Be Concise and Friendly

**Good Example:**
- Headline: "134 Agent users from 1,275 total (11%) - significant opportunity to expand automation adoption"
- Insight: "HR Generalists (140 actions/user) are 3-4x above average - great candidates to champion agent adoption"

### ✅ DO: Focus on Opportunities

Frame observations as opportunities rather than problems:
- ✅ "Opportunity to expand agent adoption from 11% to 30%"
- ❌ "Only 11% adoption shows deployment failure"

### ✅ DO: Provide Specific Numbers (Match Units Exactly)

Every headline and at least one insight should include concrete numbers:
- ✅ "217 prompts per user (5.7x average) demonstrates high-value workflow"
- ❌ "Some users show higher engagement than others"

**CRITICAL: Match number units EXACTLY as shown in the visual:**
- Dashboard shows "13" → Use "13 actions" (NOT "13K")
- Dashboard shows "13K" → Use "13K actions" (NOT "13,000" or "13")
- Dashboard shows "13M" → Use "13M actions" (NOT "13 million")

### ✅ DO: Answer "So What?"

Every insight should help IT decision maker take action:
- ✅ "Client Finance's 217 prompts/user pattern ready to replicate across Corporate Finance"
- ❌ "Client Finance has 217 prompts per user"

### ✅ DO: Analyze Platform/Feature Patterns

**CRITICAL: Only mention entities (platforms, teams, features) that are VISIBLE on the dashboard page.**

- ✅ "Teams integration shows 3x higher engagement than Outlook" — **ONLY if both are visible**
- ❌ Mentioning "Teams" when it's not shown on this specific page

**Rule: Can you point to where this name appears on the dashboard?**
- **YES** → Safe to mention
- **NO** → Do NOT mention

### ❌ DON'T: Be Critical or Verbose

Avoid:
- Critical language: "critical failure", "deployment disaster", "massive gap"
- Verbose explanations: Keep insights to 1-2 sentences max
- Technical jargon: Use business language

### ❌ DON'T: Make Generic / Vanilla Statements

An executive reading your deck should never think "I already knew that" or "so what?" Generic observations destroy credibility and waste slide real-estate.

**The "So What" test:** Before writing any headline or insight, ask: *"Would a VP forward this bullet to their boss?"* If the answer is no, rewrite it.

**Vanilla patterns to NEVER use:**
- ❌ "There are 1,275 active users" → ✅ "1,275 active users represent 68% penetration — 480-seat expansion opportunity remains"
- ❌ "Platform shows engagement" → ✅ "Power users average 33 actions/week — 2.5× the org baseline"
- ❌ "Usage varies by department" → ✅ "Operations outpaces Marketing 3:1 on weekly actions — replicate their onboarding playbook"
- ❌ "Four groups form the org structure" → ✅ "Top quartile (4 teams, 38% of users) generates 71% of all Copilot actions — concentrate enablement here"
- ❌ "The report covers AI skills" → ✅ "Business Management dominates confirmed skills at 45% — AI-specific skill gaps signal training opportunity"
- ❌ "Skills are distributed across categories" → ✅ "3 of 8 skill categories account for 80% of confirmations — focus L&D investment on the long tail"

**Anti-vanilla checklist for every headline:**
1. Does it contain a specific number? (If not, add one)
2. Does it imply an action or decision? (If not, add "→ [implication]")
3. Could it apply to ANY dashboard? (If yes, make it specific to THIS data)
4. Would an executive remember it tomorrow? (If not, sharpen it)

### ❌ DON'T: Force Insights Without Data

If a dashboard page has no numbers or is blank:
- ✅ Mark as "Insufficient data for analysis"
- ❌ Generate vague insights to fill space

---

## Insight Formula

**Headline:** `[Clear Takeaway Message]` — Focus on instant clarity

**Insight format: two-part, separated by `||`**

`"[Bold punchy line, 6-8 words max] || [Supporting evidence with specific data]"`

**Insight 1 (Scale/Scope):**
`"HR Generalists lead usage by far || 140 actions/user is 3-4x the average — strong candidates to champion adoption"`

**Insight 2 (Pattern/Opportunity):**
`"Client Finance pattern ready to replicate || 217 prompts/user sets the benchmark for Corporate Finance expansion"`

**Insight 3 (Action):**
`"Pilot training with top 3 departments || Target 50+ prompts/user baseline to establish org-wide standard"`

---

## Quality Validation

After generating insights, perform these mandatory checks:

### Checklist

✅ Every headline has specific number (not "some users" or "many")
✅ Number units match exactly (13 vs 13K vs 13M - use what's shown)
✅ All entities are visible (platforms, teams, features mentioned are on this page)
✅ Insights are concise (1-2 sentences each, not paragraphs)
✅ Tone is friendly (opportunities, not failures)
✅ Focus on action (what to do, not just what is)
✅ No forced insights (if no data, mark "Insufficient data")
✅ No vanilla statements (every headline passes the "would a VP forward this?" test)
✅ Executive summary synthesizes ALL pages (not just first slide)
✅ Recommendations are actionable (specific actions, not vague suggestions)

### Chart Coverage Self-Check (MANDATORY)

Before writing `temp/insights.json`, verify that **every slide with quantitative data has at least one chart spec**. Slides without chart specs fall back to pasting the raw screenshot — which defeats the purpose of the executive deck.

**For each slide ask:** Does this page show numbers, bars, lines, donuts, tables, or KPIs?
- **YES** → At least one insight MUST carry a `"chart"` key with extracted data
- **NO** (text-only guidance page) → OK to omit chart; use plain string insights

**Common chart-miss patterns:**
- KPI cards often get described in text but not encoded as `"kpi"` or `"kpi_row"` charts
- Tables get summarised as text instead of encoded as `"table"` chart spec
- Small bar/column charts at the bottom of a page are overlooked

The build step now runs `verify_insights()` automatically and will print warnings for any slide missing a chart spec. You can also run verification standalone:
```bash
python convert_dashboard.py --verify
```

### OCR Cross-Verification (when `ocr_used: true`)

When `analysis_request.json` marks a slide with `"ocr_used": true`, the `text_layer` was generated by EasyOCR — not from embedded text. OCR text is spatially grouped by row, but label-to-number associations can still be ambiguous.

**MANDATORY:** For every number you cite from an OCR-enriched slide, verify the number-to-label pairing by reading the actual image. Do NOT blindly trust `text_metrics[].context` — cross-check it visually.

**Example of what goes wrong:**
- OCR `text_layer`: `Business Management · 151 · AI Skills · 89`
- If you read left-to-right: 151 belongs to Business Management, 89 to AI Skills
- If you just grab "151" and see "AI Skills" nearby → WRONG: "151 people have AI Skills"
- **Always verify against the image** before attributing a number to a label

---

## Example Workflow in Practice

**User says:** "Convert wpp22.pptx to executive deck"

**You do:**

1. Run extract command:
   ```bash
   python convert_dashboard.py --source wpp22.pptx --prepare --assistant copilot
   ```

2. Read analysis request and view images:
   ```
   Read temp/analysis_request.json
   Read temp/slide_2.png
   Read temp/slide_3.png
   ... (view all slides)
   ```

3. Generate concise, friendly insights for each slide following the formula

4. Save insights JSON to `temp/insights.json`

5. Build final presentation:
   ```bash
   python convert_dashboard.py --build --output wpp22_executive.pptx
   ```

6. Verify output and report success

**Total time:** < 30 seconds
**Quality:** Analyst-grade insights, professional appearance

---

## Files Reference

- `convert_dashboard.py` - Main conversion orchestrator
- `Claude PowerPoint Constitution.md` - Quality standards and governance
- `Example-Storyboard-Analytics.pptx` - Visual template reference (use for styling)
- `lib/rendering/builder.py` - Slide rendering (16:9 format)
- `lib/rendering/validator.py` - Constitution compliance checker
- `temp/` directory - Temporary files for analysis (auto-generated, not committed)

---

## Error Handling

If conversion fails:

1. **Check file paths** - Use absolute paths, proper quotes
2. **Check temp/ directory** - Should contain extracted images
3. **Check insights JSON** - Verify format and completeness
4. **Re-run failed step** - Extract, analyze, or build independently
5. **Report clearly** - State which step failed and why
