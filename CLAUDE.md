# Claude Project Instructions

This file provides persistent instructions for Claude when working in this project.

## Project Overview

This project converts Power BI dashboard exports (.pptx or .pdf) into executive-ready analytics presentations with **compelling, analyst-grade insights** using Claude's vision and analytical capabilities.

**Supported Input Formats:**
- PowerPoint (.pptx) - Exported from Power BI
- PDF (.pdf) - Exported from Power BI
- Power BI Project (.pbip) - Live model access via powerbi-modeling MCP
- **Note:** All formats produce identical output quality (16:9 PPTX)

## Core Philosophy

**Deterministic for Technical Tasks, Intelligent for Insights**

- ✅ **Deterministic Path:** Extract images, parse structure, render slides (fast, reliable)
- ✅ **Intelligent Path:** Analyze dashboards, identify gaps, generate strategic insights (high quality)
- ✅ **No API Key Required:** Uses current Claude Code session
- ✅ **No OCR Dependencies:** Works with any Power BI dashboard export

## Standard Workflow: Claude-Powered Conversion

When a user requests dashboard conversion, they run a **single command** that orchestrates all steps:

```bash
python convert_dashboard.py "dashboard.pptx"
```

The output file will be automatically named `dashboard_executive.pptx` (or use `--output` for a custom name).

### What Happens Automatically:

**Step 1: Extract (5 seconds)**
- Script extracts each dashboard slide/page as PNG image to `temp/` directory
- **For PPTX:** Extracts embedded images from slides
- **For PDF:** Converts each page to PNG image at 150 DPI
- **PPTX workflow:** Automatically skips first slide (title page with metadata like "last refreshed", "view in PowerBI")
- **PDF workflow:** Includes all pages (PDF exports typically have dashboard content on page 1)
- Parses slide/page titles and structure
- Creates `temp/analysis_request.json` with slide metadata

**Step 2: Claude Analysis (You!)**
- Script displays clear request for Claude to analyze dashboards
- You (Claude Code) automatically see the request and respond
- **Your task:** Analyze each dashboard image and generate insights
- **Save results to:** `temp/insights.json`

**Step 3: Build (3 seconds)**
- After you finish, user presses Enter
- Script loads your insights from JSON
- **Default**: Embeds PBI page screenshots with insights commentary
- **Optional**: Use `--vector-charts` flag for matplotlib-rendered charts
- Creates professional slides (16:9 widescreen)
- Applies Analytics template styling
- Validates against Constitution

**Total time:** ~30 seconds including your analysis

---

## Your Task: Generate Analyst-Grade Insights

When the script displays the analysis request, you should:

### 1. Read the analysis request
```
Read temp/analysis_request.json
```

### 2. For each slide, read the dashboard image
```
Read temp/slide_1.png
Read temp/slide_2.png
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
- Be specific: "Pilot agent training with HR Generalists (140 actions/user) to establish best practices" NOT "Improve training"

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

## Rendering Mode: Screenshots vs Vector Charts

> **Default: PBI page screenshots. Use `--vector-charts` to generate matplotlib vector graphics instead.**

The builder has two rendering modes:

| Mode | Default for | Triggered by | Visual result |
|---|---|---|---|
| **Screenshot (default)** | PBIP, PBIX sources | Default behaviour | Original PBI page capture embedded on slide |
| **Vector charts** | PPTX, PDF sources | `--vector-charts` flag | matplotlib-rendered charts from insight chart specs |

**Screenshot mode** (default for PBIP/PBIX) embeds the actual Power BI page capture on the left with insight commentary on the right. This preserves the original dashboard visuals exactly as they appear in Power BI.

**Vector chart mode** (`--vector-charts`) uses chart specs from insights.json to render polished matplotlib charts. Use this when you want clean, re-styled charts instead of raw screenshots.

### When to use `--vector-charts`

- You have **exact DAX-queried data** (via MCP) and want polished re-rendered charts
- The PBI screenshots are low quality or cluttered
- You need to highlight specific data points or change chart types

### When to keep the default (screenshots)

- The PBI dashboard is well-designed and visually clean
- You want to preserve the original dashboard formatting exactly
- The dashboard has complex visuals (dual-axis, heatmaps, scatter plots) that are hard to reconstruct

### Chart spec guidance (only needed with `--vector-charts`)

When generating insights.json for vector-chart mode, include `"chart"` specs in insight objects.
When generating for screenshot mode (default), set `"chart": null` on all insights — the builder uses page captures automatically.

### How to extract chart data from a screenshot

For every slide with a visible chart or table:
1. Read the axis labels, legend items, and bar/line heights carefully
2. Transcribe each series as `{"label": "...", "value": N}` entries
3. Choose the appropriate chart type (see table below)
4. Set `"title"` to a short label matching the dashboard chart title

You do NOT need exact precision — reasonable visual estimates are fine. The goal is a clean rendered chart, not perfect numbers. If a bar reaches roughly 60% of the axis max, use that value.

### Metric Segment Isolation Rule (CRITICAL)

> **Never mix numbers from different platform segments or visual panels.**

Dashboard pages often show multiple segments side-by-side (e.g. Licensed vs Unlicensed vs Agent). OCR text and `text_metrics` flatten all of these into a single stream, destroying the spatial grouping. This causes **cross-contamination** — attributing a metric from one segment to another.

**Before assigning any number to a chart series, verify these three things:**

1. **Which visual does this number belong to?** Check `pbip_context.json` — each visual has a `visual_type` and listed `measures[]`. A number for `NoOfActiveChatUsers (Licensed)` cannot appear in a chart labelled "Unlicensed".
2. **Which panel/region of the page is it in?** If the page has left/right or top/bottom panels for different segments, numbers from the left panel do not belong in the right panel's chart.
3. **Does the measure name match the chart series label?** The measure name in `pbip_context.json` tells you exactly which KPI a visual renders. Use this as ground truth.

**Concrete example of what goes wrong:**
- Page shows 6 KPI cards: `180 Agent Users | 3.60 Agent Sessions/User | 2,585 Chat Users | 3.3 Chat Sessions/User | 1,616 Copilot Users | 3.1 Copilot Sessions/User`
- OCR stream: `180 · 3.60 · 2,585 · 3.3 · 1,616 · 3.1`
- ❌ WRONG: Assigning `3.3` to Agent Sessions and `3.60` to Chat Sessions (shifted by one position)
- ✅ RIGHT: Reading the label row below each number to verify `3.60` pairs with "Agent Sessions Per User"

**When `pbip_context.json` is available**, always cross-reference:
- `visual_type` → determines chart type
- `measures[].name` → determines which metric a visual shows
- `measures[].entity` → confirms the data table (e.g. "Chat + Agent Interactions")
- If a measure name contains "Licensed" → only use in Licensed charts
- If a measure name contains "Unlicensed" → only use in Unlicensed charts
- If a measure name contains "Agent" → only use in Agent charts

### Chart Fidelity Tiers (when to use screenshot fallback)

Not all chart types can be faithfully reconstructed from OCR-estimated data.
Use this confidence tier to decide **vector chart vs. screenshot fallback**:

| Tier | Chart types | Action |
|---|---|---|
| **High fidelity** — reconstruct as vector | `kpi`, `kpi_row`, `bar`, `column`, `donut`, `pie`, `table` | Always generate chart spec |
| **Medium fidelity** — reconstruct with care | `line`, `area`, `column_stacked`, `bar_stacked`, `treemap`, `funnel` | Generate chart spec; verify series shape matches original |
| **Low fidelity** — prefer screenshot | `scatter` (many overlapping points), `combo` (dual-axis scale-sensitive), `heatmap`, `ribbon`, multi-axis area/line | Use `"chart": null` and let builder paste original screenshot **UNLESS** you have exact DAX-queried data |

**Rule: If you cannot confidently reconstruct the axis scales, series relationships, or point positions, use screenshot fallback.** A screenshot preserves the original chart perfectly. A bad vector chart actively misleads the reader.

To force screenshot fallback for a specific slide, set ALL insight objects' `"chart"` to `null`:
```json
{"text": "Insight text here || Supporting detail", "chart": null}
```

### Multi-Axis Chart Rule

When a dashboard visual has **two Y-axes** (e.g. left axis = Users, right axis = Sessions), this is a **dual-axis chart**. These require special handling:

1. **Identify dual-axis visuals** from `pbip_context.json`: if a visual has measures with fundamentally different units (counts vs rates, users vs sessions), it's likely dual-axis
2. **Use `"combo"` type** for dual-axis charts — it renders bars on the left Y-axis and lines on the right Y-axis via `twinx()`
3. **If the original uses area+line or line+line dual-axis**, use screenshot fallback — the `area` and `line` renderers do NOT support secondary axes
4. **Never plot both series on the same single axis** if their scales differ by >3× — this compresses the smaller series into a flat line

**Dual-axis JSON pattern:**
```json
{
  "chart": {
    "type": "combo",
    "title": "Active Users vs Sessions/User",
    "y_label": "Active Users",
    "x_label": "Sessions/User",
    "data": [
      {"label": "Mar", "value": 142},
      {"label": "Apr", "value": 164},
      {"label": "May", "value": 192}
    ],
    "series": [
      {
        "name": "Sessions/User",
        "points": [{"x": "Mar", "y": 2.6}, {"x": "Apr", "y": 3.9}, {"x": "May", "y": 5.5}]
      }
    ]
  }
}
```
> `y_label` = left axis label, `x_label` = right axis label (reused field).
> `data` = bar series (left axis), `series` = line overlay (right axis).

### Chart types supported

| Dashboard visual | Use `type:` |
|---|---|
| Horizontal bars (manager leaderboard, ranked list) | `"bar"` |
| Stacked horizontal bars | `"bar_stacked"` |
| 100% stacked horizontal bars | `"bar_stacked_100"` |
| Vertical columns (time series, category comparison) | `"column"` |
| Stacked vertical columns | `"column_stacked"` |
| 100% stacked vertical columns | `"column_stacked_100"` |
| Line chart (trend over time) | `"line"` |
| Area chart (filled line) | `"area"` |
| Combo chart (columns + line overlay) | `"combo"` or `"column_line"` |
| Waterfall (incremental +/− changes) | `"waterfall"` |
| Ribbon chart (rank changes over time) | `"ribbon"` |
| Pie chart | `"pie"` |
| Donut / doughnut | `"donut"` |
| KPI card (single big number) | `"kpi"` or `"card"` |
| Multiple KPI cards in a row | `"kpi_row"` or `"multi_row_card"` |
| Data table / matrix | `"table"` |
| Heatmap (color-coded matrix) | `"heatmap"` |
| Treemap | `"treemap"` |
| Scatter plot | `"scatter"` |
| Bubble chart | `"bubble"` |
| Radar / spider | `"radar"` |
| Funnel | `"funnel"` |
| Gauge (half-donut) | `"gauge"` |

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

For a waterfall chart (incremental changes):
```json
{
  "chart": {
    "type": "waterfall",
    "title": "Revenue Bridge Q1→Q2",
    "data": [
      {"label": "Q1 Revenue", "value": 100, "color": "#003278"},
      {"label": "New Sales",  "value": 30},
      {"label": "Churn",      "value": -12},
      {"label": "Upsells",    "value": 8},
      {"label": "Q2 Total",   "value": 126, "color": "#003278"}
    ]
  }
}
```
> Positive values = green increase bars, negative = red decrease bars.
> Items with explicit `"color"` or labels containing "total"/"net" anchor at zero.

For a combo chart (columns + line overlay):
```json
{
  "chart": {
    "type": "combo",
    "title": "Actions vs Adoption Rate",
    "data": [
      {"label": "Jan", "value": 120},
      {"label": "Feb", "value": 145},
      {"label": "Mar", "value": 160}
    ],
    "series": [
      {
        "name": "Adoption %",
        "points": [
          {"x": "Jan", "y": 45.2},
          {"x": "Feb", "y": 52.1},
          {"x": "Mar", "y": 61.8}
        ]
      }
    ]
  }
}
```
> `data` → column bars (left Y-axis), `series` → line overlay (right Y-axis).

For a scatter plot:
```json
{
  "chart": {
    "type": "scatter",
    "title": "Usage vs Satisfaction",
    "x_label": "Weekly Actions",
    "y_label": "Satisfaction Score",
    "series": [
      {"name": "Team A", "x": 45, "y": 8.2},
      {"name": "Team B", "x": 32, "y": 7.5, "highlight": true},
      {"name": "Team C", "x": 60, "y": 9.1}
    ]
  }
}
```

For a line chart:
```json
{
  "chart": {
    "type": "line",
    "title": "Monthly Active Users",
    "series": [
      {
        "name": "Agents",
        "points": [{"x": "Jan", "y": 120}, {"x": "Feb", "y": 145}, {"x": "Mar", "y": 180}]
      },
      {
        "name": "Chat",
        "points": [{"x": "Jan", "y": 80}, {"x": "Feb", "y": 95}, {"x": "Mar", "y": 110}]
      }
    ]
  }
}
```

For a donut chart:
```json
{
  "chart": {
    "type": "donut",
    "title": "Usage Distribution",
    "data": [
      {"label": "Power Users", "value": 52.6},
      {"label": "Regular",     "value": 31.2},
      {"label": "Light",       "value": 16.2}
    ]
  }
}
```

For a treemap:
```json
{
  "chart": {
    "type": "treemap",
    "title": "Feature Usage Share",
    "data": [
      {"label": "Email Drafts", "value": 340},
      {"label": "Summarize",    "value": 280},
      {"label": "Chat",         "value": 210},
      {"label": "Search",       "value": 120}
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

## PBIP Workflow (when source is a .pbip file)

When a user has a `.pbip` Power BI project open in Power BI Desktop, this
path queries the **live in-memory model** directly — no screenshots needed.

```bash
python convert_dashboard.py "MyReport.pbip"
```

### What's different from the standard workflow

| | Standard (PDF/PPTX) | PBIP |
|---|---|---|
| Data source | Claude reads images | Claude queries live model via MCP |
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

Example MCP call pattern:
```
execute_query(database="...", dax="EVALUATE ROW(\"Total Revenue\", [Total Revenue])")
```

**Step 3: Drill deeper with custom DAX (optional)**

Modify pre-built queries to apply filters or explore dimensions:
```dax
-- Filter to last 30 days
EVALUATE
CALCULATETABLE(
    SUMMARIZECOLUMNS('Date'[Month], "Revenue", [Total Revenue]),
    DATESINPERIOD('Date'[Date], TODAY(), -30, DAY)
)
```

**Step 4: Use `measure_dax` to understand HOW each KPI is calculated**
- Measure uses `DATESYTD()` → note it's a year-to-date figure in insights
- Measure uses `DIVIDE()` with an `IF()` → zero-division guard; note denominator context
- Measure uses `CALCULATE()` with a filter → be explicit about what filter applies

**Step 4a: MANDATORY — Preserve visual type for Table and Matrix visuals**

> ❌ THE SINGLE MOST COMMON MISTAKE: Converting a Table or Matrix visual into a bar/column chart.

For every visual on each page, check the visual type in `pbip_context.json` → `pages[n].visuals[m].type`:
- If `type` is `"tableEx"`, `"pivotTable"`, `"matrix"`, or any table variant → ChartSpec **MUST** use `type: "table"`
- **Never** collapse a table into a bar or column chart, even if the data could be displayed that way
- Preserve every row and column exactly as the table shows them
- Use `table_columns` for column headers and `table_rows` for all data rows

| Power BI visual type | Required ChartSpec type |
|---|---|
| `tableEx` / `table` | `"table"` |
| `pivotTable` / `matrix` | `"table"` |
| `barChart` / `clusteredBarChart` | `"bar"` |
| `stackedBarChart` | `"bar_stacked"` |
| `hundredPercentStackedBarChart` | `"bar_stacked_100"` |
| `columnChart` / `clusteredColumnChart` | `"column"` |
| `stackedColumnChart` | `"column_stacked"` |
| `hundredPercentStackedColumnChart` | `"column_stacked_100"` |
| `lineChart` | `"line"` |
| `areaChart` / `stackedAreaChart` | `"area"` |
| `lineClusteredColumnComboChart` / `lineStackedColumnComboChart` | `"combo"` |
| `waterfallChart` | `"waterfall"` |
| `ribbonChart` | `"ribbon"` |
| `donutChart` | `"donut"` |
| `pieChart` | `"pie"` |
| `scatterChart` | `"scatter"` |
| `gauge` / `singleValueGauge` | `"gauge"` |
| `card` / `multiRowCard` | `"kpi"` or `"kpi_row"` |
| `treemap` | `"treemap"` |
| `funnel` | `"funnel"` |
| `filledMap` / `map` / `shapeMap` | `"heatmap"` (approximate) |
| `decompositionTreeVisual` | `"treemap"` (approximate) |

**Step 4b: MANDATORY — Cross-check every number: DAX result vs visual display**

Before writing any number in an insight or headline, execute its DAX query and compare to what the visual shows:

```
DAX result  →  89%
Visual shows → 89%   ✅ Use 89%

DAX result  →  89%
Visual shows → 87%   ✅ Use 89% (DAX is authoritative); note the discrepancy silently
```

- The DAX query result is always the source of truth
- If your planned insight says 87% but DAX returned 89%, use 89%
- Never state a number you have not verified with a DAX query (or visual, if DAX is unavailable)

**Step 4c: MANDATORY — Match metric labels to DAX formula AND page filter context**

Two sources define how a metric should be labelled. Check **both** before writing any label.

**Source A — The DAX formula** (what the measure mathematically computes):
- `DIVIDE([Sessions], [Users])` → "sessions per user" — **no time unit; stop here**
- `DIVIDE([Sessions], [Users]) / [ActiveWeeks]` → "sessions per user per week" (week divisor is explicit)
- `CALCULATE([M], DATESINPERIOD(..., -30, DAY))` → "over the last 30 days" (period baked into DAX)

**Source B — Active page/visual filters** (what scope the measure evaluates over):
- Check `pbip_context.json` → `pages[n].filters` and `pages[n].visuals[m].filters`
- A date slicer set to "Mar–Jun 2025" means the measure runs over those 4 months → append as scope context: "sessions per user (Mar–Jun 2025)"
- A slicer set to "Last 7 Days" → "sessions per user (last 7 days)"
- Page/visual filters define the **evaluation window**, not an additional mathematical divisor

**Combining both sources:**

| DAX formula | Active filter | Correct label |
|---|---|---|
| `DIVIDE([S],[U])` | Slicer: Mar–Jun 2025 | "sessions per user (Mar–Jun 2025)" |
| `DIVIDE([S],[U])` | Slicer: Last 7 days | "sessions per user (last 7 days)" |
| `DIVIDE([S],[U])` | No date filter | "sessions per user (selected period)" |
| `DIVIDE([S],[U]) / [Weeks]` | Any | "sessions per user per week" |

**Rule: Never add a time unit that appears in neither the DAX formula nor the active page/visual filters. Every part of a metric label must be traceable to one of these two sources.**

**Step 5: Generate insights from queried values — NOT from estimation**
- Every number in insights must come from an executed DAX query result
- Follow the same insight formula as the standard workflow
- Reference measure names in parentheses: "47.3% completion rate ([Project Completion Rate])"

**Step 6: Write insights to `temp/insights.json`** (same format as always)

### If DAX queries return no data

- Power BI Desktop may not be open, or the MCP may not be connected
- Fall back to any companion images if available (`temp/slide_N.png`)
- If no images exist either, mark slide as "Insufficient data for analysis"

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

## Manual Workflow (Optional)

If users want step-by-step control:

```bash
# Step 1: Prepare only
python convert_dashboard.py "dashboard.pptx" --prepare

# Step 2: You analyze (same as above)

# Step 3: Build only
python convert_dashboard.py --build --output "executive.pptx"
```

---

## Insight Generation Guidelines

**CRITICAL: Before analyzing any dashboard, read `DASHBOARD_READING_RULES.md` for data extraction accuracy rules.**

When generating insights (Step 2), follow these principles:

### ✅ DO: Be Concise and Friendly

**Good Example:**
- Headline: "134 Agent users from 1,275 total (11%) - significant opportunity to expand automation adoption"
- Insight: "HR Generalists (140 actions/user) are 3-4x above average - great candidates to champion agent adoption"

**Why it works:** Specific numbers, identifies opportunity, suggests action - all in one sentence

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
- Dashboard shows "13.5K" → Use "13.5K actions" (NOT "13K")

**Test each number:**
1. Can I point to this exact number on the dashboard?
2. Does the unit (K, M, %, decimal) match exactly?
3. If NO to either → Do NOT use that number

**Why this matters:** Unit mismatches (13 vs 13K) destroy credibility instantly

### ✅ DO: Answer "So What?"

Every insight should help IT decision maker take action:
- ✅ "Client Finance's 217 prompts/user pattern ready to replicate across Corporate Finance"
- ❌ "Client Finance has 217 prompts per user"

### ✅ DO: Analyze Platform/Feature Patterns

**CRITICAL: Only mention entities (platforms, teams, features) that are VISIBLE on the dashboard page.**

Always look for and call out platform, app, or feature variations **IF they appear on the visual:**
- ✅ "Teams integration shows 3x higher engagement than Outlook" — **ONLY if both "Teams" and "Outlook" are visible on this page**
- ✅ "PowerPoint Copilot leads with 450 actions while Excel shows only 89" — **ONLY if both apps are shown on this page**
- ✅ "Chat (Web) dominates with 1,275 users" — **ONLY if "Chat (Web)" label is visible on this page**
- ❌ "Users engage with the platform" (missing which platform/feature)
- ❌ Mentioning "Teams" when it's not shown on this specific page

**Rule: Can you point to where this name appears on the dashboard?**
- **YES** → Safe to mention
- **NO** → Do NOT mention, even if you think it's on other pages

**What to look for (only if visible on this page):**
- Microsoft 365 apps: Outlook, Teams, Word, Excel, PowerPoint, OneNote
- Copilot features: Chat, Agents, M365 Copilot, Business Chat
- Integration patterns: Web vs. integrated, standalone vs. embedded
- Feature adoption: Which capabilities drive engagement
- Platform concentration: Where is usage concentrated
- Department names, team names, location names

**Why this matters:**
- Reveals which workflows deliver value (content creation vs. analysis vs. communication)
- Identifies integration successes and failures
- Shows where to focus training and enablement
- Guides feature prioritization decisions

**Why the visibility rule matters:**
- Every claim must be verifiable from the visual
- Users will check if "Teams" actually appears on that page
- One unverifiable claim destroys credibility

### ❌ DON'T: Be Critical or Verbose

Avoid:
- Critical language: "critical failure", "deployment disaster", "massive gap"
- Extreme emotional language: "crisis", "catastrophically", "blocking" - use "opportunity", "challenge", "limiting" instead
- Verbose explanations: Keep insights to 1-2 sentences max
- Criticizing the analytics: Don't say "missing data" or "blind spot"
- Technical jargon: Use business language

**Tone guidelines for executives:**
- ✅ "habit formation opportunity" ❌ "habit formation crisis"
- ✅ "significant cleanup opportunity" ❌ "massive cleanup required"
- ✅ "opportunity to unlock features" ❌ "deployment gap blocking features"
- ✅ "concentration in infrequent tier" ❌ "catastrophically low usage"

**Why:** Extreme language triggers defensive reactions. Data-driven opportunities encourage action.

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

**CRITICAL: Credibility over completeness**

If a dashboard page has no numbers or is blank:
- ✅ **DO:** Mark as "Insufficient data for analysis"
- ✅ **DO:** Skip the slide entirely
- ✅ **DO:** Explain what data is needed
- ❌ **DON'T:** Generate vague insights to fill space
- ❌ **DON'T:** Make assumptions about missing data

**Why this matters:**
- User trust depends on accuracy and honesty
- One unsupported insight undermines all insights
- Better to deliver 8 strong insights than 10 with 2 weak ones

**Example when data is missing:**
```json
{
  "slide_number": 7,
  "title": "Agents - Habit Formation",
  "headline": "Insufficient data for analysis",
  "insights": [
    "Dashboard page contains no extractable usage frequency metrics",
    "To generate habit formation insights, need: daily/weekly/monthly usage distribution, engagement tiers, or frequency patterns",
    "Recommend adding frequency tracking to dashboard for next analysis cycle"
  ],
  "numbers_used": []
}
```

---

## Insight Formula

Follow this proven formula for each slide:

**Headline:** `[Clear Takeaway Message]` - Focus on instant clarity without requiring numbers

**IMPORTANT:** Headlines should answer "so what?" and be memorable. Numbers are NOT required in headlines - they belong in the supporting insights where they provide evidence.

Good headline examples:
- ✅ "Multi-platform adoption demonstrates organization-wide AI readiness"
- ✅ "Visual Creator emerges as winning agent template for scaling"
- ✅ "Habit formation gap prevents agents from becoming daily workflows"
- ✅ "Operations establishes benchmark for M365 Copilot excellence"

Bad headline examples:
- ❌ "4,381 total active users across Agent (180), Unlicensed Chat (2,585)..." (too many numbers, unclear takeaway)
- ❌ "101 of 116 total agent users stay in infrequent tier with zero daily users" (data dump, not insight)

**Insight format: two-part, separated by `||`**

Each insight must follow this exact structure:
`"[Bold punchy line, 6-8 words max] || [Supporting evidence with specific data]"`

The bold line is rendered large and bold on the slide. The detail is rendered below in normal weight. Keep them clearly separated — the bold line is the "so what", the detail is the proof.

**Insight 1 (Scale/Scope):**

Example: `"HR Generalists lead usage by far || 140 actions/user is 3-4x the average — strong candidates to champion adoption"`

**Insight 2 (Pattern/Opportunity):**

Example: `"Client Finance pattern ready to replicate || 217 prompts/user sets the benchmark for Corporate Finance expansion"`

**Insight 3 (Action):**

Example: `"Pilot training with top 3 departments || Target 50+ prompts/user baseline to establish org-wide standard"`

**Remember:** The headline is what an executive will remember 3 days later. Make it clear, memorable, and clearly supported by the numbered evidence in your insights.

---

## Technical Requirements

### Image Analysis
- Dashboard images are 2560x1460 pixels (Power BI export default)
- Look for: KPI cards, trends charts, tables, leaderboards
- Extract: Numbers, percentages, trends, comparisons, distributions
- **CRITICAL: Identify platforms, apps, and features** (Outlook, Teams, PowerPoint, Excel, Agents, Chat, etc.)
- Notice engagement variations across different platforms/features
- Use platform patterns to identify high-value workflows

### Data Extraction Accuracy Rules (CRITICAL)

**Read `DASHBOARD_READING_RULES.md` for complete guidelines. Key rules:**

1. **Match numbers to EXACT labels** - Read table labels carefully; "Team C: 73,071" not "Team B: 73,071"
2. **Never assume first row is aggregate** - Read actual label; could be "Team B" not "Total"
3. **Read chart axes precisely** - Verify scale/units; don't guess values
4. **Identify selected filters** - Black/highlighted = active; chart shows only selected filter data
5. **Match units exactly** - "13K" shown → write "13K" (not "13,000")
6. **User categories exact** - "148 Unlicensed" not "148 Bottom 25%" if label differs
7. **Only cite visible numbers** - Can you point to it? If no, don't use it

**Pre-analysis checklist for each page:**
- [ ] Identify filter selections (black vs grey elements)
- [ ] Read all table row/column labels
- [ ] Check chart axis scales and units
- [ ] Read legend for color meanings
- [ ] Verify every number matches its label

### Slide Types to Recognize
- **Trends:** Time-series charts → Focus on momentum, growth, patterns
- **Leaderboards:** Rankings/tables → Focus on top performers, concentration, gaps
- **Health Check:** KPI dashboards → Focus on portfolio health, key metrics
- **Habit Formation:** Frequency distributions → Focus on engagement tiers, usage patterns
- **License Priority:** User segments → Focus on upgrade candidates, ROI opportunities
- **Platform/Feature Comparison:** Usage across apps → Focus on which platforms drive value, integration patterns

### Output Format
- 16:9 widescreen (13.333" x 7.5")
- Image on left (6.5" max width)
- Insights on right (5.5" width)
- 3 bullet points per slide maximum
- Segoe UI font, blue accent colors

---

## Quality Validation

After generating insights, perform these mandatory checks:

### Checklist

✅ **Every headline has specific number** (not "some users" or "many")
✅ **Number units match exactly** (13 vs 13K vs 13M - use what's shown)
✅ **All entities are visible** (platforms, teams, features mentioned are on this page)
✅ **Insights are concise** (1-2 sentences each, not paragraphs)
✅ **Tone is friendly** (opportunities, not failures)
✅ **Focus on action** (what to do, not just what is)
✅ **Numbers are traceable** (can point to exact location on dashboard)
✅ **No criticism** (don't critique the analytics or report)
✅ **No forced insights** (if no data, mark "Insufficient data" rather than generating generic content)
✅ **No vanilla statements** (every headline passes the "would a VP forward this?" test)
✅ **Platform patterns identified** (mention specific apps/features ONLY when visible on this page)
✅ **Executive summary synthesizes ALL pages** (not just first slide)
✅ **Recommendations are actionable** (specific actions, not vague suggestions)
✅ **Both sections grounded in data** (can trace each point to dashboard numbers)

### Chart Coverage Self-Check (MANDATORY)

Before writing `temp/insights.json`, verify that **every slide with quantitative data has at least one chart spec**. Slides without chart specs fall back to pasting the raw screenshot — which defeats the purpose of the executive deck.

**For each slide ask:** Does this page show numbers, bars, columns, lines, donuts, scatter plots, treemaps, waterfall charts, combo charts, tables, or KPIs?
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

## Success Criteria

When the user says "Convert X to executive deck":

✅ Output file created at expected path (16:9 widescreen)
✅ All source slides have corresponding output slides (1:1 mapping)
✅ Dashboard images embedded on left side of each slide
✅ Headlines are insight-driven with specific numbers
✅ Insights are concise (1-2 sentences), friendly, actionable
✅ No generic statements like "there are X users"
✅ Visual formatting follows Analytics template (blue, professional)
✅ Constitution rules followed (Section 4-6)
✅ Images placed without overlap with text
✅ 2-3 insights per slide (executive-focused)
✅ **Execution time < 30 seconds total**

---

## Example Workflow in Practice

**User says:** "Convert wpp22.pptx to executive deck"

**You do:**

1. Run prepare command:
   ```bash
   python convert_dashboard.py wpp22.pptx --prepare
   ```

2. Read analysis request and view images:
   ```
   Read temp/analysis_request.json
   Read temp/slide_2.png
   Read temp/slide_3.png
   ... (view all slides)
   ```

3. Generate concise, friendly insights for each slide following formula

4. Save insights JSON to temp/insights.json

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
4. **Re-run failed step** - Prepare, analyze, or build independently
5. **Report clearly** - State which step failed and why

---

## Key Advantages of This Approach

✅ **No API key needed** - Uses current Claude Code session
✅ **High-quality insights** - Real intelligence, not rule-based patterns
✅ **Fast execution** - < 30 seconds total
✅ **Scalable** - Works for any user without configuration
✅ **Maintainable** - Separate deterministic and intelligent layers
✅ **Professional output** - 16:9 widescreen, Analytics template styling
✅ **Constitution-compliant** - Automated validation built-in
