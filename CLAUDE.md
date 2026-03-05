# Claude Project Instructions

## Project Overview

Converts Power BI dashboard exports (.pptx, .pdf, .pbip) into executive-ready 16:9 PPTX presentations using Claude's vision and analytical capabilities. No API key needed — uses current Claude Code session.

## Workflow

```bash
python convert_dashboard.py "dashboard.pptx"             # Full run → dashboard_executive.pptx
python convert_dashboard.py "file.pptx" --prepare        # Step 1 only (extract)
python convert_dashboard.py --build --output "exec.pptx" # Step 3 only (build)
python convert_dashboard.py --verify                     # Validate insights.json
```

**Three steps:**
1. **Extract (auto, ~5s):** Images → `temp/slide_N.png`, metadata → `temp/analysis_request.json`. PPTX skips first slide (title/metadata); PDF includes all pages.
2. **Analyze (you):** Read images → generate insights → write `temp/insights.json`.
3. **Build (auto, ~3s):** Renders slides, validates against Constitution.

**Rendering modes:**

| Mode | Default | Flag | Output |
|---|---|---|---|
| Screenshot | ✅ | — | PBI screenshot embedded + insight commentary |
| Vector charts | — | `--vector-charts` | matplotlib charts from chart specs in insights.json |

Use `--vector-charts` when: exact DAX data available, screenshots are low quality, or specific data points need highlighting. Keep default when: PBI visuals are clean, or complex chart types (dual-axis, heatmaps) are hard to reconstruct.

---

## Analysis Task

1. `Read temp/analysis_request.json`
2. `Read temp/slide_N.png` for each slide listed
3. Act as senior analyst advisor to an IT decision maker
4. Write `temp/insights.json`

**CRITICAL: Read `DASHBOARD_READING_RULES.md` before analyzing.**

### insights.json Root Structure

```json
{
  "deck_title": "Compelling title: specific finding or story",
  "deck_subtitle": "Platform · Org · Month Year – Month Year",
  "executive_summary": [
    "Finding with specific number → business implication"
  ],
  "recommendations": ["Action verb + what to do + expected outcome"],
  "slides": [ /* per-slide objects */ ]
}
```

**deck_title** — 5–10 words, compelling, positively framed, specific to content.
- ✅ "Copilot Impact Confirmed: $14.7M in Value with 84% Active Adoption"
- ✅ "From Reach to Routine: Building AI Habits Across the Organization"
- ❌ "Executive Insights" | "Copilot Usage Report" | any negative framing

**deck_subtitle** — `"[Platform 1] · [Platform 2] · [Org] · [Date range]"` — always populate, never leave blank.

**executive_summary** — 5 bullets synthesizing ALL pages, highest business impact first, each with specific numbers, format: `"[finding] → [implication]"`.

**recommendations** — 3–5 specific actions traceable to dashboard data. "Pilot agent training with HR Generalists (140 actions/user) to establish best practices" ✅ — "Improve training" ❌.

### Per-Slide Structure

```json
{
  "slide_number": 1,
  "title": "Slide Title",
  "headline": "Memorable insight answering 'so what?' — not a data dump",
  "insights": [
    {"text": "Bold line 6-8 words || Supporting evidence with specific data", "chart": { /* spec or null */ }},
    {"text": "Second insight || Detail", "chart": null},
    {"text": "Third insight || Detail", "chart": null}
  ],
  "numbers_used": ["134", "1,275", "11%"]
}
```

**Insight text format:** `"[Bold line, 6-8 words max] || [Supporting evidence with data]"`
The `||` splits bold heading (rendered large) from normal-weight detail. Max 3 insights per slide.
- Insight 1 (Scale/Scope): `"HR Generalists lead usage by far || 140 actions/user is 3-4x the average — strong candidates to champion adoption"`
- Insight 2 (Pattern/Opportunity): `"Client Finance pattern ready to replicate || 217 prompts/user sets the benchmark for Corporate Finance expansion"`
- Insight 3 (Action): `"Pilot training with top 3 departments || Target 50+ prompts/user baseline to establish org-wide standard"`

**Missing data:** If a page has no numbers → set headline to "Insufficient data for analysis", list what data is needed, set `"numbers_used": []`. Do NOT generate vague insights to fill space.

---

## Chart Specs

In **screenshot mode** (default): set `"chart": null` on all insights.
In **vector chart mode** (`--vector-charts`): include chart specs. **MANDATORY — every slide with quantitative data needs at least one chart spec.** Common misses: KPI cards encoded as text, tables summarized instead of `"table"` spec, small charts overlooked.

### Chart Types

| Visual | type |
|---|---|
| Horizontal bars | `"bar"` |
| Stacked horizontal | `"bar_stacked"` / `"bar_stacked_100"` |
| Vertical columns | `"column"` |
| Stacked vertical | `"column_stacked"` / `"column_stacked_100"` |
| Line trend | `"line"` |
| Area filled | `"area"` |
| Columns + line (dual-axis) | `"combo"` |
| Waterfall | `"waterfall"` |
| Pie / Donut | `"pie"` / `"donut"` |
| KPI single | `"kpi"` |
| KPI row | `"kpi_row"` |
| Table / Matrix | `"table"` — **NEVER convert to bar/column** |
| Heatmap | `"heatmap"` |
| Treemap | `"treemap"` |
| Scatter | `"scatter"` |
| Gauge | `"gauge"` |
| Funnel | `"funnel"` |
| Ribbon | `"ribbon"` |

### Fidelity Tiers

| Tier | Types | Action |
|---|---|---|
| **High** — always reconstruct | `kpi`, `kpi_row`, `bar`, `column`, `donut`, `pie`, `table` | Always include chart spec |
| **Medium** — reconstruct with care | `line`, `area`, `column_stacked`, `bar_stacked`, `treemap`, `funnel` | Include; verify series shape |
| **Low** — prefer screenshot | `scatter`, `combo` (dual-axis), `heatmap`, `ribbon`, multi-axis | Use `"chart": null` unless exact DAX data |

### Dual-Axis Charts

Use `"combo"` for visuals with two Y-axes. Never plot series with >3× scale difference on a single axis. If original uses area+line or line+line dual-axis → screenshot fallback.

`data` = bar series (left Y-axis), `series` = line overlay (right Y-axis), `y_label` = left axis, `x_label` = right axis.

### Chart Spec Examples

```json
// bar / column
{"type": "bar", "title": "Weekly Actions", "data": [{"label": "Dana Bourque", "value": 48.95}, {"label": "Matt Sheard", "value": 38.09}]}

// kpi
{"type": "kpi", "value": "5.35", "label": "Hours/Week/Person"}

// kpi_row
{"type": "kpi_row", "items": [{"value": "256", "label": "Active Users"}, {"value": "33.77", "label": "Weekly Actions"}, {"value": "52.6%", "label": "Power Users"}]}

// table
{"type": "table", "columns": ["Manager", "Users", "Weekly Actions"], "rows": [["Dana Bourque", "4", "48.95"], ["Matt Sheard", "52", "38.09"]]}

// line (multi-series)
{"type": "line", "title": "Monthly Active Users", "series": [{"name": "Agents", "points": [{"x": "Jan", "y": 120}, {"x": "Feb", "y": 145}]}, {"name": "Chat", "points": [{"x": "Jan", "y": 80}, {"x": "Feb", "y": 95}]}]}

// donut
{"type": "donut", "title": "Usage Distribution", "data": [{"label": "Power Users", "value": 52.6}, {"label": "Regular", "value": 31.2}, {"label": "Light", "value": 16.2}]}

// waterfall (positive=green, negative=red; explicit "color" or "total"/"net" labels anchor at zero)
{"type": "waterfall", "title": "Revenue Bridge", "data": [{"label": "Q1 Revenue", "value": 100, "color": "#003278"}, {"label": "New Sales", "value": 30}, {"label": "Churn", "value": -12}, {"label": "Q2 Total", "value": 118, "color": "#003278"}]}

// combo (dual-axis)
{"type": "combo", "title": "Users vs Sessions/User", "y_label": "Active Users", "x_label": "Sessions/User",
 "data": [{"label": "Mar", "value": 142}, {"label": "Apr", "value": 164}],
 "series": [{"name": "Sessions/User", "points": [{"x": "Mar", "y": 2.6}, {"x": "Apr", "y": 3.9}]}]}

// scatter
{"type": "scatter", "title": "Usage vs Satisfaction", "x_label": "Weekly Actions", "y_label": "Score", "series": [{"name": "Team A", "x": 45, "y": 8.2}, {"name": "Team B", "x": 32, "y": 7.5, "highlight": true}]}

// treemap
{"type": "treemap", "title": "Feature Usage Share", "data": [{"label": "Email Drafts", "value": 340}, {"label": "Summarize", "value": 280}]}
```

---

## Critical Data Rules

### Metric Segment Isolation (CRITICAL)

**Never mix numbers from different platform segments (Licensed / Unlicensed / Agent).**

OCR and `text_metrics` flatten all segments into one stream, causing cross-contamination. Before assigning any number to a chart series, verify:
1. Which visual does it belong to? (`pbip_context.json` → `measures[]`)
2. Which page region? Left panel numbers ≠ right panel chart.
3. Does the measure name match the series label? "Licensed" measure → Licensed charts only.

**Example:** Page shows `180 Agent Users | 3.60 Agent Sessions/User | 2,585 Chat Users | 3.3 Chat Sessions/User`. OCR stream flattens to `180 · 3.60 · 2,585 · 3.3`. Read the label below each number to confirm `3.60` pairs with "Agent Sessions Per User", not Chat.

### Units and Labels

- **Match units exactly:** "13K" shown → write "13K", not "13,000" or "13".
- **Only cite visible entities:** Can you point to the name on this exact page? YES → cite it. NO → don't.

**Pre-analysis checklist (every page):**
- [ ] Filter selections identified (black/highlighted = active; chart shows only selected filter data)
- [ ] All table row/column labels read
- [ ] Chart axis scales and units checked
- [ ] Legend colors verified
- [ ] Every number matched to its label

### OCR Cross-Verification (when `ocr_used: true`)

When `analysis_request.json` marks a slide `"ocr_used": true`, always verify number-to-label pairing against the actual image. Do NOT trust `text_metrics[].context` — OCR flattens spatial groupings. Example: seeing `151` near "AI Skills" does NOT mean "151 people have AI Skills" — verify which label is above/below the number in the image.

---

## PBIP Workflow

When `temp/pbip_context.json` exists, query the live model via MCP instead of estimating from screenshots.

1. `Read temp/pbip_context.json` — contains pages, visuals, measures (with full DAX), dax_queries
2. Execute DAX via `powerbi-modeling` MCP for each page's queries: `execute_query(database="...", dax="EVALUATE ...")`
3. All numbers MUST come from DAX results, not visual estimation.
4. If DAX returns no data → fall back to `temp/slide_N.png` → else mark "Insufficient data for analysis".

**MANDATORY — Preserve table visuals as tables:**
`tableEx`, `pivotTable`, `matrix` → ChartSpec type = `"table"`. **NEVER collapse to bar/column**, even if the data could be displayed that way.

**PBI visual type → ChartSpec mapping:**

| PBI type | ChartSpec type |
|---|---|
| `tableEx` / `table` / `pivotTable` / `matrix` | `"table"` |
| `barChart` / `clusteredBarChart` | `"bar"` |
| `stackedBarChart` / `hundredPercentStackedBarChart` | `"bar_stacked"` / `"bar_stacked_100"` |
| `columnChart` / `clusteredColumnChart` | `"column"` |
| `stackedColumnChart` / `hundredPercentStackedColumnChart` | `"column_stacked"` / `"column_stacked_100"` |
| `lineChart` | `"line"` |
| `areaChart` / `stackedAreaChart` | `"area"` |
| `lineClusteredColumnComboChart` / `lineStackedColumnComboChart` | `"combo"` |
| `donutChart` / `pieChart` | `"donut"` / `"pie"` |
| `card` / `multiRowCard` | `"kpi"` / `"kpi_row"` |
| `waterfallChart` | `"waterfall"` |
| `scatterChart` | `"scatter"` |
| `treemap` | `"treemap"` |
| `filledMap` / `map` / `shapeMap` | `"heatmap"` |
| `decompositionTreeVisual` | `"treemap"` |

**MANDATORY — Cross-check every number:** DAX result is authoritative. If DAX returns 89% and visual shows 87%, use 89%.

**MANDATORY — Metric label = DAX formula + active page filter:**

| DAX formula | Active filter | Correct label |
|---|---|---|
| `DIVIDE([S],[U])` | Slicer: Mar–Jun 2025 | "sessions per user (Mar–Jun 2025)" |
| `DIVIDE([S],[U])` | No date filter | "sessions per user (selected period)" |
| `DIVIDE([S],[U]) / [Weeks]` | Any | "sessions per user per week" |

Rule: Never add a time unit that appears in neither the DAX formula nor the active page/visual filters.

**Use `measure_dax` to understand how each KPI is calculated:** `DATESYTD()` → year-to-date figure; `DIVIDE()` with `IF()` → zero-division guard; `CALCULATE()` with filter → note what filter applies.

---

## Insight Quality Rules

### Headline
- Clear, memorable "so what?" — not a data dump of numbers
- Numbers NOT required in headline (they belong in supporting insights as evidence)
- ✅ "Multi-platform adoption demonstrates organization-wide AI readiness"
- ✅ "Visual Creator emerges as winning agent template for scaling"
- ❌ "4,381 total active users across Agent (180), Unlicensed Chat (2,585)..." (data dump)

### DO
- **Opportunity framing:** "opportunity to expand from 11%" ✅ — "only 11% shows deployment failure" ❌
- **Exact units:** "217 prompts/user (5.7× average)" ✅ — "some users show higher engagement" ❌
- **Answer "So What?":** Every insight implies an action. "Client Finance pattern ready to replicate across Corporate Finance" ✅ — "Client Finance has 217 prompts/user" ❌
- **Visibility rule:** Only mention platforms/teams/features visible on this specific page. "Teams integration shows 3× engagement" only if "Teams" label appears on this page.
- **VP forward test:** "Would a VP forward this bullet to their boss?" If no → rewrite.

### DON'T
- Critical/extreme language: "crisis", "catastrophically", "blocking" → use "opportunity", "challenge", "limiting"
- Mention entities not visible on this specific page
- Generic statements that could apply to any dashboard
- Numbers you can't point to on the dashboard
- Force insights on blank pages → mark "Insufficient data for analysis"

### Anti-vanilla transformations
- ❌ "There are 1,275 active users" → ✅ "1,275 active users = 68% penetration — 480-seat expansion opportunity"
- ❌ "Usage varies by department" → ✅ "Operations outpaces Marketing 3:1 on weekly actions — replicate their onboarding playbook"
- ❌ "Skills are distributed across categories" → ✅ "3 of 8 skill categories account for 80% of confirmations — focus L&D on the long tail"

---

## Quality Checklist (run before saving insights.json)

✅ Units match exactly (13 vs 13K vs 13M — use what's shown)
✅ All entities mentioned are visible on that specific page
✅ Insights concise (1-2 sentences each, not paragraphs)
✅ Opportunity framing (not failure/crisis language)
✅ Every number traceable to dashboard (can point to it)
✅ No vanilla statements (pass VP-forward test for every headline)
✅ Every slide with quantitative data has ≥1 chart spec
✅ `deck_title` populated and compelling (never "Executive Insights")
✅ `deck_subtitle` populated with scope and date range
✅ Executive summary covers ALL pages, highest impact first, with specific numbers
✅ Recommendations are specific actions traceable to dashboard numbers
✅ No forced insights on empty/blank pages

---

## Slide Types to Recognize

| Slide type | Focus |
|---|---|
| Trends | Momentum, growth, patterns in time-series |
| Leaderboards | Top performers, concentration, gaps |
| Health Check | Portfolio health, KPI status |
| Habit Formation | Engagement tiers, usage frequency distribution |
| License Priority | Upgrade candidates, ROI opportunities |
| Platform/Feature Comparison | Which platforms/apps drive value, integration patterns |

---

## Output Specs

- 16:9 widescreen (13.333" × 7.5"), Segoe UI font, blue accent colors
- Dashboard image: left side (6.5" max width)
- Insights panel: right side (5.5" width)
- 3 bullet points per slide maximum

---

## Files Reference

| File | Purpose |
|---|---|
| `convert_dashboard.py` | Main orchestrator |
| `DASHBOARD_READING_RULES.md` | **Read before any analysis** |
| `Claude PowerPoint Constitution.md` | Quality standards and governance |
| `lib/rendering/builder.py` | Slide renderer (16:9 format) |
| `lib/rendering/validator.py` | Constitution compliance checker |
| `temp/` | Working directory (auto-generated, not committed) |
