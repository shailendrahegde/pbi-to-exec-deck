# Dashboards → Decks

### From raw Power BI data to executive-ready presentations — in minutes.

Turn Power BI dashboards into polished, insight-driven presentations automatically. No design skills, no manual slide-building, no copy-pasting numbers.

Two modes depending on how much access you have to the source data:

| | Quick Mode | Deep Analysis Mode |
|---|---|---|
| **Input** | PDF or PPTX export | `.pbix` or `.pbip` file |
| **Data source** | Dashboard screenshots | Live Power BI model via MCP |
| **Insight quality** | Visual analysis | Exact DAX-verified numbers |
| **Setup** | Claude Code only | Claude Code + MCP server |
| **Time** | ~3 minutes | ~5–10 minutes |

---

## Mode 1 — Quick Mode (PDF or PPTX)

Export your dashboard from Power BI and feed it directly. Claude reads each page as an image and generates analyst-grade insights.

### Prerequisites

- [Claude Code CLI](https://docs.anthropic.com/en/docs/claude-code/overview) — the only requirement. Python and dependencies are auto-installed on first run.

### Step 1: Export your dashboard from Power BI

**PDF** — Power BI Desktop: `File → Export → Export to PDF`

**PPTX** — Power BI Service: `File → Export → PowerPoint`

Both formats work identically.

### Step 2: Clone the repo and open Claude Code

```bash
git clone https://github.com/shailendrahegde/pbi-to-exec-deck.git
cd pbi-to-exec-deck
claude
```

### Step 3: Run the conversion

```
convert to an executive deck "C:\path\to\your\dashboard.pdf"
```

Or with a PPTX export:

```
convert to an executive deck "C:\path\to\your\dashboard.pptx"
```

### Step 4: Done

Your presentation is saved as `dashboard_executive.pptx` in the project folder.

**What Claude does automatically:**
- Extracts each dashboard page as an image
- Reads KPIs, trends, tables, and comparisons as a senior analyst
- Generates insight-driven headlines with specific numbers
- Builds an executive summary and action recommendations
- Renders clean 16:9 slides with SVG charts

---

## Mode 2 — Deep Analysis Mode (PBIX or PBIP)

Feed the Power BI source file directly. Claude connects to the live data model via MCP and queries exact values using DAX — no visual estimation.

**Why this mode is more powerful:**
- Every number comes from a DAX query, not a screenshot read
- Can drill into any dimension, filter, or time period
- Measure logic is transparent (Claude reads the DAX expressions)
- Handles reports where screenshots don't capture the full picture

### Prerequisites

**1. Claude Code CLI**

Install from [claude.ai/code](https://claude.ai/code) or:
```bash
npm install -g @anthropic-ai/claude-code
```

**2. Power BI Desktop**

Download from the [Microsoft Store](https://apps.microsoft.com/store/detail/power-bi-desktop/9NTXR16HNW1T) or the Power BI download page. Must be installed on Windows.

**3. Power BI Modeling MCP Server**

This MCP server lets Claude query your live Power BI Desktop model via DAX. Install it by running the included setup script:

```bash
python setup_pbi_mcp.py
```

The script will:
- Detect if the MCP server is already installed (via VS Code extension)
- Download and configure it automatically if not found
- Register it in `.mcp.json` so Claude Code picks it up

To verify the setup:
```bash
python setup_pbi_mcp.py --check
```

> **Note:** The MCP server is sourced from [microsoft/powerbi-modeling-mcp](https://github.com/microsoft/powerbi-modeling-mcp). It runs locally and communicates with Power BI Desktop on your machine — no data leaves your environment.

### Step 1: Open your report in Power BI Desktop

Power BI Desktop must be **running with the report open** before you run the conversion. The MCP server connects to the in-memory model.

### Step 2: Clone the repo and open Claude Code

```bash
git clone https://github.com/shailendrahegde/pbi-to-exec-deck.git
cd pbi-to-exec-deck
claude
```

### Step 3: Run the conversion

With a `.pbix` file:
```
convert to an executive deck "C:\path\to\your\report.pbix"
```

With a `.pbip` project folder:
```
convert to an executive deck "C:\path\to\your\report.pbip"
```

### Step 4: Done

Your presentation is saved as `report_executive.pptx`.

**What Claude does automatically:**
- Extracts report page structure and visual definitions
- Reads every measure's DAX expression to understand how KPIs are calculated
- Executes DAX queries against the live model for exact values
- Cross-checks numbers between query results and on-screen visuals
- Generates insights grounded in verified data (no estimation)
- Renders clean 16:9 slides with SVG charts matching the original visual types

---

## Output

Both modes produce the same presentation format:

- **16:9 widescreen** slides (1920×1080)
- **SVG charts** rendered from extracted data — bar, line, donut, heatmap, treemap, scatter, KPI cards, tables, and more
- **Insight-driven headlines** that answer "so what?" for each dashboard page
- **Executive summary** — 5 synthesized findings across all pages
- **Action recommendations** — specific, data-grounded next steps
- **Analytics template styling** — clean, professional, blue accent

---

## Example

**Raw dashboard numbers:**
> "1,275 active Copilot users", "134 Agent users"

**Turned into analyst-grade insight:**
> "134 Agent users from 1,275 total (11%) — significant opportunity to expand automation adoption. HR Generalists at 140 actions/user are 3–4x above average: strong candidates to champion agent adoption org-wide."

---

## What Makes This Different

| | This tool | Manual approach |
|---|---|---|
| Insight quality | Analyst-grade, data-grounded | Varies by author |
| Time | 3–10 minutes | Hours |
| Numbers | Exact (DAX) or verified (visual) | Copy-paste error risk |
| Charts | SVG rendered from data | Screenshot embeds |
| Consistency | Templated, constitutionally validated | Inconsistent |

---

## Key Files

| File | Purpose |
|---|---|
| `convert_dashboard_claude.py` | Main entry point for all conversions |
| `setup_pbi_mcp.py` | One-time MCP server setup for PBIX/PBIP mode |
| `lib/extraction/pbix_extractor.py` | PBIX ZIP extraction and static screenshot handling |
| `lib/extraction/pbip_extractor.py` | PBIP folder parsing and DAX query generation |
| `lib/rendering/chart_builder_mpl.py` | SVG chart rendering (matplotlib) |
| `lib/rendering/builder.py` | Slide layout and PPTX assembly |
| `CLAUDE.md` | Full instructions for Claude's analysis workflow |

---

## License

MIT — use freely within your organization.

---

**Found this useful?** [Star the repo](https://github.com/shailendrahegde/pbi-to-exec-deck) to help others find it.
