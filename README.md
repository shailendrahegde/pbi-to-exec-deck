# Dashboards → Decks

### From raw Power BI data to executive-ready presentations — in minutes.

Turn Power BI dashboards into polished, insight-driven presentations automatically. No design skills, no manual slide-building, no copy-pasting numbers.

---

## Prerequisite

**[Claude Code CLI](https://docs.anthropic.com/en/docs/claude-code/overview)** — required for both modes. Python and all other dependencies are auto-installed on first run.

---

## Mode 1 — Quick Mode

**Input:** PDF or PPTX export from Power BI
**How it works:** Claude reads each dashboard page as an image and generates analyst-grade insights

```
convert to an executive deck "C:\path\to\dashboard.pdf"
```

**Get started:**

```bash
git clone https://github.com/shailendrahegde/pbi-to-exec-deck.git
cd pbi-to-exec-deck
claude
```

Then paste the command above with your file path. Output is saved as `dashboard_executive.pptx` (~3 minutes).

**How to export from Power BI:**
- **PDF** — Power BI Desktop: `File → Export → Export to PDF`
- **PPTX** — Power BI Service: `File → Export → PowerPoint`

---

## Mode 2 — Deep Analysis Mode

**Input:** `.pbix` or `.pbip` source file
**How it works:** Claude connects to the live Power BI Desktop model via MCP and queries exact values using DAX

```
convert to an executive deck "C:\path\to\report.pbix"
```

```
convert to an executive deck "C:\path\to\report.pbip"
```

**Extra prerequisite:** Install the Power BI Modeling MCP server (one-time setup)

```bash
python setup_pbi_mcp.py
```

This downloads the official [microsoft/powerbi-modeling-mcp](https://github.com/microsoft/powerbi-modeling-mcp) server, extracts it to `C:\MCPServers\PowerBIModelingMCP\`, and registers it in `.mcp.json`. To verify:

```bash
python setup_pbi_mcp.py --check
```

**Before running:** Open your report in Power BI Desktop — the MCP connects to the running Desktop process to query the live model.

**Without MCP installed:** The tool still runs but falls back to image-only analysis (same as Mode 1). You'll see a clear message indicating this.

> The MCP server runs locally and communicates only with Power BI Desktop on your machine — no data leaves your environment.

---

## What you get

Both modes produce the same output format:

- **16:9 widescreen** slides
- **SVG charts** matching the original visual types — bar, line, donut, heatmap, treemap, scatter, KPI cards, tables
- **Insight-driven headlines** that answer "so what?" for each dashboard page
- **Executive summary** — 5 synthesized findings across all pages
- **Action recommendations** — specific, data-grounded next steps

| | Mode 1 (PDF/PPTX) | Mode 2 (PBIX/PBIP) |
|---|---|---|
| Data source | Screenshots | Live DAX queries |
| Numbers | Visually read | Exact from model |
| Measure logic | Unknown | Full DAX expressions |
| Extra setup | None | MCP server |

---

## Example

**Raw dashboard numbers:**
> "1,275 active Copilot users", "134 Agent users"

**Turned into analyst-grade insight:**
> "134 Agent users from 1,275 total (11%) — significant opportunity to expand automation adoption. HR Generalists at 140 actions/user are 3–4x above average: strong candidates to champion agent adoption org-wide."

---

## Key Files

| File | Purpose |
|---|---|
| `convert_dashboard_claude.py` | Main entry point for all conversions |
| `setup_pbi_mcp.py` | One-time MCP server setup for Mode 2 |
| `lib/extraction/pbix_extractor.py` | PBIX ZIP extraction and screenshot handling |
| `lib/extraction/pbip_extractor.py` | PBIP folder parsing and DAX query generation |
| `lib/rendering/chart_builder_mpl.py` | SVG chart rendering |
| `lib/rendering/builder.py` | Slide layout and PPTX assembly |
| `CLAUDE.md` | Full instructions for Claude's analysis workflow |

---

## License

MIT — use freely within your organization.

---

**Found this useful?** [Star the repo](https://github.com/shailendrahegde/pbi-to-exec-deck) to help others find it.
