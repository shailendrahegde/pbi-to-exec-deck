# Dashboards → Decks

### From raw Power BI data to executive-ready presentations — in minutes.

Turn Power BI dashboards into polished, insight-driven presentations automatically. No design skills, no manual slide-building, no copy-pasting numbers.

This tool works with **[GitHub Copilot Chat](#option-a--github-copilot-chat)** (in VS Code) or **[Claude Code](#option-b--claude-code-cli)** — pick whichever you already have.

---

## Two Modes

| | ⚡ Quick Mode | 🔬 Deep Analysis Mode |
|---|---|---|
| **Input** | PDF or PPTX export | PBIP or PBIX project file |
| **Data source** | AI reads page / slide images (OCR) | Live DAX queries via Power BI MCP (exact values) |
| **Requires MCP?** | No | Yes ([one-time setup](#deep-analysis-setup-pbip--pbix)) — falls back to image analysis if unavailable |
| **Time** | **~3–5 min** | **~15–20 min** |

> **Preferred AI model:** Claude **Opus** or **Sonnet**. Set your model in Copilot Chat settings or Claude Code config.

---

## Option A — GitHub Copilot Chat

1. `git clone https://github.com/shailendrahegde/pbi-to-exec-deck.git`
2. Open the folder in VS Code: `code pbi-to-exec-deck`
3. Open **Copilot Chat** → switch mode to **Agent** → *(optional)* select **Claude Sonnet/Opus**
4. Type:

> Create exec deck `"C:\path\to\dashboard.pdf"`

That's it. Dependencies auto-install on first run.

---

## Option B — Claude Code (CLI)

```bash
git clone https://github.com/shailendrahegde/pbi-to-exec-deck.git
cd pbi-to-exec-deck
claude
> create exec deck "C:\path\to\dashboard.pdf"
```

![Demo](demo.gif)

Dependencies auto-install on first run. Optional alias: `./install-alias.ps1`

---

## What You Get

- **16:9 widescreen** slides with dark professional theme
- **PBI page screenshots** embedded (or **vector charts** with `--vector-charts`)
- **Insight-driven headlines** that answer "so what?" for each dashboard page
- **Executive summary** and **action recommendations** grounded in data
- **Constitution validation** — automated quality checks on the output

---

## Advanced Options

### Vector Charts

By default, the deck embeds PBI screenshots. Add `--vector-charts` to generate clean charts (bar, column, line, donut, KPI tiles, treemap, table, etc.) directly from the data.

**When to use:** screenshots are low-quality, you want resolution-independent visuals, or exact DAX data is available.

> Create exec deck `"C:\path\to\report.pbip"` --vector-charts

### Contextual Focus

Add context to steer the analysis toward a specific department, audience, or theme:

> Create exec deck `"C:\path\to\report.pbip"` — focus on the Finance department

> Create exec deck `"C:\path\to\dashboard.pdf"` — this is for the CISO; emphasize security and compliance metrics

> Create exec deck `"C:\path\to\report.pbip"` --vector-charts — focus on May 2025 month-over-month changes

> Create exec deck `"C:\path\to\report.pbip"` --vector-charts — focus on HR and Operations; frame recommendations around reducing onboarding time

Everything after the file path is treated as analyst guidance — adjust framing, tone, and priorities as needed.

---

## Deep Analysis Setup (PBIP / PBIX)

For `.pbip` or `.pbix` files, the tool connects to Power BI Desktop via the **Power BI MCP** server and queries exact values using DAX.

> **Prefer `.pbip` over `.pbix`** — PBIP exposes measure DAX expressions so the AI understands how each KPI is calculated.

### One-time setup

```bash
python setup_pbi_mcp.py          # Download & register MCP server
python setup_pbi_mcp.py --check  # Verify installation
```

Downloads the official [microsoft/powerbi-modeling-mcp](https://github.com/microsoft/powerbi-modeling-mcp) server and registers it in `.mcp.json`. Runs locally — no data leaves your machine.

### Running

1. Open your report in **Power BI Desktop**
2. Restart your assistant session (Claude Code: `/exit` → `claude`; Copilot: reload window)
3. `Create exec deck "C:\path\to\report.pbip"`

Without MCP installed, PBIP/PBIX falls back to image-only analysis automatically.

---

<details>
<summary><strong>Supported Inputs</strong></summary>

| Input | Format | Mode | Data Source |
|---|---|---|---|
| **PDF export** | `.pdf` | ⚡ Quick | AI reads page images |
| **PPTX export** | `.pptx` | ⚡ Quick | AI reads slide images + OCR |
| **PBIP project** | `.pbip` | 🔬 Deep | Live DAX queries via MCP + images |
| **PBIX file** | `.pbix` | 🔬 Deep | Live DAX queries via MCP + images |

**How to export from Power BI (Quick Mode):**
- **PDF** — Power BI Desktop: `File → Export → Export to PDF`
- **PPTX** — Power BI Service: `File → Export → PowerPoint`

</details>

<details>
<summary><strong>Key Files</strong></summary>

| File | Purpose |
|---|---|
| `convert_dashboard.py` | Main entry point — extract, analyze, build |
| `run_pipeline.py` | Single-command wrapper (deps + convert + validate) |
| `setup_pbi_mcp.py` | One-time MCP server setup for PBIP/PBIX mode |
| `lib/extraction/` | Source file parsing (PDF, PPTX, PBIP, PBIX, OCR) |
| `lib/rendering/builder.py` | Slide layout and PPTX assembly |
| `lib/rendering/chart_builder_mpl.py` | Vector chart rendering (matplotlib) |
| `lib/rendering/validator.py` | Constitution compliance checker |
| `CLAUDE.md` / `COPILOT.md` | AI workflow instructions |

</details>

---

## License

MIT — use freely within your organization.

---

**Found this useful?** [Star the repo](https://github.com/shailendrahegde/pbi-to-exec-deck) to help others find it.
