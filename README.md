# Dashboards → Decks

### From raw Power BI data to executive-ready presentations — in minutes.

Turn Power BI dashboards into polished, insight-driven presentations automatically. No design skills, no manual slide-building, no copy-pasting numbers.

This tool works with **[GitHub Copilot Chat](#option-a--github-copilot-chat-recommended)** (in VS Code) or **[Claude Code](#option-b--claude-code-cli)** — pick whichever you already have.

---

## Two Modes

| | ⚡ Quick Mode | 🔬 Deep Analysis Mode |
|---|---|---|
| **Input** | PDF or PPTX export | PBIP or PBIX project file |
| **Data source** | AI reads page / slide images (OCR) | Live DAX queries via Power BI MCP (exact values) ; higher accuracy |
| **Requires MCP?** | No | Yes (one-time setup) — falls back to image analysis if MCP is unavailable |
| **Time** | **~3–5 minutes** | **~15–20 minutes** (~1 min per dashboard page) |


> **Preferred AI model:** Claude **Opus** or **Sonnet** — both provide the best insight quality. Set your model in Copilot Chat settings or Claude Code config.

---

## Option A — GitHub Copilot Chat (Recommended)

### 1. Download the repo

```bash
git clone https://github.com/shailendrahegde/pbi-to-exec-deck.git
```

### 2. Open the folder in VS Code

```bash
code pbi-to-exec-deck
```

### 3. Switch Copilot Chat to Agent Mode

1. Open the **Copilot Chat** panel (Ctrl + Shift + I, or click the Copilot icon in the sidebar).
2. In the chat input area, click the **mode selector** dropdown (it may say "Ask" or "Edit") and switch to **Agent**.
3. *(Optional)* Click the **model picker** next to it and select **Claude Sonnet** or **Claude Opus** for best results.

### 4. Ask Copilot to convert your dashboard

Type in the chat:

> Create exec deck `"C:\path\to\dashboard.pdf"`

That's it. Copilot reads the project instructions automatically ([COPILOT.md](COPILOT.md) via [.github/copilot-instructions.md](.github/copilot-instructions.md)) and runs the full pipeline — extract, analyze, build — producing a polished PPTX.

**Quick Mode (PDF / PPTX):**

> Create exec deck `"C:\path\to\dashboard.pdf"`

**Deep Analysis Mode (PBIP / PBIX):**

> Create exec deck `"C:\path\to\report.pbip"`

*(Make sure Power BI Desktop has the report open and MCP is configured — see [Deep Analysis Setup](#deep-analysis-setup-pbip--pbix) below.)*

Dependencies (`requirements-copilot.txt`) are auto-installed on first run.

---

## Option B — Claude Code (CLI)

### 1. Install Claude Code

Follow the [Claude Code CLI docs](https://docs.anthropic.com/en/docs/claude-code/overview) to install.

### 2. Clone and enter the repo

```bash
git clone https://github.com/shailendrahegde/pbi-to-exec-deck.git
cd pbi-to-exec-deck
```

### 3. Enter Claude and run the conversion

```
claude
> convert to an executive deck "C:\path\to\dashboard.pdf"
```

![Demo](demo.gif)

**Quick Mode (PDF / PPTX)** — ~3–5 minutes. **Deep Analysis (PBIP / PBIX)** — ~15–20 minutes.

Dependencies (`requirements.txt`) are auto-installed on first run.

Optional PowerShell alias (run `convert-to-exec-deck` from anywhere):

```powershell
./install-alias.ps1
```

---

## How It Works

Both assistants follow the same three-step pipeline:

### 1. Extract
Parse the source file into per-page PNG screenshots and metadata.

### 2. Analyze
The AI reads each dashboard image, extracts numbers, identifies trends, and generates analyst-grade insights following the [Insight Formula](#insight-formula). In Deep Analysis mode, the AI also runs DAX queries against the live Power BI model for exact figures.

### 3. Build
Assemble a polished 16:9 PPTX with embedded PBI screenshots, insight headlines, executive summary, and recommendations.

---

## Supported Inputs

| Input | Format | Mode | Data Source |
|---|---|---|---|
| **PDF export** | `.pdf` | ⚡ Quick | AI reads page images |
| **PPTX export** | `.pptx` | ⚡ Quick | AI reads slide images + OCR text extraction |
| **PBIP project** | `.pbip` | 🔬 Deep | Live DAX queries via MCP (exact values) + image analysis |
| **PBIX file** | `.pbix` | 🔬 Deep | Live DAX queries via MCP (exact values) + image analysis |

**How to export from Power BI (for Quick Mode):**
- **PDF** — Power BI Desktop: `File → Export → Export to PDF`
- **PPTX** — Power BI Service: `File → Export → PowerPoint`

> **Accuracy ranking:** PBIP + MCP (highest) > PBIX + MCP > PDF / PPTX (good).
> Without MCP installed, PBIP and PBIX fall back to image-only analysis automatically — still produces a complete deck, just without DAX-queried precision.

---

## Deep Analysis Setup (PBIP / PBIX)

For `.pbip` or `.pbix` files, the tool connects to the live Power BI Desktop model via the **Power BI MCP** server and queries exact values using DAX — no visual estimation needed.

> **Prefer `.pbip` over `.pbix` when you have the choice.**
> PBIP stores the semantic model as plain-text TMDL files, so the assistant can read every measure's DAX expression and understand exactly how each KPI is calculated.

### One-time setup

```bash
python setup_pbi_mcp.py          # Download & register MCP server
python setup_pbi_mcp.py --check  # Verify installation
```

This downloads the official [microsoft/powerbi-modeling-mcp](https://github.com/microsoft/powerbi-modeling-mcp) server and registers it in `.mcp.json`.

> The MCP server runs locally and communicates only with Power BI Desktop on your machine — no data leaves your environment.

### Running a deep analysis

1. Open your report in **Power BI Desktop**
2. Restart your assistant session (Claude Code: `/exit` → `claude`; Copilot: reload the window)
3. Run the conversion:

**Copilot Chat (Agent Mode):**
> Create exec deck `"C:\path\to\report.pbip"`

**Claude Code:**
```
convert to an executive deck "C:\path\to\report.pbip"
```

**Without MCP installed:** The tool falls back to image-only analysis automatically (~3–5 min instead of ~15–20 min, with lower numeric precision).

---

## What You Get

- **16:9 widescreen** slides with dark professional theme
- **PBI page screenshots** embedded by default for pixel-perfect dashboard visuals
- **Vector charts** (opt-in via `--vector-charts`) — bar, column, line, donut, KPI cards, tables, treemap, scatter, funnel, gauge
- **Insight-driven headlines** that answer "so what?" for each dashboard page
- **Executive summary** — 5 synthesized findings across all pages
- **Action recommendations** — specific, data-grounded next steps
- **Constitution validation** — automated quality checks on the output

---

## Insight Formula

Every insight follows this format:

**Headline:** `[Clear Takeaway Message]`

**Insight:** `"[Bold punchy line, 6-8 words] || [Supporting evidence with specific data]"`

**Example:**

Raw dashboard numbers:
> "1,275 active Copilot users", "134 Agent users"

Turned into analyst-grade insight:
> "134 Agent users from 1,275 total (11%) — significant opportunity to expand automation adoption. HR Generalists at 140 actions/user are 3–4x above average: strong candidates to champion agent adoption org-wide."

---

## Key Files

| File | Purpose |
|---|---|
| `convert_dashboard.py` | Main entry point — extract, analyze, build |
| `run_pipeline.py` | Single-command wrapper (deps + convert + validate) |
| `convert-to-exec-deck.cmd` | Windows batch launcher for `run_pipeline.py` |
| `check_setup.py` | Dependency validation (`--profile claude\|copilot`) |
| `setup_pbi_mcp.py` | One-time MCP server setup for PBIP/PBIX mode |
| **Extraction** | |
| `lib/extraction/extractor.py` | Markitdown-based text + metric extraction |
| `lib/extraction/text_layer_extractor.py` | PPTX/PDF text-layer enrichment |
| `lib/extraction/ocr_extractor.py` | EasyOCR fallback for dashboard PNGs |
| `lib/extraction/pdf_extractor.py` | PDF page rendering (PyMuPDF / pypdfium2) |
| `lib/extraction/pbip_extractor.py` | PBIP folder parsing and DAX query generation |
| `lib/extraction/pbix_extractor.py` | PBIX ZIP extraction and screenshot handling |
| **Rendering** | |
| `lib/rendering/builder.py` | Slide layout and PPTX assembly |
| `lib/rendering/chart_builder_mpl.py` | SVG chart rendering (matplotlib) |
| `lib/rendering/validator.py` | Constitution compliance checker |
| **Instructions** | |
| `CLAUDE.md` | Full instructions for Claude's analysis workflow |
| `COPILOT.md` | Full instructions for Copilot Chat workflow |
| `.github/copilot-instructions.md` | VS Code auto-discovery for Copilot |
| `Claude PowerPoint Constitution.md` | Quality standards and governance |

---

## License

MIT — use freely within your organization.

---

**Found this useful?** [Star the repo](https://github.com/shailendrahegde/pbi-to-exec-deck) to help others find it.
