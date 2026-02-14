# Power BI to Executive Deck Converter

## Stop wrestling with slides. Get executive-ready presentations in 30 seconds.

Building slides is time-consuming. Extracting insights from dashboards is hard. Packaging everything for executives takes real effort. This tool handles all three automatically.

Transform Power BI dashboards (PDF or PPTX) into professional presentations with analyst-grade insights. No design skills required.

## What You Get

**Input:** Raw Power BI dashboard export
**Output:** Professional 16:9 presentation with:
- Clean, formatted slides
- Analyst-grade insights with specific numbers
- Actionable recommendations
- Executive-friendly language

**Time:** ~30 seconds total

### Example Transformation

**Generic Statement:**
> "There are 1,275 active Copilot users"

**Analyst-Grade Insight:**
> "134 Agent users from 1,275 total (11%) - significant opportunity to expand automation adoption. HR Generalists (140 actions/user) are 3-4x above average - great candidates to champion agent adoption"

---

## Installation & Setup

### Step 1: Clone the repository
```bash
git clone https://github.com/shailendrahegde/pbi-to-exec-deck.git
cd pbi-to-exec-deck
```

### Step 2: Open in Claude Code (IMPORTANT)
```bash
claude
```

That's it! All dependencies are already included. No API keys, no OCR installation required.

---

## Quick Start (3 Steps)

### 1. Export your Power BI dashboard

**Option A: PDF Export (from Power BI Desktop)**
- In Power BI Desktop: **File → Export → Export to PDF**
- Exports all dashboard pages as PDF
- Save the `.pdf` file to the project directory

**Option B: PPTX Export (from Power BI Online)**
- Publish your report from PBI Desktop to Power BI Online (My Workspace)
- In Power BI Online: **File → Export → PowerPoint (Static Images)**
- Exports dashboard as PPTX with static images
- Save the `.pptx` file to the project directory

**Both formats work identically** - the converter produces the same high-quality executive presentation output

### 2. Run the converter (single command)

```bash
python convert_dashboard_claude.py --source "your-dashboard.pdf" --auto
```

Works with both `.pdf` and `.pptx` files. Output will be `your-dashboard_executive.pptx` (or use `--output` for a custom name).

### 3. Workflow completion
The script will:
- ✅ Extract dashboard images (PDF or PPTX)
- ✅ Claude automatically analyzes and generates insights
- ✅ Build your executive-ready presentation (16:9 PPTX format)

**Done!** Open `your-dashboard_executive.pptx` to see your professional presentation.

---

## What Makes This Different

✅ **Analyst-grade insights** - Not just data restatements
✅ **No API key needed** - Uses your Claude Code session
✅ **16:9 widescreen** - Modern presentation format
✅ **Specific numbers** - Every insight backed by data
✅ **Friendly tone** - Opportunities, not criticisms
✅ **Fast** - Complete in under 30 seconds
✅ **No setup required** - All dependencies included

---

## Sample Insights Generated

**Real examples from dashboard conversions:**
- "Finance achieves 217 prompts per user (5.7x average) - proven pattern to replicate"
- "HR Generalists (22 users, 140 prompts) have embedded workflows - leverage as champions"
- "SW Builder leads with 546 actions - successful use case to build on"

---

## How It Works

**Step 1 (Deterministic):** Python extracts dashboard images and structure
**Step 2 (Intelligent):** Claude analyzes dashboards as senior analyst advisor
**Step 3 (Deterministic):** Python renders professional slides with insights

**Key insight:** Technical tasks are automated, strategic analysis uses AI.

---

## Key Files

- `convert_dashboard_claude.py` - Main converter
- `CLAUDE.md` - Detailed workflow for Claude
- `Claude PowerPoint Constitution.md` - Quality standards
- `Example-Storyboard-Analytics.pptx` - Visual template reference



## License

MIT - Use freely for your organization

---

## Ready to Try?

**Run the 3-step process above** on your Power BI dashboard.

**You'll have a professional executive deck in under 30 seconds.**

---

**Found this useful?** ⭐ [Star this repo](https://github.com/shailendrahegde/pbi-to-exec-deck) to help others discover it!

That's it! 🚀
