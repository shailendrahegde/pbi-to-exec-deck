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

## Prerequisites

**Required:**
- **Claude Code CLI** ([Installation guide](https://docs.anthropic.com/claude-code))

**Auto-detected and installed by Claude:**
- **Python 3.8+** - Claude checks if installed, guides installation if needed
- **Python packages** - Claude automatically installs: `python-pptx`, `Pillow`, `PyMuPDF`, `markitdown`

**That's it!** Claude Code handles all dependency detection and installation automatically.

---

## Installation & Setup (2 Steps)

### Step 1: Clone and open in Claude Code
```bash
git clone https://github.com/shailendrahegde/pbi-to-exec-deck.git
cd pbi-to-exec-deck
claude
```

### Step 2: Let Claude handle the rest
Just ask Claude to convert your dashboard (provide the file path):
```
"Convert C:\Users\you\Downloads\dashboard.pdf to an executive deck"
```

Or with relative path:
```
"Convert ./my-dashboard.pdf to an executive deck"
```

**Claude automatically:**
- ✓ Detects if Python and dependencies are installed
- ✓ Installs missing packages (`python-pptx`, `Pillow`, `PyMuPDF`, `markitdown`)
- ✓ Runs the conversion
- ✓ Creates your executive presentation

**No manual pip install needed!** This is the advantage of running through Claude Code.

---

### Advanced: Running Without Claude Code (Optional)

If you prefer to run the script directly:

```bash
# Install dependencies manually
pip install -r requirements.txt

# Verify setup (optional)
python check_setup.py

# Run conversion
python convert_dashboard_claude.py --source your-dashboard.pdf
```

**Note:** Running through Claude Code is recommended for the best experience.

---

## Quick Start

### 1. Export your Power BI dashboard

**Option A: PDF Export (from Power BI Desktop)**
- In Power BI Desktop: **File → Export → Export to PDF**
- Exports all dashboard pages as PDF
- Save the `.pdf` file

**Option B: PPTX Export (from Power BI Online)**
- Publish your report from PBI Desktop to Power BI Online (My Workspace)
- In Power BI Online: **File → Export → PowerPoint (Static Images)**
- Exports dashboard as PPTX with static images
- Save the `.pptx` file

**Both formats work identically** - the converter produces the same high-quality executive presentation output

### 2. Open in Claude Code and convert

```bash
cd pbi-to-exec-deck
claude
```

Then in Claude Code, provide the path to your dashboard:
```
"Convert C:\Users\you\Downloads\my-dashboard.pdf to an executive deck"
```

Or with relative path if file is in the project directory:
```
"Convert ./my-dashboard.pdf to an executive deck"
```

Or run the script directly:
```bash
python convert_dashboard_claude.py --source "your-dashboard.pdf"
```

### 3. Get your executive deck

Claude will:
- ✅ Install any missing dependencies automatically
- ✅ Extract dashboard images (PDF or PPTX)
- ✅ Analyze and generate analyst-grade insights
- ✅ Build your executive-ready presentation (16:9 PPTX format)

**Done in ~30 seconds!** Open `your-dashboard_executive.pptx` to see your professional presentation.

---

## What Makes This Different

✅ **Analyst-grade insights** - Not just data restatements
✅ **No API key needed** - Uses your Claude Code session
✅ **Automatic setup** - Claude installs dependencies for you
✅ **16:9 widescreen** - Modern presentation format
✅ **Specific numbers** - Every insight backed by data
✅ **Friendly tone** - Opportunities, not criticisms
✅ **Fast** - Complete in under 30 seconds

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
