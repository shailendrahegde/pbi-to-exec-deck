# Power BI to Executive Deck Converter

Transform Power BI dashboards into executive-ready presentations with **analyst-grade insights** in under 30 seconds.

## What You Get

**Before (Power BI):**
- Raw dashboard, interpretation burden is on you
- Generating insights and packaging it for your execs is on you

**After (Executive Deck):**
- 📊 Professional 16:9 widescreen slides
- 💡 Compelling and actionable insights
- ⚡ Generated in < 30 seconds

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
- If you are in PBI Desktop, publish the report to PBI Online in My Workspace
- In Power BI Online: **File → Export → PowerPoint**
- Save the `.pptx` file to the project directory

### 2. Run the converter (single command, within Claude Code)
```bash
python convert_dashboard_claude.py --source "your-dashboard.pptx"
```

This will automatically create `your-dashboard_executive.pptx` (or use `--output` for a custom name).

### 3. Press Enter when prompted
The script will:
- ✅ Extract dashboard images
- ✅ Claude automatically analyzes and generates insights
- ✅ Wait for you to press Enter
- ✅ Build your executive-ready presentation

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

## Need Help?

**Common issues:**

**Q: Want to run steps manually?**
A: Use `--prepare` and `--build` flags separately:
```bash
python convert_dashboard_claude.py --source "dashboard.pptx" --prepare
# Then tell Claude to analyze
python convert_dashboard_claude.py --build
```
Output will be auto-generated as `dashboard_executive.pptx`

**Q: Want different insight style?**
A: Edit the insights in `temp/claude_insights.json` before pressing Enter to continue

**Q: Claude didn't see the analysis request?**
A: Explicitly say: "Analyze the dashboards in temp/analysis_request.json and save to temp/claude_insights.json"

---

## Key Files

- `convert_dashboard_claude.py` - Main converter
- `CLAUDE.md` - Detailed workflow for Claude
- `Claude PowerPoint Constitution.md` - Quality standards
- `Example-Storyboard-Analytics.pptx` - Visual template reference

---

## Insight Quality Examples

### ✅ Good (What You'll Get)
- Concise (1-2 sentences)
- Specific numbers included
- Friendly opportunity framing
- Actionable recommendation

**Example:** "HR Generalists (140 actions/user) are 3-4x above average - great candidates to champion agent adoption across their teams"

### ❌ Bad (What We Avoid)
- Verbose paragraphs
- Generic statements without numbers
- Critical/negative tone
- No actionable takeaway

**Example:** "The data shows that there are users with varying levels of engagement across different departments which may indicate potential areas for improvement"

---

## Advanced: Batch Processing

Process multiple dashboards one at a time:
```bash
# Process each dashboard with single command
for file in dashboards/*.pptx; do
    python convert_dashboard_claude.py --source "$file"
    # Script will pause for Claude analysis, then press Enter
    # Output: dashboards/filename_executive.pptx
done
```

---

## License

MIT - Use freely for your organization

---

## Ready to Try?

**Run the 3-step process above** on your Power BI dashboard.

**You'll have a professional executive deck in under 30 seconds.**

That's it! 🚀
