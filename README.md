# Power BI to Executive Deck Converter

Transform Power BI dashboards into executive-ready presentations with **analyst-grade insights** in under 30 seconds.

## What You Get

**Before (Power BI Export):**
- Raw dashboard screenshots
- No insights or narrative
- Just data visualizations

**After (Executive Deck):**
- 📊 Professional 16:9 widescreen slides
- 💡 Compelling insights with specific numbers
- 🎯 Actionable recommendations
- ⚡ Generated in < 30 seconds

### Example Transformation

**Generic Statement:**
> "There are 1,275 active Copilot users"

**Analyst-Grade Insight:**
> "134 Agent users from 1,275 total (11%) - significant opportunity to expand automation adoption. HR Generalists (140 actions/user) are 3-4x above average - great candidates to champion agent adoption"

---

## Quick Start (3 Steps)

### 1. Export your Power BI dashboard
- In Power BI: **File → Export → PowerPoint**
- Save the `.pptx` file

### 2. Run the converter
```bash
# Step 1: Extract dashboards
python convert_dashboard_claude.py --source "your-dashboard.pptx" --prepare

# Step 2: In Claude Code, say:
"Generate analyst insights for the dashboards"

# Step 3: Build final deck
python convert_dashboard_claude.py --build --output "executive-deck.pptx"
```

### 3. Open your executive-ready deck
Done! You now have a professional presentation with compelling insights.

---

## What Makes This Different

✅ **Analyst-grade insights** - Not just data restatements
✅ **No API key needed** - Uses your Claude Code session
✅ **16:9 widescreen** - Modern presentation format
✅ **Specific numbers** - Every insight backed by data
✅ **Friendly tone** - Opportunities, not criticisms
✅ **Fast** - Complete in under 30 seconds

---

## Installation

Already done if you're in this repo! Dependencies:
- `python-pptx` ✓
- `markitdown` ✓
- `Pillow` ✓

No OCR or API keys required.

---

## Sample Insights Generated

**Real examples from dashboard conversions:**
- "Client Finance achieves 217 prompts per user (5.7x average) - proven pattern to replicate"
- "HR Generalists (22 users, 140 prompts) have embedded workflows - leverage as champions"
- "Draft as IP Agent leads with 546 actions - successful use case to build on"

---

## How It Works

**Step 1 (Deterministic):** Python extracts dashboard images and structure
**Step 2 (Intelligent):** Claude analyzes dashboards as senior analyst advisor
**Step 3 (Deterministic):** Python renders professional slides with insights

**Key insight:** Technical tasks are automated, strategic analysis uses AI.

---

## Need Help?

**Common issues:**

**Q: Claude didn't generate insights?**
A: After running `--prepare`, explicitly ask Claude: "Generate analyst insights for the dashboards in temp/analysis_request.json"

**Q: Want different insight style?**
A: Edit the insights in `temp/claude_insights.json` before running `--build`

**Q: Output not wide format?**
A: Delete old output and re-run `--build` - latest version uses 16:9

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

Process multiple dashboards:
```bash
for file in dashboards/*.pptx; do
    python convert_dashboard_claude.py --source "$file" --prepare
done

# Claude generates insights for all

for file in dashboards/*.pptx; do
    output="${file%.pptx}_executive.pptx"
    python convert_dashboard_claude.py --build --output "$output"
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
