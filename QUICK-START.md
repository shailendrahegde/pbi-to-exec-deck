# Quick Start Guide

Get up and running with Power BI to Executive Deck in 5 minutes.

## Prerequisites Checklist

- [ ] Claude Code CLI installed
- [ ] Power BI report exported as `.pptx`
- [ ] Working directory created (e.g., `C:\Users\[username]\claudex\`)

## 3-Step Setup

### 1. Download & Place Files

```bash
cd C:\Users\[your-username]\claudex\

# Place these files in the directory:
# ✓ Claude PowerPoint Constitution.md (from this repo)
# ✓ [Your-Dashboard].pptx (exported from Power BI)
```

### 2. Customize Constitution (Optional)

**Skip this if using default styling.**

If you have a corporate template:

1. Open `Claude PowerPoint Constitution.md`
2. Find **Section 10: Reference Templates**
3. Update file paths:

```markdown
"C:\Users\[you]\claudex\Your-Template.pptx"
```

### 3. Run the Prompt

Open Claude Code CLI and paste:

```
"[Your-Dashboard].pptx" Read this file and convert into an executive ready analytics deck following the Claude PowerPoint Constitution. Copy and paste appropriate screenshots in this new deck. Use Analytics template as a reference guide for visuals and formatting expectations. Enforce section 6 of the constitution. Each page should have compelling insights and a persuasive headline.
```

**Done!** Claude Code will generate `[Your-Dashboard] - Executive Analytics.pptx`

## Example Command

```
"AI-in-One Dashboard - v10.1.pptx" Read this file and convert into an executive ready analytics deck following the Claude PowerPoint Constitution. Copy and paste appropriate screenshots in this new deck. Use Analytics template as a reference guide for visuals and formatting expectations. Enforce section 6 of the constitution. Each page should have compelling insights and a persuasive headline.
```

## What You'll Get

**Input:** Power BI dashboard export (raw visualizations)

**Output:** Executive deck with:
- Compelling cover page and executive summary
- Insight-driven headlines (not descriptive labels)
- 2-3 insights per slide (executive-focused)
- Data-grounded recommendations (Section 6 source fidelity)
- Professional formatting

## Prompt Customization

### For deeper insights:

```
Read every page of this report and extract the most compelling insight a CIO or ITDM could glean from the data points available. Stick to the info within the report and index on actionability bringing out the so what. Generalize the patterns. State this as a storyboard. For formatting rules, refer to constitution.
```

### For specific focus areas:

```
Focus on [adoption trends/license optimization/user behavior] when creating insights.
```

### For strict data adherence:

```
Enforce section 6 of the constitution - use ONLY data visible in the dashboards.
```

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Claude can't find constitution | Ensure `Claude PowerPoint Constitution.md` is in working directory |
| Generic insights | Add "Enforce section 6" and "Extract specific data points" to prompt |
| Too much text per slide | Remind: "2-3 insights per slide for executive audience" |
| Descriptive headlines | Emphasize: "compelling insights and persuasive headlines" |

## File Locations Reference

```
Working Directory: C:\Users\[your-username]\claudex\
Constitution: Claude PowerPoint Constitution.md
Input: [Your-Dashboard].pptx
Output: [Your-Dashboard] - Executive Analytics.pptx
```

## Next Steps

1. Review generated presentation
2. Provide feedback to Claude Code for refinements
3. Iterate until presentation meets standards
4. Distribute final deck

## Need Help?

- Read full [README.md](README.md) for detailed instructions
- Check [EXAMPLE-SETUP.md](EXAMPLE-SETUP.md) for directory structure
- Review constitution file for specific rules

---

**Pro Tip:** Keep the constitution file in your working directory permanently. You can reuse it for all future dashboard conversions.
