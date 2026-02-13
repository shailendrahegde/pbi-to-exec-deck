# Power BI to Executive Deck

Transform Power BI dashboards into executive-ready PowerPoint presentations in seconds using Claude Code.

## Quick Start

### 1. Clone this repository
```bash
git clone https://github.com/shailendrahegde/pbi-to-exec-deck
cd pbi-to-exec-deck
```

### 2. Export your Power BI dashboard
- In Power BI: **File → Export → PowerPoint**
- Save the `.pptx` file to this directory

### 3. Run this command in Claude Code
```
"[your-dashboard.pptx]" Read this file and convert into an executive ready analytics deck following the Claude PowerPoint Constitution. Copy and paste appropriate screenshots in this new deck. Each page in the source report should have a corresponding page in the new output deck. Use Analytics template as a reference guide for visuals and formatting expectations. Enforce section 6 of the constitution. Each page should have compelling insights and a persuasive headline.
```

**Example:**
```
"Sales-Dashboard.pptx" Read this file and convert into an executive ready analytics deck following the Claude PowerPoint Constitution. Copy and paste appropriate screenshots in this new deck. Each page in the source report should have a corresponding page in the new output deck. Use Analytics template as a reference guide for visuals and formatting expectations. Enforce section 6 of the constitution. Each page should have compelling insights and a persuasive headline.
```

That's it! Claude will generate an executive-ready presentation.

---

## What You Get

✅ **Compelling insights** - Not just data summaries
✅ **Persuasive headlines** - Insight-driven, not descriptive
✅ **Professional formatting** - Corporate template styling
✅ **Source fidelity** - All insights traceable to your data
✅ **Executive-focused** - 2-3 insights per slide, no information overload

**See example output:** Open `Example-Storyboard-Analytics.pptx` in this repository

---

## How It Works

The **Claude PowerPoint Constitution** enforces quality standards:
- Storyboard-driven narrative (not disconnected slides)
- Actionable insights with clear "so what" messaging
- Professional visual consistency
- All content grounded in your source data (no external facts)

---

## Optional Customization

Want to use your own corporate templates? Edit section 11 in `Claude PowerPoint Constitution.md` and add your template paths:

```markdown
**Optional: Customize with Your Corporate Templates**
- Analytics Decks: `C:\Path\To\Your\Analytics-Template.pptx`
- Strategy Narrative: `C:\Path\To\Your\Strategy-Template.pptx`
```

---

## Troubleshooting

**Generic insights?**
→ Add "Enforce section 6 - extract specific data points" to your prompt

**Too much text?**
→ Remind Claude: "Limit to 2-3 insights per slide for executive audience"

**Descriptive headlines?**
→ Emphasize: "compelling insights and persuasive headlines" in your prompt

---

## Learn More

- **[QUICK-START.md](QUICK-START.md)** - Detailed walkthrough
- **[EXAMPLE-STORYBOARD.md](EXAMPLE-STORYBOARD.md)** - Example output breakdown
- **[Claude PowerPoint Constitution.md](Claude%20PowerPoint%20Constitution.md)** - Full governance rules

---

## License

MIT - Use and modify freely for your organization's needs
