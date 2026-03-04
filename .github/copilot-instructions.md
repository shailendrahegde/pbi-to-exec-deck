# Copilot Project Instructions

Read and follow the instructions in [COPILOT.md](../COPILOT.md) for the full workflow.
That file is the authoritative guide for converting Power BI dashboards into executive presentations.

## Quick Reference

When a user asks you to convert a dashboard (PPTX, PDF, PBIP, or PBIX):

1. **Extract:** `python convert_dashboard_claude.py --source "<path>" --prepare --assistant copilot`
2. **Analyze:** Read `temp/analysis_request.json`, read each slide image (`temp/slide_N.png`), generate analyst-grade insights, write `temp/claude_insights.json`
3. **Build:** `python convert_dashboard_claude.py --build --output "<output>.pptx"`

**Always follow the full instructions in COPILOT.md** — it contains the JSON schema, insight formula, chart spec format, quality rules, and accuracy guidelines.
