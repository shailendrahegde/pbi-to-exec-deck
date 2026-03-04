# Copilot Project Instructions

Read and follow the instructions in [COPILOT.md](../COPILOT.md) for the full workflow.
That file is the authoritative guide for converting Power BI dashboards into executive presentations.

## Quick Reference

When a user asks you to convert a dashboard (PPTX, PDF, PBIP, or PBIX):

1. **Extract:** `python convert_dashboard.py --source "<path>" --prepare --assistant copilot`
2. **Analyze:** Read `temp/analysis_request.json`, read each slide image (`temp/slide_N.png`), generate analyst-grade insights, write `temp/insights.json`
3. **Verify:** `python convert_dashboard.py --verify` (catches missing charts, vanilla headlines, slide count mismatches)
4. **Build:** `python convert_dashboard.py --build --output "<output>.pptx"` (also runs verify automatically)

**Critical quality rules:**
- Every slide with quantitative data MUST have at least one `"chart"` spec — don't let slides fall back to raw screenshots
- If a slide has `"ocr_used": true`, cross-verify every number-to-label pairing against the image — OCR context can misattribute
- Every headline must pass the "Would a VP forward this?" test — no vanilla statements like "shows the breakdown"

**Always follow the full instructions in COPILOT.md** — it contains the JSON schema, insight formula, chart spec format, quality rules, and accuracy guidelines.
