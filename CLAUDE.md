# Claude Project Instructions

This file provides persistent instructions for Claude when working in this project.

## Project Overview

This project converts Power BI dashboard exports (.pptx or .pdf) into executive-ready analytics presentations with **compelling, analyst-grade insights** using Claude's vision and analytical capabilities.

**Supported Input Formats:**
- PowerPoint (.pptx) - Exported from Power BI
- PDF (.pdf) - Exported from Power BI
- **Note:** Both formats produce identical output quality (16:9 PPTX)

## Core Philosophy

**Deterministic for Technical Tasks, Intelligent for Insights**

- ✅ **Deterministic Path:** Extract images, parse structure, render slides (fast, reliable)
- ✅ **Intelligent Path:** Analyze dashboards, identify gaps, generate strategic insights (high quality)
- ✅ **No API Key Required:** Uses current Claude Code session
- ✅ **No OCR Dependencies:** Works with any Power BI dashboard export

## Standard Workflow: Claude-Powered Conversion

When a user requests dashboard conversion, they run a **single command** that orchestrates all steps:

```bash
python convert_dashboard_claude.py --source "dashboard.pptx"
```

The output file will be automatically named `dashboard_executive.pptx` (or use `--output` for a custom name).

### What Happens Automatically:

**Step 1: Extract (5 seconds)**
- Script extracts each dashboard slide/page as PNG image to `temp/` directory
- **For PPTX:** Extracts embedded images from slides
- **For PDF:** Converts each page to PNG image at 150 DPI
- **PPTX workflow:** Automatically skips first slide (title page with metadata like "last refreshed", "view in PowerBI")
- **PDF workflow:** Includes all pages (PDF exports typically have dashboard content on page 1)
- Parses slide/page titles and structure
- Creates `temp/analysis_request.json` with slide metadata

**Step 2: Claude Analysis (You!)**
- Script displays clear request for Claude to analyze dashboards
- You (Claude Code) automatically see the request and respond
- **Your task:** Analyze each dashboard image and generate insights
- **Save results to:** `temp/claude_insights.json`

**Step 3: Build (3 seconds)**
- After you finish, user presses Enter
- Script loads your insights from JSON
- Creates professional slides (16:9 widescreen)
- Embeds dashboard images with insights
- Applies Analytics template styling
- Validates against Constitution

**Total time:** ~30 seconds including your analysis

---

## Your Task: Generate Analyst-Grade Insights

When the script displays the analysis request, you should:

### 1. Read the analysis request
```
Read temp/analysis_request.json
```

### 2. For each slide, read the dashboard image
```
Read temp/slide_1.png
Read temp/slide_2.png
... (for all slides listed)
```

### 3. Act as senior analyst advisor to IT decision maker
- Identify key numbers, trends, patterns
- Spot gaps and opportunities
- Generate strategic insights with "so what"
- Provide actionable recommendations

### 4. Write insights to JSON

**IMPORTANT: After analyzing all slides, also generate:**

**Executive Summary (5 bullets):**
- Synthesize the most compelling insights across ALL dashboard pages
- Prioritize by business impact (highest value first)
- Each bullet must reference specific numbers from the data
- Answer: "What are the 5 things an executive MUST know?"
- Format: "[Specific finding] → [Business implication]"

**Next Steps (3-5 recommendations):**
- Actionable recommendations grounded in the data
- Prioritize by implementation value and feasibility
- Each must trace back to specific insights from the dashboard
- Format: "[Action verb] + [what to do] + [expected outcome]"
- Be specific: "Pilot agent training with HR Generalists (140 actions/user) to establish best practices" NOT "Improve training"

```json
{
  "deck_title": "Compelling, persuasive title for the presentation",
  "deck_subtitle": "Optional subtitle — e.g. date range, org name, or focus area",
  "executive_summary": [
    "Finding with specific number → business implication",
    "Finding with specific number → business implication",
    "Finding with specific number → business implication",
    "Finding with specific number → business implication",
    "Finding with specific number → business implication"
  ],
  "recommendations": [
    "Action: Specific recommendation with expected outcome",
    "Action: Specific recommendation with expected outcome",
    "Action: Specific recommendation with expected outcome"
  ],
  "slides": [
    {
      "slide_number": 1,
      "title": "Slide Title",
      "headline": "[Number] + [Insight]",
      "insights": [
        "Punchy bold line under 8 words || Supporting evidence with specific data and business implication",
        "Punchy bold line under 8 words || Pattern or opportunity identified with data",
        "Punchy bold line under 8 words || Actionable recommendation with expected outcome"
      ],
      "numbers_used": ["134", "1,275", "11%"]
    }
  ]
}
```

**Save to:** `temp/claude_insights.json`

---

## Deck Title Guidelines

**ALWAYS generate `deck_title` and `deck_subtitle`.** Never leave these blank — the default "Executive Insights" is generic and forgettable.

**`deck_title`** — The cover slide headline. Should be:
- Compelling and persuasive, not descriptive
- Positive framing: opportunities and wins, not gaps or problems
- Specific to the content: reference the org, product, or key finding
- 5–10 words max

**Good examples:**
- ✅ "Swetha Org: Strong Copilot Momentum with Clear Scaling Opportunities"
- ✅ "M365 Copilot: 40% Power Users and Growing"
- ✅ "Copilot Adoption Accelerating — Efficiency and Wellbeing Gains Confirmed"

**Bad examples:**
- ❌ "Executive Insights" (generic default)
- ❌ "Copilot Usage Report" (descriptive, not persuasive)
- ❌ "Low Adoption in Some Teams" (negative framing)

**`deck_subtitle`** — One line of context. Use the date range from the dashboard filters and/or the org/leader name.
- Example: "Swetha Org · Dec 2025 – Feb 2026"
- Example: "Super User Analysis · Q1 2026"

---

## Manual Workflow (Optional)

If users want step-by-step control:

```bash
# Step 1: Prepare only
python convert_dashboard_claude.py --source "dashboard.pptx" --prepare

# Step 2: You analyze (same as above)

# Step 3: Build only
python convert_dashboard_claude.py --build --output "executive.pptx"
```

---

## Insight Generation Guidelines

**CRITICAL: Before analyzing any dashboard, read `DASHBOARD_READING_RULES.md` for data extraction accuracy rules.**

When generating insights (Step 2), follow these principles:

### ✅ DO: Be Concise and Friendly

**Good Example:**
- Headline: "134 Agent users from 1,275 total (11%) - significant opportunity to expand automation adoption"
- Insight: "HR Generalists (140 actions/user) are 3-4x above average - great candidates to champion agent adoption"

**Why it works:** Specific numbers, identifies opportunity, suggests action - all in one sentence

### ✅ DO: Focus on Opportunities

Frame observations as opportunities rather than problems:
- ✅ "Opportunity to expand agent adoption from 11% to 30%"
- ❌ "Only 11% adoption shows deployment failure"

### ✅ DO: Provide Specific Numbers (Match Units Exactly)

Every headline and at least one insight should include concrete numbers:
- ✅ "217 prompts per user (5.7x average) demonstrates high-value workflow"
- ❌ "Some users show higher engagement than others"

**CRITICAL: Match number units EXACTLY as shown in the visual:**
- Dashboard shows "13" → Use "13 actions" (NOT "13K")
- Dashboard shows "13K" → Use "13K actions" (NOT "13,000" or "13")
- Dashboard shows "13M" → Use "13M actions" (NOT "13 million")
- Dashboard shows "13.5K" → Use "13.5K actions" (NOT "13K")

**Test each number:**
1. Can I point to this exact number on the dashboard?
2. Does the unit (K, M, %, decimal) match exactly?
3. If NO to either → Do NOT use that number

**Why this matters:** Unit mismatches (13 vs 13K) destroy credibility instantly

### ✅ DO: Answer "So What?"

Every insight should help IT decision maker take action:
- ✅ "Client Finance's 217 prompts/user pattern ready to replicate across Corporate Finance"
- ❌ "Client Finance has 217 prompts per user"

### ✅ DO: Analyze Platform/Feature Patterns

**CRITICAL: Only mention entities (platforms, teams, features) that are VISIBLE on the dashboard page.**

Always look for and call out platform, app, or feature variations **IF they appear on the visual:**
- ✅ "Teams integration shows 3x higher engagement than Outlook" — **ONLY if both "Teams" and "Outlook" are visible on this page**
- ✅ "PowerPoint Copilot leads with 450 actions while Excel shows only 89" — **ONLY if both apps are shown on this page**
- ✅ "Chat (Web) dominates with 1,275 users" — **ONLY if "Chat (Web)" label is visible on this page**
- ❌ "Users engage with the platform" (missing which platform/feature)
- ❌ Mentioning "Teams" when it's not shown on this specific page

**Rule: Can you point to where this name appears on the dashboard?**
- **YES** → Safe to mention
- **NO** → Do NOT mention, even if you think it's on other pages

**What to look for (only if visible on this page):**
- Microsoft 365 apps: Outlook, Teams, Word, Excel, PowerPoint, OneNote
- Copilot features: Chat, Agents, M365 Copilot, Business Chat
- Integration patterns: Web vs. integrated, standalone vs. embedded
- Feature adoption: Which capabilities drive engagement
- Platform concentration: Where is usage concentrated
- Department names, team names, location names

**Why this matters:**
- Reveals which workflows deliver value (content creation vs. analysis vs. communication)
- Identifies integration successes and failures
- Shows where to focus training and enablement
- Guides feature prioritization decisions

**Why the visibility rule matters:**
- Every claim must be verifiable from the visual
- Users will check if "Teams" actually appears on that page
- One unverifiable claim destroys credibility

### ❌ DON'T: Be Critical or Verbose

Avoid:
- Critical language: "critical failure", "deployment disaster", "massive gap"
- Extreme emotional language: "crisis", "catastrophically", "blocking" - use "opportunity", "challenge", "limiting" instead
- Verbose explanations: Keep insights to 1-2 sentences max
- Criticizing the analytics: Don't say "missing data" or "blind spot"
- Technical jargon: Use business language

**Tone guidelines for executives:**
- ✅ "habit formation opportunity" ❌ "habit formation crisis"
- ✅ "significant cleanup opportunity" ❌ "massive cleanup required"
- ✅ "opportunity to unlock features" ❌ "deployment gap blocking features"
- ✅ "concentration in infrequent tier" ❌ "catastrophically low usage"

**Why:** Extreme language triggers defensive reactions. Data-driven opportunities encourage action.

### ❌ DON'T: Make Generic Statements

Avoid statements that don't help decision making:
- ❌ "There are 1,275 active users"
- ❌ "Platform shows engagement"
- ❌ "Usage varies by department"

### ❌ DON'T: Force Insights Without Data

**CRITICAL: Credibility over completeness**

If a dashboard page has no numbers or is blank:
- ✅ **DO:** Mark as "Insufficient data for analysis"
- ✅ **DO:** Skip the slide entirely
- ✅ **DO:** Explain what data is needed
- ❌ **DON'T:** Generate vague insights to fill space
- ❌ **DON'T:** Make assumptions about missing data

**Why this matters:**
- User trust depends on accuracy and honesty
- One unsupported insight undermines all insights
- Better to deliver 8 strong insights than 10 with 2 weak ones

**Example when data is missing:**
```json
{
  "slide_number": 7,
  "title": "Agents - Habit Formation",
  "headline": "Insufficient data for analysis",
  "insights": [
    "Dashboard page contains no extractable usage frequency metrics",
    "To generate habit formation insights, need: daily/weekly/monthly usage distribution, engagement tiers, or frequency patterns",
    "Recommend adding frequency tracking to dashboard for next analysis cycle"
  ],
  "numbers_used": []
}
```

---

## Insight Formula

Follow this proven formula for each slide:

**Headline:** `[Clear Takeaway Message]` - Focus on instant clarity without requiring numbers

**IMPORTANT:** Headlines should answer "so what?" and be memorable. Numbers are NOT required in headlines - they belong in the supporting insights where they provide evidence.

Good headline examples:
- ✅ "Multi-platform adoption demonstrates organization-wide AI readiness"
- ✅ "Visual Creator emerges as winning agent template for scaling"
- ✅ "Habit formation gap prevents agents from becoming daily workflows"
- ✅ "Operations establishes benchmark for M365 Copilot excellence"

Bad headline examples:
- ❌ "4,381 total active users across Agent (180), Unlicensed Chat (2,585)..." (too many numbers, unclear takeaway)
- ❌ "101 of 116 total agent users stay in infrequent tier with zero daily users" (data dump, not insight)

**Insight format: two-part, separated by `||`**

Each insight must follow this exact structure:
`"[Bold punchy line, 6-8 words max] || [Supporting evidence with specific data]"`

The bold line is rendered large and bold on the slide. The detail is rendered below in normal weight. Keep them clearly separated — the bold line is the "so what", the detail is the proof.

**Insight 1 (Scale/Scope):**

Example: `"HR Generalists lead usage by far || 140 actions/user is 3-4x the average — strong candidates to champion adoption"`

**Insight 2 (Pattern/Opportunity):**

Example: `"Client Finance pattern ready to replicate || 217 prompts/user sets the benchmark for Corporate Finance expansion"`

**Insight 3 (Action):**

Example: `"Pilot training with top 3 departments || Target 50+ prompts/user baseline to establish org-wide standard"`

**Remember:** The headline is what an executive will remember 3 days later. Make it clear, memorable, and clearly supported by the numbered evidence in your insights.

---

## Technical Requirements

### Image Analysis
- Dashboard images are 2560x1460 pixels (Power BI export default)
- Look for: KPI cards, trends charts, tables, leaderboards
- Extract: Numbers, percentages, trends, comparisons, distributions
- **CRITICAL: Identify platforms, apps, and features** (Outlook, Teams, PowerPoint, Excel, Agents, Chat, etc.)
- Notice engagement variations across different platforms/features
- Use platform patterns to identify high-value workflows

### Data Extraction Accuracy Rules (CRITICAL)

**Read `DASHBOARD_READING_RULES.md` for complete guidelines. Key rules:**

1. **Match numbers to EXACT labels** - Read table labels carefully; "Team C: 73,071" not "Team B: 73,071"
2. **Never assume first row is aggregate** - Read actual label; could be "Team B" not "Total"
3. **Read chart axes precisely** - Verify scale/units; don't guess values
4. **Identify selected filters** - Black/highlighted = active; chart shows only selected filter data
5. **Match units exactly** - "13K" shown → write "13K" (not "13,000")
6. **User categories exact** - "148 Unlicensed" not "148 Bottom 25%" if label differs
7. **Only cite visible numbers** - Can you point to it? If no, don't use it

**Pre-analysis checklist for each page:**
- [ ] Identify filter selections (black vs grey elements)
- [ ] Read all table row/column labels
- [ ] Check chart axis scales and units
- [ ] Read legend for color meanings
- [ ] Verify every number matches its label

### Slide Types to Recognize
- **Trends:** Time-series charts → Focus on momentum, growth, patterns
- **Leaderboards:** Rankings/tables → Focus on top performers, concentration, gaps
- **Health Check:** KPI dashboards → Focus on portfolio health, key metrics
- **Habit Formation:** Frequency distributions → Focus on engagement tiers, usage patterns
- **License Priority:** User segments → Focus on upgrade candidates, ROI opportunities
- **Platform/Feature Comparison:** Usage across apps → Focus on which platforms drive value, integration patterns

### Output Format
- 16:9 widescreen (13.333" x 7.5")
- Image on left (6.5" max width)
- Insights on right (5.5" width)
- 3 bullet points per slide maximum
- Segoe UI font, blue accent colors

---

## Quality Validation

After generating insights, verify:

✅ **Every headline has specific number** (not "some users" or "many")
✅ **Number units match exactly** (13 vs 13K vs 13M - use what's shown)
✅ **All entities are visible** (platforms, teams, features mentioned are on this page)
✅ **Insights are concise** (1-2 sentences each, not paragraphs)
✅ **Tone is friendly** (opportunities, not failures)
✅ **Focus on action** (what to do, not just what is)
✅ **Numbers are traceable** (can point to exact location on dashboard)
✅ **No criticism** (don't critique the analytics or report)
✅ **No forced insights** (if no data, mark "Insufficient data" rather than generating generic content)
✅ **Platform patterns identified** (mention specific apps/features ONLY when visible on this page)
✅ **Executive summary synthesizes ALL pages** (not just first slide)
✅ **Recommendations are actionable** (specific actions, not vague suggestions)
✅ **Both sections grounded in data** (can trace each point to dashboard numbers)

---

## Success Criteria

When the user says "Convert X to executive deck":

✅ Output file created at expected path (16:9 widescreen)
✅ All source slides have corresponding output slides (1:1 mapping)
✅ Dashboard images embedded on left side of each slide
✅ Headlines are insight-driven with specific numbers
✅ Insights are concise (1-2 sentences), friendly, actionable
✅ No generic statements like "there are X users"
✅ Visual formatting follows Analytics template (blue, professional)
✅ Constitution rules followed (Section 4-6)
✅ Images placed without overlap with text
✅ 2-3 insights per slide (executive-focused)
✅ **Execution time < 30 seconds total**

---

## Example Workflow in Practice

**User says:** "Convert wpp22.pptx to executive deck"

**You do:**

1. Run prepare command:
   ```bash
   python convert_dashboard_claude.py --source wpp22.pptx --prepare
   ```

2. Read analysis request and view images:
   ```
   Read temp/analysis_request.json
   Read temp/slide_2.png
   Read temp/slide_3.png
   ... (view all slides)
   ```

3. Generate concise, friendly insights for each slide following formula

4. Save insights JSON to temp/claude_insights.json

5. Build final presentation:
   ```bash
   python convert_dashboard_claude.py --build --output wpp22_executive.pptx
   ```

6. Verify output and report success

**Total time:** < 30 seconds
**Quality:** Analyst-grade insights, professional appearance

---

## Files Reference

- `convert_dashboard_claude.py` - Main conversion orchestrator
- `Claude PowerPoint Constitution.md` - Quality standards and governance
- `Example-Storyboard-Analytics.pptx` - Visual template reference (use for styling)
- `lib/rendering/builder.py` - Slide rendering (16:9 format)
- `lib/rendering/validator.py` - Constitution compliance checker
- `temp/` directory - Temporary files for analysis (auto-generated, not committed)

---

## Error Handling

If conversion fails:

1. **Check file paths** - Use absolute paths, proper quotes
2. **Check temp/ directory** - Should contain extracted images
3. **Check insights JSON** - Verify format and completeness
4. **Re-run failed step** - Prepare, analyze, or build independently
5. **Report clearly** - State which step failed and why

---

## Key Advantages of This Approach

✅ **No API key needed** - Uses current Claude Code session
✅ **High-quality insights** - Real intelligence, not rule-based patterns
✅ **Fast execution** - < 30 seconds total
✅ **Scalable** - Works for any user without configuration
✅ **Maintainable** - Separate deterministic and intelligent layers
✅ **Professional output** - 16:9 widescreen, Analytics template styling
✅ **Constitution-compliant** - Automated validation built-in
