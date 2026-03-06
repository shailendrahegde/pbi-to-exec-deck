# Credibility Over Completeness Principle

## Core Philosophy

**Better to skip 2 slides than generate 2 weak insights that undermine trust in 8 strong insights.**

---

## The Trust Problem

### Scenario: Analyst generates 10 insights

**Case 1: All 10 have data support**
- User confidence: HIGH
- All insights trusted and acted upon
- Analyst credibility established

**Case 2: 8 have data, 2 are generic filler**
- User discovers 2 weak insights
- Questions validity of ALL 10 insights
- User confidence: LOW
- Strong insights now doubted
- Analyst credibility damaged

**Result:** The 2 generic insights don't add value—they subtract trust.

---

## When to Skip or Mark "Insufficient Data"

### Clear Skip Situations

**1. Dashboard page is blank or empty**
- No charts, no numbers, no visualizations
- **Action:** Skip entirely or mark "No data available"

**2. Dashboard shows structure but no extractable numbers**
- Chart exists but values not visible
- Table present but cells empty
- **Action:** Mark "Insufficient data for analysis"

**3. Only title/labels visible, no actual data**
- Section headers without content
- Placeholder text
- **Action:** Skip or note "Awaiting data population"

**4. Image quality too poor to read numbers**
- Blurry screenshots
- Text too small to extract
- **Action:** Note "Data not readable—higher resolution image needed"

---

## Examples: Honesty Builds Trust

### ✅ Good: Honest About Limitations

**Dashboard:** "Agents - Habit Formation" page with no frequency distribution visible

**Insight:**
```json
{
  "headline": "Insufficient data for habit formation analysis",
  "insights": [
    "Dashboard page contains no usage frequency distribution data",
    "To analyze habit formation, need: daily/weekly/monthly usage breakdown or engagement tier distribution",
    "Recommend adding frequency tracking to dashboard for comprehensive adoption analysis"
  ],
  "numbers_used": []
}
```

**User reaction:** "Fair point. I'll add that data to the dashboard for next time."
**Trust level:** HIGH (analyst is rigorous and honest)

---

### ❌ Bad: Forced Generic Content

**Dashboard:** Same "Agents - Habit Formation" page with no data

**Insight:**
```json
{
  "headline": "Usage frequency patterns show opportunities for increased engagement",
  "insights": [
    "Different users demonstrate varying levels of engagement with the platform",
    "Opportunity exists to move users from light to heavy usage tiers",
    "Champion programs can help drive adoption across user segments"
  ],
  "numbers_used": []
}
```

**User reaction:** "This is completely generic. What frequency patterns? There's no data here!"
**Trust level:** LOW (analyst appears to make things up)
**Damage:** User now questions ALL previous insights, even the strong ones

---

## Partial Data Scenarios

### When Some Data Exists

**Approach:** Generate insights ONLY for available data, acknowledge gaps

**Example:**
```json
{
  "headline": "92 active agent users tracked (frequency data pending)",
  "insights": [
    "Agent user base of 92 provides good foundation for adoption analysis",
    "Unable to assess engagement patterns without frequency distribution—recommend tracking daily vs. weekly vs. monthly usage",
    "Once frequency data available, can segment power users for champion programs and identify at-risk users for retention"
  ],
  "numbers_used": ["92"]
}
```

**Why this works:**
- Uses available data (92 users)
- Honest about what's missing (frequency data)
- Explains impact of gap (can't segment users yet)
- Suggests path forward (add tracking)

---

## Implementation Guidelines

### Step 1: Analyze Dashboard Image
```
IF dashboard has extractable numbers:
    → Generate data-driven insights

IF dashboard has structure but no numbers:
    → Mark "Insufficient data" + explain what's needed

IF dashboard is blank/placeholder:
    → Skip slide entirely
```

### Step 2: Quality Check Each Insight
```
FOR EACH insight:
    IF contains specific number from dashboard:
        ✅ KEEP
    ELSE IF makes specific observation from visible pattern:
        ✅ KEEP (if clearly supported)
    ELSE IF generic statement without data support:
        ❌ REMOVE (damages credibility)
```

### Step 3: Validate Trust
```
Ask: "Would the user trust this insight?"

IF insight could apply to ANY dashboard (not specific to this one):
    → Too generic, remove

IF insight contradicts visible data or makes unsupported claims:
    → Remove immediately

IF insight based on assumption rather than observation:
    → Remove or mark as speculation
```

---

## Impact on Output Quality

### Scenario: Dashboard with 14 slides

**Old Approach (forced completeness):**
- Generate insights for all 14 slides
- 3 slides have no real data
- Generate generic filler for those 3
- **Result:** 14 slides, but user distrusts all insights

**New Approach (credibility-first):**
- Generate insights for 11 slides with data
- Mark 3 slides as "Insufficient data"
- Be explicit about what's missing
- **Result:** 11 trusted insights + 3 honest gaps = HIGH credibility

**User outcome:**
- Acts on all 11 strong insights
- Improves dashboard to fill 3 gaps
- Returns for next analysis with better data
- **Relationship built on trust**

---

## Communication Templates

### Template 1: No Data Available
```
Headline: "Insufficient data for analysis"
Insight: "This dashboard page contains no extractable metrics or visualizations. Analysis requires [specific data needed]."
```

### Template 2: Partial Data
```
Headline: "[Number from visible data] - [available insight] (additional metrics needed for full analysis)"
Insight 1: "[Insight from available data]"
Insight 2: "Unable to assess [specific analysis] without [missing data]"
Insight 3: "Recommend adding [specific metrics] to enable [specific analysis type]"
```

### Template 3: Data Quality Issue
```
Headline: "Data visibility insufficient for detailed analysis"
Insight: "Dashboard image resolution or visualization format prevents extracting specific numbers. Recommend: [higher resolution export / data table view / specific metric cards]"
```

---

## Data Accuracy Rules (CRITICAL)

### Rule 1: Only Mention What's Visible
**Every entity mentioned in insights MUST be visible on the dashboard page.**

❌ **WRONG:** "Teams integration shows 3x higher engagement than Outlook"
- If "Teams" and "Outlook" are not visible as labels/categories on this specific page

✅ **RIGHT:** Only mention platforms, teams, functions, departments, or any nouns that are explicitly shown in the visual

**Examples of entities that must be visible:**
- Platform names (Teams, Outlook, Excel, PowerPoint, etc.)
- Department names (Finance, HR, IT, Sales, etc.)
- Feature names (Chat, Agents, Copilot, etc.)
- Team names, location names, product names

**Test:** Can I point to the exact spot on the dashboard where this name appears?
- **YES** → Safe to mention in insight
- **NO** → Do NOT mention, even if you think it's likely

### Rule 2: Match Number Units Exactly
**Numbers and units MUST match exactly what's shown in the visual.**

❌ **WRONG Examples:**
- Dashboard shows "13" → Insight says "13K actions" (incorrect unit)
- Dashboard shows "13K" → Insight says "13,000 actions" (unit mismatch)
- Dashboard shows "13M" → Insight says "13 million" (format mismatch)
- Dashboard shows "13.5K" → Insight says "13K" (precision lost)

✅ **RIGHT Examples:**
- Dashboard shows "13" → Insight says "13 actions"
- Dashboard shows "13K" → Insight says "13K actions"
- Dashboard shows "13M" → Insight says "13M actions"
- Dashboard shows "13.5K" → Insight says "13.5K actions"

**Critical:** Mismatched units destroy credibility instantly. Users will catch this.

### Rule 3: Numbers Must Be Traceable
**Every number in an insight must be directly attributable to something visible on the page.**

**Test for each number:**
1. Can I point to where this exact number appears on the dashboard?
2. Is the unit (K, M, %, etc.) exactly as shown?
3. If it's a calculation (e.g., percentage, ratio), are both source numbers visible?

**If answer is NO to any question → Do NOT use that number**

---

## Key Principles

### 1. One Generic Insight Undermines Ten Strong Ones
User discovers one unsupported claim → questions everything else

### 2. Honesty About Gaps Builds Expertise
"I need X data to analyze Y" → demonstrates analytical rigor

### 3. Skip Rather Than Speculate
Better to deliver 8/10 slides excellently than 10/10 with 2 weak

### 4. Incomplete Dashboard ≠ Failed Analysis
Missing data is feedback for dashboard improvement, not analysis failure

### 5. Trust Is Fragile, Hard to Rebuild
Lose credibility once → very difficult to regain

### 6. Every Entity Must Be Visible
Never mention platforms, teams, or any nouns not shown on the page

### 7. Units Must Match Exactly
13 ≠ 13K ≠ 13M — use exactly what the visual shows

---

## Constitutional Mandate

From **Claude PowerPoint Constitution Section 5A & Section 6.6:**

> "Claude MUST prioritize credibility over producing content for every slide."

> "If the source is insufficient to support a slide, Claude MUST state that limitation explicitly."

> "NEVER generate vague or generic content to fill the space."

> "Credibility over completeness: Skip slides rather than compromise trust."

---

## Success Metrics

**Old metric:** Did we produce insights for all slides?
**New metric:** Does the user trust and act on our insights?

**Result:** Better to deliver:
- 8 slides with 100% user trust and action
- Than 10 slides with 60% trust and skepticism

**Long-term outcome:**
- User improves dashboards based on honest feedback
- Next analysis has better data
- Trust compounds over time
- Analyst becomes go-to source of truth

---

## Summary

✅ **DO:**
- Generate insights only when data supports them
- Mark "Insufficient data" when numbers missing
- Explain what data would enable better analysis
- Skip slides rather than force generic content

❌ **DON'T:**
- Generate vague insights to fill space
- Make assumptions about missing data
- Create generic statements without specific numbers
- Sacrifice credibility for completeness

**Remember:** Your credibility is your most valuable asset. Protect it by being honest about data limitations.
