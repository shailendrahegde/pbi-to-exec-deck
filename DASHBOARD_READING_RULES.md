# Dashboard Reading & Validation Rules

## Critical Rules for Accurate Data Extraction

### Rule 1: Match Numbers to EXACT Labels
**Problem:** Misattributing values to wrong entities (e.g., Team C's 73,071 attributed to Team B)

**Rule:**
- Read table row labels character-by-character
- Match each number to its corresponding row/column label
- Never guess or assume label mappings
- If table has: `Team C | 73,071`, then state "Team C: 73,071" (not Team B)

**Validation:**
- ✅ Each number citation includes its exact label
- ❌ Generic references like "the first team" or "top row"

---

### Rule 2: Never Assume First Row is Aggregate
**Problem:** Assuming first row is "Total" or "Organization" when it could be a specific division

**Rule:**
- Read the actual label in first row
- Could be: "Team B", "Division A", "Total", "Organization (Aggregated)", or any entity
- Only call it "aggregate" if label explicitly says so
- State what the label actually says: "Team B shows 207 users" not "Organization shows..."

**Validation:**
- ✅ Use exact label text from dashboard
- ❌ Assumptions about row meaning without reading label

---

### Rule 3: Read Chart Axes Precisely
**Problem:** Misreading trend values (e.g., "1.64M declining to 0.51K" when actual values differ)

**Rule:**
- Read axis scale carefully (is it 0-1K, 0-100K, 0-1M?)
- Trace data points to exact axis values
- Don't interpolate or guess - if unsure, describe the trend qualitatively
- Verify units match (K vs M vs raw numbers)

**Validation:**
- ✅ Can point to exact position on axis where value appears
- ✅ Units consistent with axis labels
- ❌ Numbers that don't align with visible axis markings

---

### Rule 4: Match User Categories to Exact Labels
**Problem:** Saying "148 Bottom 25%" when label actually says "148 Unlicensed"

**Rule:**
- User categories have specific names: "Unlicensed", "Bottom 25%", "Top 10%", etc.
- Match the number to the exact category label shown
- Don't substitute similar-sounding categories

**Validation:**
- ✅ "148 Unlicensed users" (if label says "Unlicensed")
- ❌ "148 Bottom 25%" (when label says "Unlicensed")

---

### Rule 5: Identify Selected Filters (Highlighted Elements)
**Problem:** Missing that "External" is highlighted (black) while "Networks", "Meetings" (grey) are not selected

**Rule:**
- **Visual indicators of selection:**
  - Black/dark = Selected/Active filter
  - Grey/light = Not selected/Inactive
  - Outlined/bordered = Selected
  - Filled vs unfilled = Active vs inactive
- Chart below filter shows data for ONLY the selected filter
- State the filter context: "External collaboration (selected filter) shows..."

**Validation:**
- ✅ Mention which filter is active when describing chart
- ✅ "External collaboration chart shows..." (when External is highlighted)
- ❌ Describing chart without noting what filter is selected

---

### Rule 6: Verify Column Headers in Tables
**Problem:** Misreading columns (mentioned in feedback for Page 4)

**Rule:**
- Read column headers left to right
- Match each column header to its data values below
- Don't skip columns or assume column order
- If table is: `Name | Value | Rank`, read all three columns

**Validation:**
- ✅ Each data point references correct column
- ❌ Mixing up which column a value belongs to

---

### Rule 7: Unit Consistency Check
**Problem:** Mixing units (K, M, raw numbers) inconsistently

**Rule:**
- Use units EXACTLY as shown in dashboard
- If dashboard shows "13K", write "13K" (not "13,000" or "13")
- If dashboard shows "1.64M", write "1.64M" (not "1,640K")
- If dashboard shows "500", write "500" (not "0.5K")

**Validation:**
- ✅ Unit matches what's displayed on screen
- ❌ Converting units (unless explicitly needed for comparison)

---

### Rule 8: Don't Invent Numbers
**Problem:** Stating numbers that don't appear on the dashboard

**Rule:**
- Only cite numbers you can see and point to
- If doing math (e.g., "2.6x difference"), show the calculation
- If number is unclear/pixelated, describe qualitatively instead
- Never fabricate specific values

**Validation:**
- ✅ Every number appears somewhere visible on dashboard
- ❌ Calculated or inferred numbers without showing work

---

### Rule 9: Read Legends for Chart Context
**Problem:** Describing chart without understanding what colors/segments represent

**Rule:**
- Read legend before describing chart
- Match colors/patterns to legend labels
- State what each visual element represents
- Don't guess what chart segments mean

**Validation:**
- ✅ "Blue bars represent meetings, grey bars represent emails (per legend)"
- ❌ "These bars show activity" (without reading legend)

---

### Rule 10: Context Before Numbers
**Problem:** Numbers without context are meaningless

**Rule:**
- State WHAT the number represents before the value
- Format: "[Label]: [Value] [Unit]"
- Example: "Team B: 207 users" not "207 for team B"
- Include full context: "Team B shows 207 active users generating 42 assisted hours"

**Validation:**
- ✅ Clear subject + verb + object structure
- ❌ Numbers floating without clear referent

---

### Rule 11: Executive Summary Must Synthesize All Pages
**Problem:** Summary only reflects first few pages or misses key insights from later pages

**Rule:**
- Read and analyze ALL dashboard pages before writing summary
- Identify the 5 most business-critical findings across entire dashboard
- Prioritize by impact, not by page order
- Each bullet must include specific numbers from the data
- Cross-reference: Do these 5 points cover the most important themes?

**Validation:**
- ✅ Summary references insights from multiple pages (not just page 1-2)
- ✅ Each bullet has specific number from dashboard
- ❌ Generic statements without data backing

---

### Rule 12: Recommendations Must Be Actionable and Grounded
**Problem:** Vague recommendations like "improve training" or "increase adoption" without specifics

**Rule:**
- Each recommendation must start with action verb (Pilot, Expand, Target, Launch, etc.)
- Include WHO should do it (which team, department, or user group)
- Include WHAT specifically (which feature, workflow, or capability)
- Include expected outcome or success metric
- Trace recommendation back to specific insight from dashboard

**Examples:**
- ✅ "Pilot agent training with HR Generalists (140 actions/user) to establish replicable best practices for departments below 50 actions/user baseline"
- ❌ "Improve agent adoption" (too vague, no specifics)
- ✅ "Expand PowerPoint Copilot access to Finance team (217 prompts/user pattern) to replicate proven high-engagement workflow"
- ❌ "Get more users" (no grounding in data)

**Validation:**
- ✅ Action verb + specific target + expected outcome
- ✅ Can point to dashboard data supporting this recommendation
- ❌ Generic advice that could apply to any dashboard

---

### Rule 13: Handle Strip Edges — Never Invent Missing Chart Data

**Context:** Portrait PDF pages are split into horizontal strips at whitespace gaps between sections. The splitter snaps to background-coloured rows, so charts should not be bisected. However, if a chart title appears at the very bottom of a strip without its chart body (or a chart body appears at the top of the next strip without its title), treat those two strips as a pair for that chart.

**Rule:**
- If content is visually truncated at the top or bottom edge of a strip, acknowledge that you are seeing a partial section and note "continues from previous strip" or "continues on next strip"
- Do NOT invent numbers for the unseen portion
- Do NOT analyse a chart whose axis labels, bars, or legend are cut off — mark it as partial and skip numeric claims for that chart
- When both strips of a pair are available (they always are in this report), cross-reference them before citing numbers from any chart that spans the boundary

**Validation:**
- ✅ "Chart is partially visible at strip edge — citing only fully visible data points"
- ✅ "Legend for this chart appears on the next strip — cross-referenced before citing numbers"
- ❌ Guessing or extrapolating values from a cropped chart axis
- ❌ Describing a chart as complete when its labels/legend are cut off

---

## Pre-Analysis Checklist

Before generating insights for a dashboard page:

1. [ ] Identify all filter selections (black/highlighted vs grey)
2. [ ] Read all table row/column labels
3. [ ] Check axis scales on charts
4. [ ] Read legend for color/pattern meanings
5. [ ] Verify units are consistent (K, M, %, etc.)
6. [ ] Match each number to its exact label
7. [ ] Don't assume first row = total without reading label

---

## Example: Correct Reading

**Dashboard shows:**
```
Team B    | 207 users | 42 hours
Team C    | 24 users  | 73,071 value
```

**Filter: "External" (black), "Internal" (grey)**

**Correct insight:**
"Team B leads with 207 users averaging 42 assisted hours. Team C shows 24 users generating 73,071 in assisted value. External collaboration (selected filter) chart indicates..."

**Incorrect insight:**
"Organization shows 207 users (first row assumed to be total) with Team B at 73,071 (misread table)..."

---

## Validation Questions for Each Number

Before stating a number in an insight, ask:

1. **Can I point to where this number appears?** (Yes/No)
2. **What is the exact label next to this number?** (Read it)
3. **What are the units?** (K, M, %, raw number?)
4. **Is this from selected or unselected filter?** (Check highlight)
5. **Am I reading the correct axis/column?** (Verify)

If you answer "No" or "Unsure" to any question, don't use that number.

---

## Title Slide Formatting Rule

**Rule:** Title and subtitle must be left-aligned at the same horizontal position

**Implementation:**
- Both should start at same `left` coordinate
- Title: Larger font, same X position
- Subtitle: Smaller font, same X position (below title)
- No indentation or offset between them

**Validation:**
- ✅ Both text boxes have identical `left` property
- ❌ Subtitle indented or offset from title
