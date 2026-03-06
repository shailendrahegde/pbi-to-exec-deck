# Claude PowerPoint Constitution
Version: 1.7
Effective Date: 2026-02-13

## 1. Purpose

This constitution defines mandatory rules that Claude Code MUST follow when creating or modifying PowerPoint presentations.
The objective is to produce executive-ready slides that follow a familiar corporate template, maintain narrative rigor, and are strictly grounded in the provided source material.

Failure to comply with these rules is considered an error.

---

## 2. Absolute Rules (Non-Negotiable)

1. Claude MUST follow the slide structure, formatting, and template rules defined in this document exactly.
2. Claude MUST use ONLY the information provided in the supplied source document.
3. Claude MUST NOT introduce external facts, benchmarks, assumptions, or interpretations not explicitly supported by the source.
4. Claude MUST NOT insert stock images or decorative imagery.
5. Claude MUST anchor layout and visual structure to the referenced templates in Section 11.
6. **Claude MUST ensure every number, organizational attribute, noun, and acronym is traceable to the source document.**
7. **Claude MUST NOT include any content (text, numbers, examples) from reference templates in the final output - only styling and layout patterns.**
8. **Claude MUST create properly formatted presentations with all elements within page boundaries - poor formatting is grounds for rejection.**
9. Claude MUST explain any deviation from these rules explicitly before producing output.

---

## 3. Storyboard and Narrative Structure

1. The presentation MUST be structured as a **storyboard**, not a collection of disconnected slides.
2. Each slide MUST represent a logical step in a narrative progression.
3. Claude MUST generalize patterns from the source material rather than restating raw data verbatim.
4. Each slide MUST clearly answer the question: **"So what?"**
5. Each slide MUST include 2-3 insights. Do not overcrowd the page. This is for an executive audience.
6. Every slide MUST advance understanding or decision-making for an executive audience.
7. Include a cover page and table of contents
8. **The cover page MUST follow the exact background, style, and pattern from the reference template** (Section 11)
9. Insert section breaks when you are moving from one core idea to another
10. **Section separator pages MUST match the style and pattern from the reference template**

---

## 4. Slide Headline Standards

1. Every slide MUST contain a single compelling headline at the top.
2. Headlines MUST:
   - Be insight-driven, not descriptive
   - Emphasize actionability and implications
   - Avoid generic titles (e.g., "Overview", "Data Summary")
3. Headline formatting MUST be:
   - Font: Segoe Sans Display
   - Size: 24 pt
   - Color: Blue
4. Headlines MUST be written as complete thoughts, not labels.
5. Headlines MUST start at the same vertical and horizontal position. When you move from one slide to another, there shouldn't be a visible jump.

---

## 5. Content Constraints

1. Claude MUST index heavily on **actionability** based on the information it is seeing.
2. Each slide MUST make explicit:
   - What the insight is in the form of takeaways
   - Why it matters
   - What action or implication follows
3. Claude MUST NOT:
   - Add speculative recommendations
   - Infer intent beyond what the source supports
   - Normalize or reframe data using outside context

---

## 5A. Insight Generation Guidelines (For Dashboard Conversions)

When converting dashboards or analytics reports using Claude's vision capabilities, follow these insight generation principles:

### Tone and Style
1. **Be concise:** Keep insights to 1-2 sentences maximum
2. **Be friendly:** Frame as opportunities, not criticisms or failures
3. **Be specific:** Always include concrete numbers from the dashboard
4. **Be actionable:** Help IT decision makers take action

### What Makes a Compelling Insight
1. **Good insight structure:**
   - Specific number + observation + business implication
   - Example: "HR Generalists (140 actions/user) are 3-4x above average - great candidates to champion agent adoption"

2. **Bad patterns to avoid:**
   - Generic: "There are 1,275 active users"
   - Critical: "Only 11% adoption shows deployment failure"
   - Verbose: Multi-sentence explanations or academic analysis
   - Judgmental: "Critical blind spot", "massive gap", "total failure"

### Headline Formula
`[Specific Number] + [Key Insight/Opportunity]`

**Good examples:**
- "134 Agent users from 1,275 total (11%) - significant opportunity to expand automation adoption"
- "Client Finance achieves 217 prompts per user (5.7x average) - proven pattern to replicate"

**Bad examples:**
- "Agent Adoption Overview" (no number, descriptive not insight)
- "Only 11% of users have adopted agents showing critical deployment failure" (too critical/verbose)

### Three-Insight Structure Per Slide
1. **Insight 1 (Scale/Scope):** Specific observation with number
   - "Agent adoption at 11% while web chat reaches 1,275 users shows strong platform engagement"

2. **Insight 2 (Pattern/Opportunity):** Trend or comparison identified
   - "HR Generalists (140 actions/user) are 3-4x above average - natural champions for expansion"
   - **Include platform/feature patterns:** "Teams integration shows 3x higher engagement than Outlook - collaborative workflows drive value"

3. **Insight 3 (Action):** Recommendation with expected outcome
   - "Client Delivery's concentration positions them as ideal pilot group for scaling agent workflows"

### Platform & Feature Analysis (CRITICAL)
When analyzing dashboards, Claude MUST identify and analyze platform, app, and feature patterns:

**What to look for:**
- **Microsoft 365 apps:** Outlook, Teams, Word, Excel, PowerPoint, OneNote, OneDrive
- **Copilot features:** Chat (Web), Agents, M365 Copilot, Business Chat, integrations
- **Platform variations:** Web vs. desktop, standalone vs. embedded, integrated vs. separate
- **Feature adoption:** Which specific capabilities show high engagement
- **Usage concentration:** Which platforms dominate vs. underperform

**Examples of platform insights:**
- ✅ "PowerPoint Copilot leads with 450 actions while Excel shows 89 - content creation is primary use case to scale"
- ✅ "Teams integration achieves 3x engagement vs. Outlook - collaborative workflows deliver higher value"
- ✅ "Web chat dominates (1,275 users) while M365 integrated shows zero - integration barrier preventing adoption"
- ✅ "Draft as IP Agent (546 actions) outperforms other agents 7x - IP workflow automation resonates with users"

**Why this matters:**
- Reveals which workflows and integrations deliver value
- Identifies successful platforms to replicate patterns from
- Shows integration failures requiring attention
- Guides feature prioritization and training focus
- Helps target enablement to high-value use cases

### What NOT to Do
❌ **Don't criticize the analytics:** Avoid "missing data", "blind spot", "poor reporting"
❌ **Don't be verbose:** No multi-paragraph explanations
❌ **Don't use harsh language:** Avoid "failure", "disaster", "critical gap"
❌ **Don't make generic statements:** Always include specific numbers
❌ **Don't force insights without data:** If a page has no numbers or is blank, indicate "Insufficient data for analysis" rather than generating unsupported insights

### Credibility Over Completeness (CRITICAL)
**Claude MUST prioritize credibility over producing content for every slide.**

When a dashboard page has insufficient data:
- ✅ **DO:** State explicitly "Insufficient data for insight generation"
- ✅ **DO:** Explain what data would be needed: "This slide requires usage metrics to generate insights"
- ✅ **DO:** Skip the slide entirely rather than generate generic content
- ❌ **DON'T:** Generate vague insights to fill space
- ❌ **DON'T:** Make assumptions about missing data
- ❌ **DON'T:** Create generic statements without specific numbers

**Why this matters:**
- User trust depends on accuracy and honesty
- One unsupported insight undermines confidence in all insights
- Better to deliver 8 strong insights than 10 insights with 2 weak/generic ones
- Acknowledging data gaps shows rigor and builds credibility

**Examples:**

✅ **Good (honest about limitations):**
```
Headline: "Insufficient data for analysis"
Insight: "This dashboard page contains no extractable metrics. To generate insights, need: user counts, engagement metrics, or trend data."
```

❌ **Bad (forced generic content):**
```
Headline: "Platform shows potential for growth"
Insight: "Usage patterns indicate opportunities exist across the organization"
```

**Implementation:**
- If dashboard image shows no numbers → Skip or mark as "Insufficient data"
- If only partial data visible → Generate insights only for available data, note gaps
- If data present but unclear → Request clarification rather than guessing

### Friendly vs. Critical Tone Examples

| Critical (❌ Avoid) | Friendly (✅ Use) |
|---------------------|------------------|
| "Zero M365 Copilot adoption represents total deployment failure" | "M365 Copilot launch in progress - strong web chat base provides foundation" |
| "Only 11% agent adoption shows massive awareness gap" | "134 Agent users from 1,275 total (11%) - significant expansion opportunity" |
| "Critical blind spot preventing analysis" | "Opportunity to track engagement patterns for optimization" |
| "58% abandonment rate signals value realization failure" | "36 agents marked as keepers from 85 created - healthy curation in progress" |

---

## 6. Source Fidelity Rules

1. All slide content MUST be from the provided source document.
2. **Every number, organizational attribute (department, organization, function), noun, and acronym in the output MUST be traceable to the source file.**
   - If a specific department name appears in the output, it must exist in the source
   - If a metric or number is cited, it must come directly from the source data
   - No organizational details may be inferred or assumed
3. **Claude MUST use numbers from the source file (and source file only) to present insights, headlines, and recommendations.**
   - Insights MUST include specific numbers, percentages, or metrics from the source data
   - Headlines MUST reference concrete data points, not generic statements
   - Avoid vague language like "many users" - use "324 users" instead
   - Avoid generic statements like "usage is growing" - use "usage increased 45% month-over-month" instead
   - Every headline and insight should be data-driven and specific to the source material
4. **The reference template is for styling and layout reference ONLY.**
   - ZERO content from the reference template (text, numbers, examples, department names) may appear in the final output
   - Only visual styling, color schemes, and layout patterns should be borrowed from the reference template
   - All substantive content must come exclusively from the provided source document
5. Claude MUST NOT:
   - Bring in industry norms
   - Reference comparable companies
   - Use prior knowledge or training data
   - Make things up
   - Copy example content from reference templates
   - Generate insights without supporting data from the source
6. **If the source is insufficient to support a slide, Claude MUST state that limitation explicitly.**
   - Mark slide as "Insufficient data for analysis"
   - Explain what data would be needed to generate insights
   - NEVER generate vague or generic content to fill the space
   - Credibility over completeness: Skip slides rather than compromise trust

---
## 7. Text Boxes and Formatting Quality

### Layout & Page Boundaries
- Claude MUST ensure that Text boxes **always remain fully within the page borders**.
- Claude MUST ensure that Text boxes **do not extend into or overlap page margins**.
- Claude MUST ensure that Text boxes **do not overflow** (no content or box edges may be rendered outside the printable/page boundary).
- Claude MUST ensure that images **do not overlap with text boxes or headlines**.
- **⚠️ PENALTY WARNING:** Presentations with text boxes stretching beyond page margins, overlapping content, images overlapping with text boxes or headlines, or poorly formatted layouts will be considered defective output and rejected.

### Text Wrapping & Overflow Handling
- Claude MUST ensure that text inside text boxes **must wrap automatically** when it cannot fit on a single line.
- Claude MUST ensure that horizontal overflow (e.g., clipped text, scrolling, or text rendering past the box edge) **is not permitted**.
- Claude MUST ensure that if content cannot fit vertically, the text box must be resized **only within page borders/margins constraints**; otherwise content must be edited/shortened or moved to a continued element on the next page (still respecting these rules).

### Typography
- **Minimum font size:** 12 pt (must never be smaller).
- **Preferred font size:** 14 pt.
- **Preferred font family:** Segoe Sans Display.
- If Segoe Sans Display is unavailable, use the closest approved fallback per the constitution's typography rules, while still honoring the minimum font size requirement.

### Quality Standards
- All elements MUST be properly aligned and positioned within slide boundaries
- Text MUST be readable with no clipping or truncation
- Visual elements MUST NOT overlap unless intentionally designed for layering effect
- Spacing between elements MUST be consistent and professional

## 8. Image Usage Rules

1. Claude MUST NOT insert stock images under any circumstances.
2. Images MAY ONLY be used if they are explicitly provided or explicitly requested.
3. When inserting images, Claude MUST:
   - **ALWAYS maintain the original aspect ratio** - never stretch or distort images
   - Take up to 70% of the screen width for maximum visibility
   - **NEVER let images extend beyond the page width** - images must fit completely within slide boundaries
   - Place the image on a white canvas
   - Use a blue border (2pt width)
   - Apply rounded corners where appropriate
4. Images MUST support comprehension, not decoration.
5. Image positioning MUST allow adequate white space around the image for visual breathing room.
6. All images must be properly contained within the slide margins with no overflow or clipping.
7. When including images alongside narrative text, place image to the left, and the narrative to the right.
8. **Image Orientation — AUTO-CORRECTED (Mandatory):**
   - All source images MUST be presented in landscape orientation (width > height) inside the slide.
   - **EXIF rotation metadata** embedded in image files MUST be applied before any placement or dimension calculations.
   - If a source image is portrait (height > width) after EXIF correction, it MUST be rotated 90° clockwise to restore landscape orientation.  Power BI embeds landscape dashboards in portrait PDF pages using a CCW rotation; the clockwise correction reverses this artifact.
   - Text inside images MUST read left-to-right.  A correctly oriented landscape image from a Power BI dashboard will always have left-to-right text; portrait dimensions are a reliable signal of an incorrect rotation.
   - Claude MUST NOT paste a portrait-oriented dashboard image into a slide without first correcting its orientation.

---

## 9. Visual Consistency Standards

1. **ALWAYS start from a blank template and apply styling and formatting from one of the suggested reference templates.**
   - Do NOT use the reference template file directly as the presentation base
   - Create a new blank presentation and manually apply colors, fonts, layouts, and styling patterns from the reference template
   - This ensures clean slides without inherited template artifacts
2. Slides MUST conform visually to the reference templates defined in Section 10.
3. White space MUST be preserved to emphasize hierarchy and clarity.
4. Visual elements MUST reinforce the narrative, not distract from it.
5. Layout choices MUST default to the closest matching reference template.
6. All text must fit within the page boundaries. Nothing should overflow beyond the page width
7. **Presentations MUST be created in wide format (16:9 aspect ratio)** for modern display compatibility.

---


## 10. Claude Response Protocol

When creating or modifying slides, Claude MUST:

1. Confirm understanding of this constitution and the referenced templates before generating output.
2. Explicitly state:
   - How the storyboard flows
   - The core "so what" for each slide
3. Call out any assumptions or ambiguities in the source.
4. Perform a self-audit before finalizing output.

---

## 11. Reference Templates (Authoritative)

The following templates define the **canonical look, feel, and layout patterns** that Claude MUST follow.
These templates are not inspirational — they are normative.

Claude MUST:
- Mirror slide hierarchy, spacing, and layout logic
- Reuse familiar structural patterns
- Default to these templates when making layout decisions

### Primary Corporate Templates

**Analytics Template with Example Storyboard (Default):**
- File: `./Example-Storyboard-Analytics.pptx` (included in repository)
- Purpose: Reference example showing expected output format, insights quality, and visual styling
- Use as guide for: Slide layouts, headline structure, insight formulation, image-text balance
- **This is the default reference template and works out-of-the-box after cloning the repository**

**How to Use (No Setup Required):**
1. Clone the repository
2. The reference template is automatically available at `./Example-Storyboard-Analytics.pptx`
3. Claude will use this by default for creating presentations

**Optional: Customize with Your Corporate Templates**
If you want to use your own corporate PowerPoint templates instead, you can optionally specify custom paths:
- Analytics Decks: `C:\Path\To\Your\Analytics-Template.pptx`
- Strategy Narrative: `C:\Path\To\Your\Strategy-Template.pptx`
- Townhall Presentations: `C:\Path\To\Your\Townhall-Template.pptx`

*Note: Customization is optional. The default template works immediately after cloning.*

If no reference template applies cleanly, Claude MUST state that explicitly and choose the closest equivalent.

---

## 12. Required Self-Audit (End of Output)

Claude MUST include the following section at the end of each response:

### Self-Audit
- Constitution rules followed:
- Compelling headlines and insightful key takeaways on every page:
- Reference templates used (styling only, no content borrowed):
- Source traceability verified (all numbers, departments, nouns, acronyms from source):
- Formatting quality verified (no text boxes beyond margins, no overlapping elements):
- Potential rule risks:
- Source limitations encountered:
- Any deviations (if applicable):

---

## 13. Precedence

If instructions conflict:
1. This constitution takes precedence over general prompts.
2. Explicit user instructions override this constitution only if stated clearly.
3. In case of ambiguity, Claude MUST pause and ask for clarification.

---

End of Constitution
