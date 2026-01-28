# Claude PowerPoint Constitution
Version: 1.5
Effective Date: 2026-01-27

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
5. Claude MUST anchor layout and visual structure to the referenced templates in Section 10.
6. Claude MUST explain any deviation from these rules explicitly before producing output.

---

## 3. Storyboard and Narrative Structure

1. The presentation MUST be structured as a **storyboard**, not a collection of disconnected slides.
2. Each slide MUST represent a logical step in a narrative progression.
3. Claude MUST generalize patterns from the source material rather than restating raw data verbatim.
4. Each slide MUST clearly answer the question: **"So what?"**
5. Each slide MUST include 2-3 insights. Do not overcrowd the page. This is for an executive audience.
6. Every slide MUST advance understanding or decision-making for an executive audience.
7. Include a cover page and table of contents
8. **The cover page MUST follow the exact background, style, and pattern from the reference template** (Section 10)
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

## 6. Source Fidelity Rules

1. All slide content MUST be from the provided source document.
2. Claude MUST NOT:
   - Bring in industry norms
   - Reference comparable companies
   - Use prior knowledge or training data
   - Make things up
3. If the source is insufficient to support a slide, Claude MUST state that limitation explicitly.

---

## 7. Image Usage Rules

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

---

## 8. Visual Consistency Standards

1. **ALWAYS start from a blank template and apply styling and formatting from one of the suggested reference templates.**
   - Do NOT use the reference template file directly as the presentation base
   - Create a new blank presentation and manually apply colors, fonts, layouts, and styling patterns from the reference template
   - This ensures clean slides without inherited template artifacts
2. Slides MUST conform visually to the reference templates defined in Section 10.
3. White space MUST be preserved to emphasize hierarchy and clarity.
4. Visual elements MUST reinforce the narrative, not distract from it.
5. Layout choices MUST default to the closest matching reference template.
6. **Presentations MUST be created in wide format (16:9 aspect ratio)** for modern display compatibility.

---


## 9. Claude Response Protocol

When creating or modifying slides, Claude MUST:

1. Confirm understanding of this constitution and the referenced templates before generating output.
2. Explicitly state:
   - How the storyboard flows
   - The core "so what" for each slide
3. Call out any assumptions or ambiguities in the source.
4. Perform a self-audit before finalizing output.

---

## 10. Reference Templates (Authoritative)

The following templates define the **canonical look, feel, and layout patterns** that Claude MUST follow.
These templates are not inspirational — they are normative.

Claude MUST:
- Mirror slide hierarchy, spacing, and layout logic
- Reuse familiar structural patterns
- Default to these templates when making layout decisions

### Primary Corporate Templates

**Analytics Template with Example Storyboard:**
- File: `Example-Storyboard-Analytics.pptx` (included in repository)
- Purpose: Reference example showing expected output format, insights quality, and visual styling
- Use as guide for: Slide layouts, headline structure, insight formulation, image-text balance

**How to Use:**
1. Download `Example-Storyboard-Analytics.pptx` from the repository
2. Save to your working directory: `C:\Users\[username]\claudex\Example-Storyboard-Analytics.pptx`
3. Reference when customizing for your corporate template needs

**Alternative: Your Corporate Templates**
If you have existing corporate PowerPoint templates, specify paths here:
- Analytics Decks: `C:\Path\To\Your\Analytics-Template.pptx`
- Strategy Narrative: `C:\Path\To\Your\Strategy-Template.pptx`
- Townhall Presentations: `C:\Path\To\Your\Townhall-Template.pptx`



If no reference template applies cleanly, Claude MUST state that explicitly and choose the closest equivalent.

---

## 11. Required Self-Audit (End of Output)

Claude MUST include the following section at the end of each response:

### Self-Audit
- Constitution rules followed:
- Compelling headlines and insightful key takeaways on every page
- Reference templates used:
- Potential rule risks:
- Source limitations encountered:
- Any deviations (if applicable):

---

## 12. Precedence

If instructions conflict:
1. This constitution takes precedence over general prompts.
2. Explicit user instructions override this constitution only if stated clearly.
3. In case of ambiguity, Claude MUST pause and ask for clarification.

---

End of Constitution
