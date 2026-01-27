# Power BI to Executive Deck

Transform Power BI dashboards into executive-ready PowerPoint presentations using Claude Code and a governance framework (Claude PowerPoint Constitution) that ensures consistent formatting, compelling insights, and source fidelity.

## Purpose

This constitution ensures Claude Code produces high-quality, executive-ready analytics decks by enforcing:

- **Narrative rigor**: Storyboard-driven structure with clear "so what" messaging
- **Source fidelity**: All insights traceable to source data (no external facts or assumptions)
- **Visual consistency**: Professional formatting aligned with corporate templates
- **Actionability**: Insight-driven headlines emphasizing implications over description
- **Executive focus**: 2-3 insights per slide, avoiding information overload

The result: Transform raw dashboard exports into polished executive presentations that drive decision-making.

## How It Works

Claude Code reads the constitution file and follows its rules when generating presentations. The constitution defines:
- Slide structure and narrative flow
- Headline standards (insight-driven, not descriptive)
- Content constraints (actionable, source-grounded)
- Image handling (aspect ratios, borders, placement)
- Visual consistency (colors, fonts, layouts)

## Getting Started

### Prerequisites

- Claude Code CLI installed and configured
- Power BI report access
- Local working directory that Claude can access (e.g., `C:\Users\[username]\pbi-to-exec-deck\`)

### Step-by-Step Guide

#### Step 1: Publish Your Power BI Report

1. Open your Power BI report
2. Navigate to **File → Export → PowerPoint**
3. Save the exported `.pptx` file to your local working directory
   - Example: `C:\Users\username\pbi-to-exec-deck\Dashboard-Report.pptx`

#### Step 2: Download the Constitution File

1. Download `Claude PowerPoint Constitution.md` from this repository
2. Save it to the same working directory as your Power BI export
   - Example: `C:\Users\username\pbi-to-exec-deck\Claude PowerPoint Constitution.md`

#### Step 3: Customize the Constitution (Optional)

If you want to use your own corporate template and styling:

1. Open `Claude PowerPoint Constitution.md` in a text editor
2. Navigate to **Section 8: Visual Consistency Standards**
3. Modify the color palette and template references:

```markdown
## 8. Visual Consistency Standards

1. **ALWAYS start from a blank template and apply styling and formatting from one of the suggested reference templates.**
   - Do NOT use the reference template file directly as the presentation base
   - Create a new blank presentation and manually apply colors, fonts, layouts, and styling patterns from the reference template
   - This ensures clean slides without inherited template artifacts
```

4. Update **Section 10: Reference Templates** with your template file paths:

```markdown
## 10. Reference Templates (Authoritative)

### Primary Corporate Templates
- Analytics Decks:
  "C:\Path\To\Your\Analytics-Template.pptx"
- Strategy Narrative Template:
  "C:\Path\To\Your\Strategy-Template.pptx"
```

5. If using custom colors, document your corporate color palette:

```markdown
## Corporate Color Palette (Update with your brand colors)
- Primary Dark: RGBColor(0, 32, 96)
- Accent Blue: RGBColor(0, 114, 198)
- Background Beige: RGBColor(245, 239, 231)
```

#### Step 4: Run the Prompt in Claude Code

Open Claude Code CLI and run the following prompt:

```
"[Path-to-your-dashboard.pptx]" Read every page of this report and extract the most compelling insight a CIO or ITDM could glean from the data points available. Stick to the info within the report and index on actionability bringing out the so what. Generalize the patterns. State this as a story board. For formatting rules,      
  refer to constitution.
```

**Example:**
```
"C:\Users\username\pbi-to-exec-deck\AI-in-One Dashboard - v10.1.pptx" Read every page of this report and extract the most compelling insight a CIO or ITDM could glean from the data points available. Stick to the info within the report and index on actionability bringing out the so what. Generalize the patterns. State this as a story board. For formatting rules,      
  refer to constitution.
```

#### Step 5: Review and Refine

1. Claude Code will generate an executive analytics deck (typically named `[Original-Name] - Executive Analytics.pptx`)
2. Review the output presentation
3. If adjustments needed, provide specific feedback to Claude Code or perform manually edit.
4. Iterate until the presentation meets your standards

## Constitution Sections Overview

| Section | Focus | Key Requirements |
|---------|-------|------------------|
| **1. Purpose** | Overall objectives | Executive-ready, template-driven, source-grounded |
| **2. Absolute Rules** | Non-negotiable standards | Follow structure exactly, source fidelity only |
| **3. Storyboard Structure** | Narrative flow | Logical progression, 2-3 insights per slide, "so what" clarity |
| **4. Headline Standards** | Slide titles | Insight-driven (not descriptive), actionable implications |
| **5. Content Constraints** | Messaging rules | Index on actionability, explicit implications |
| **6. Source Fidelity** | Data integrity | **All content traceable to source, no external knowledge** |
| **7. Image Usage** | Visual handling | Maintain aspect ratio, 70% max width, borders, placement |
| **8. Visual Consistency** | Design standards | Start from blank, apply template styling, wide format |
| **9. Response Protocol** | Output expectations | Confirm understanding, state flow, self-audit |
| **10. Reference Templates** | Style anchors | Corporate template file paths |
| **11. Self-Audit** | Quality check | Rules followed, templates used, limitations noted |
| **12. Precedence** | Conflict resolution | Constitution overrides general prompts |

## Key Constitution Principles

### Section 6: Source Fidelity (CRITICAL)

All slide content **MUST** be traceable directly to the source document. Claude Code will **NOT**:
- Introduce external facts or industry benchmarks
- Reference comparable companies or market data
- Use training data or prior knowledge
- Add speculative recommendations beyond source support

**Example:**
- ❌ "Industry average adoption rate is 15%" (external benchmark)
- ✓ "Dashboard shows 1801 active users out of 4,381 total (40% adoption)" (source data)

### Section 4: Compelling Headlines

Headlines must be insight-driven and emphasize implications:
- ❌ "Agent Usage Trends" (descriptive label)
- ✓ "37% of Agent users remain infrequent - usage has not embedded into workflows" (insight + implication)

### Section 3.5: Executive Focus

Each slide limited to **2-3 insights** to avoid overcrowding for executive audience.

## Example Output

**Input:** Power BI dashboard export (14 slides of raw visualizations)

**Output:** Executive analytics deck (16 slides) with:
- Compelling cover page and executive summary
- Section separators for narrative flow
- Data-driven insights with persuasive headlines
- Image-left, text-right layout with 2-3 insights per slide
- Actionable recommendations grounded in source data
- Professional formatting matching corporate templates

## Troubleshooting

### "Claude isn't following the template styling"

- Ensure the template file path in Section 10 is correct and accessible
- Verify the constitution file is in Claude's working directory
- Check that Section 8 specifies starting from blank and applying styling

### "Insights seem generic or lack specifics"

- Emphasize "Enforce section 6" in your prompt
- Request "Extract specific data points from dashboards"
- Ask Claude to "Read every page and extract the most compelling insights"

### "Too much text on slides"

- Remind Claude of Section 3.5: "2-3 insights per slide for executive audience"
- Request "Generalize patterns" rather than listing all details

### "Headlines are descriptive, not insight-driven"

- Emphasize Section 4 in your prompt: "compelling insights and persuasive headlines"
- Provide example: "Not 'Usage Trends' but 'Usage patterns reveal X requiring action Y'"

## Contributing

Improvements to the constitution are welcome! Consider:
- Additional template styles for different presentation types
- Industry-specific customizations
- Enhanced quality control rules
- Better examples of compelling headlines

## License

This constitution is provided as-is for use with Claude Code. Modify freely for your organization's needs.

## Version History

- **v1.4 (Current)**: Added blank template requirement, 2-3 insights per slide, image-left/text-right layout
- **v1.3**: Image aspect ratio and placement rules
- **v1.2**: Executive quality standards
- **v1.1**: Initial release

## Support

For issues or questions:
- Review the constitution file sections
- Check troubleshooting guide above
- Ensure file paths are correct and accessible to Claude Code
- Verify Claude Code CLI is properly configured

---

**Quick Start Command:**

```bash
cd C:\Users\[your-username]\pbi-to-exec-deck
# Place your Power BI export and constitution.md here
# Then run in Claude Code:
"[your-dashboard.pptx]" Read every page of this report and extract the most compelling insight a CIO or ITDM could glean from the data points available. Stick to the info within the report and index on actionability bringing out the so what. Generalize the patterns. State this as a story board. For formatting rules,      
  refer to constitution.
```


