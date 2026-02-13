# Project Structure

Clean, minimal structure for converting Power BI dashboards to executive presentations.

---

## Core Files (Required)

### User-Facing
- **`README.md`** (5KB) - Quick start guide for users
- **`LICENSE`** (1KB) - MIT license

### Conversion Script
- **`convert_dashboard_claude.py`** (8KB) - Main converter (3-step workflow)

### Templates & Standards
- **`Example-Storyboard-Analytics.pptx`** (1MB) - Visual template reference
- **`Claude PowerPoint Constitution.md`** (19KB) - Quality standards and rules

### Claude Instructions
- **`CLAUDE.md`** (12KB) - Detailed workflow for Claude Code

---

## Documentation (Optional but Recommended)

### Verification & Examples
- **`CONSTITUTION_COMPLIANCE.md`** (8KB) - Proves Constitution compliance
- **`PLATFORM_INSIGHTS_EXAMPLES.md`** (8KB) - Before/after insight examples
- **`CREDIBILITY_PRINCIPLE.md`** (8KB) - Philosophy on data honesty

**Purpose:** Help users understand quality standards and see examples of good insights

**Can be removed:** Yes, if you want absolute minimum. Core functionality works without these.

---

## Code Libraries

### `lib/` directory
```
lib/
├── __init__.py
├── extraction/
│   ├── __init__.py
│   └── extractor.py          # Extracts dashboards and numbers
├── analysis/
│   ├── __init__.py
│   ├── gaps.py               # Gap detection (adoption, habit, license)
│   ├── implications.py       # Business implication mapping
│   └── insights.py           # Insight composition
├── templates/
│   ├── __init__.py
│   ├── patterns.py           # Headline and insight patterns
│   └── classifiers.py        # Slide type detection
└── rendering/
    ├── __init__.py
    ├── builder.py            # PowerPoint slide builder (16:9)
    └── validator.py          # Constitution compliance checker
```

**Purpose:** Modular Python code for the converter
**Required:** Yes

---

## Test Suite

### `tests/` directory
```
tests/
├── __init__.py
├── test_extraction.py        # Tests data extraction
├── test_insights.py          # Tests insight generation
└── test_validation.py        # Tests Constitution compliance
```

**Purpose:** Unit tests for developers
**Required:** No (useful for development, not needed for users)

---

## Working Directory

### `temp/` directory
- Auto-generated during conversion
- Contains: extracted dashboard images, analysis JSON, insights JSON
- **`.gitignored`** - not committed to repo
- Created automatically when converter runs

---

## What Was Removed

### Old Converters (Obsolete)
- ❌ `convert_dashboard.py` - Original version
- ❌ `convert_dashboard_v2.py` - OCR-based version
- ❌ `convert_dashboard_fast.py` - Cached metrics version
- ❌ `convert_dashboard_smart.py` - Rule-based version
- ❌ `create_exec_deck.py` - Early prototype
- ❌ `create_exec_deck_v2.py` - Early prototype

**Why removed:** Replaced by `convert_dashboard_claude.py`

### Old Documentation (Redundant)
- ❌ `CONSISTENCY-GUIDE.md` - Superseded by Constitution
- ❌ `EXAMPLE-SETUP.md` - Obsolete setup instructions
- ❌ `EXAMPLE-STORYBOARD.md` - Redundant with template
- ❌ `INSTALL_TESSERACT.md` - OCR no longer needed
- ❌ `QUICK-START.md` - Merged into README
- ❌ `SMART_CONVERTER_SUMMARY.md` - About obsolete approach

**Why removed:** Information now in README, CLAUDE.md, or Constitution

### Output/Temp Files (Generated)
- ❌ `analytics_template.txt` - Temp extraction file
- ❌ `dashboard_analysis_ocr.json` - Old OCR output
- ❌ `dashboard_analysis_v2.json` - Old analysis output
- ❌ `wpp22_metrics.json` - Example metrics
- ❌ `extracted_text_debug.txt` - Debug output
- ❌ `output_preview.txt` - Preview file
- ❌ Various `.txt` debug files
- ❌ `temp_slide_2.png` - Test image
- ❌ `test_number_extraction.py` - Test script

**Why removed:** Generated during conversion, shouldn't be committed

### Cleaned Directories
- ❌ `extracted_images/` - Old output directory (deleted)
- ✅ `temp/` - Now empty, auto-populated during conversion

---

## Minimal Installation

**What users need to clone:**
```
pbi-to-exec-deck/
├── README.md                               # Start here
├── convert_dashboard_claude.py             # Run this
├── CLAUDE.md                               # Claude reads this
├── Claude PowerPoint Constitution.md       # Quality standards
├── Example-Storyboard-Analytics.pptx       # Template reference
└── lib/                                    # Python modules
```

**Total size:** ~1.5 MB

**Optional additions:**
- Documentation files (CREDIBILITY_PRINCIPLE.md, etc.) - Add 24KB
- Test suite (`tests/`) - Add 50KB

---

## Quick Start Checklist

✅ Clone repository
✅ Ensure files exist:
   - `convert_dashboard_claude.py`
   - `Example-Storyboard-Analytics.pptx`
   - `CLAUDE.md`
   - `lib/` directory
✅ Export Power BI dashboard to .pptx
✅ Run: `python convert_dashboard_claude.py --source dashboard.pptx --prepare`
✅ Ask Claude to generate insights
✅ Run: `python convert_dashboard_claude.py --build --output result.pptx`
✅ Done!

---

## File Size Summary

**Core files:** 1.5 MB
- Python scripts: 8 KB
- Documentation: 40 KB
- Template: 1 MB
- Libraries: 50 KB

**Optional docs:** 24 KB
**Tests:** 50 KB

**Total with everything:** ~1.6 MB

---

## Maintenance

**Never commit:**
- `temp/*` (working files)
- `*_executive*.pptx` (generated outputs)
- `*.txt` debug files
- Test output files

**Always keep:**
- Core 8 files listed above
- `lib/` directory
- `.gitignore`

Clean, focused structure for production use.
