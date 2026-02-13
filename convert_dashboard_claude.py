#!/usr/bin/env python3
"""
Claude-Powered Dashboard Converter

Uses current Claude Code session to generate analyst-grade insights.
No API key needed - leverages the interactive Claude session.

Workflow:
1. Extract slides as images (deterministic)
2. Prepare analysis request for Claude
3. Claude analyzes each image and generates insights
4. Build final presentation with insights
"""

import argparse
import json
import sys
from pathlib import Path
from pptx import Presentation
from PIL import Image
import io


def extract_slide_as_image(prs, slide_idx, output_path):
    """Extract a slide's dashboard image"""
    slide = prs.slides[slide_idx]

    # Find the main dashboard image
    for shape in slide.shapes:
        if shape.shape_type == 13:  # Picture
            image_stream = io.BytesIO(shape.image.blob)
            img = Image.open(image_stream)
            img.save(output_path)
            return True

    return False


def extract_slide_title(slide):
    """Extract slide title, removing emojis"""
    import re

    for shape in slide.shapes:
        if hasattr(shape, 'text') and shape.text.strip():
            title = shape.text.strip()
            # Remove emoji characters
            title = re.sub(r'[\U00010000-\U0010ffff]', '', title).strip()
            return title

    return "Untitled Slide"


def classify_slide_type(title):
    """Classify slide type for context"""
    title_lower = title.lower()

    if 'trend' in title_lower or 'over time' in title_lower:
        return 'trends'
    elif 'leaderboard' in title_lower or 'top' in title_lower:
        return 'leaderboard'
    elif 'health' in title_lower or 'overview' in title_lower:
        return 'health_check'
    elif 'habit' in title_lower or 'frequency' in title_lower:
        return 'habit_formation'
    elif 'license' in title_lower or 'priority' in title_lower:
        return 'license_priority'
    else:
        return 'general'


def prepare_for_claude_analysis(source_path):
    """
    Extract slides and prepare analysis request for Claude.

    Returns the analysis request structure.
    """
    print("=" * 70)
    print("PREPARING SLIDES FOR CLAUDE ANALYSIS")
    print("=" * 70)

    prs = Presentation(source_path)

    # Create temp directory for images
    Path('temp').mkdir(exist_ok=True)

    slides_to_analyze = []

    print(f"\nExtracting {len(prs.slides)} slides...")
    print("  Skipping slide 1 (title page)...")

    for idx, slide in enumerate(prs.slides):
        # Skip the first slide (title page with metadata, no insights)
        if idx == 0:
            continue

        title = extract_slide_title(slide)
        image_path = f"temp/slide_{idx+1}.png"

        # Extract dashboard image
        has_image = extract_slide_as_image(prs, idx, image_path)

        if has_image:
            slide_info = {
                'slide_number': idx + 1,
                'title': title,
                'image_path': image_path,
                'slide_type': classify_slide_type(title)
            }

            slides_to_analyze.append(slide_info)
            print(f"  OK Slide {idx+1}: {title[:50]}...")

    # Save analysis request
    request_file = 'temp/analysis_request.json'
    with open(request_file, 'w', encoding='utf-8') as f:
        json.dump({
            'source_file': source_path,
            'total_slides': len(slides_to_analyze),
            'slides': slides_to_analyze
        }, f, indent=2)

    print(f"\nOK Prepared {len(slides_to_analyze)} slides for analysis")
    print(f"OK Analysis request saved to: {request_file}")

    return request_file


def show_claude_instructions(request_file):
    """Show instructions for Claude to generate insights (manual mode)"""

    print("\n" + "=" * 70)
    print("NEXT STEP: CLAUDE GENERATES INSIGHTS")
    print("=" * 70)

    print(f"""
Claude will now analyze each dashboard image and generate insights.

Please say to Claude Code:

    "Analyze the dashboards in {request_file} and generate analyst-grade insights.
    Save results to temp/claude_insights.json"

After Claude completes, run:
    python convert_dashboard_claude.py --build --output [output.pptx]
""")


def trigger_claude_analysis(request_file):
    """Trigger Claude Code to analyze dashboards (automatic mode)"""

    print("\n" + "=" * 70)
    print("STEP 2: CLAUDE ANALYSIS")
    print("=" * 70)

    # Load request to show Claude what to analyze
    with open(request_file, 'r', encoding='utf-8') as f:
        request = json.load(f)

    print(f"\nClaude Code: Please analyze these {request['total_slides']} dashboard images.\n")

    print("For each slide:")
    for slide in request['slides']:
        print(f"  - Slide {slide['slide_number']}: {slide['title']}")
        print(f"    Image: {slide['image_path']}")
        print(f"    Type: {slide['slide_type']}")

    print("\n" + "-" * 70)
    print("CLAUDE CODE TASK:")
    print("-" * 70)
    print("""
Act as senior analyst advisor to IT decision maker.

For EACH dashboard image above:
1. Read the image file
2. Extract specific numbers visible in the dashboard
3. Generate compelling headline with specific number ([Number] + [Insight])
4. Generate 2-3 concise insights (1-2 sentences each):
   - Insight 1: Specific observation with number
   - Insight 2: Pattern/opportunity identified (include platform/feature analysis)
   - Insight 3: Actionable recommendation
5. Use friendly tone (opportunities, not criticisms)
6. If insufficient data visible, mark as "Insufficient data"

Follow Claude PowerPoint Constitution Section 5A guidelines.

Save results to: temp/claude_insights.json

Format:
{
  "slides": [
    {
      "title": "slide title",
      "headline": "[Number] + [Insight]",
      "insights": ["insight 1", "insight 2", "insight 3"],
      "numbers_used": ["123", "45%"]
    }
  ]
}
""")
    print("-" * 70)


def build_presentation_from_insights(source_path, output_path, insights_file):
    """Build final presentation using Claude's insights"""

    print("\n" + "=" * 70)
    print("BUILDING FINAL PRESENTATION")
    print("=" * 70)

    # Load Claude's insights
    with open(insights_file, 'r', encoding='utf-8') as f:
        insights_data = json.load(f)

    print(f"\nOK Loaded insights for {len(insights_data.get('slides', []))} slides")

    # Use existing smart converter's rendering engine
    from lib.rendering.builder import render_presentation
    from lib.analysis.insights import Insight

    # Convert Claude's insights to expected format
    insights = {}
    for slide_insight in insights_data.get('slides', []):
        title = slide_insight['title']
        insights[title] = Insight(
            headline=slide_insight['headline'],
            bullet_points=slide_insight['insights'],
            source_numbers=slide_insight.get('numbers_used', [])
        )

    # Render presentation
    print(f"\nRendering presentation...")
    render_presentation(source_path, insights, output_path)

    print(f"\nOK Created: {output_path}")

    # Validate
    from lib.rendering.validator import validate_output
    passed, report = validate_output(insights)
    print("\n" + report)

    print("\n" + "=" * 70)
    print("OK CONVERSION COMPLETE")
    print("=" * 70)


def generate_output_filename(source_path):
    """Generate output filename from source (e.g., dashboard.pptx -> dashboard_executive.pptx)"""
    from pathlib import Path
    source = Path(source_path)
    return str(source.parent / f"{source.stem}_executive{source.suffix}")


def main():
    parser = argparse.ArgumentParser(
        description='Convert Power BI dashboards using Claude Code for insights',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Single command (automatic workflow):
  python convert_dashboard_claude.py --source dashboard.pptx

  # With custom output name:
  python convert_dashboard_claude.py --source dashboard.pptx --output executive.pptx

  # Manual workflow (step-by-step):
  python convert_dashboard_claude.py --source dashboard.pptx --prepare
  python convert_dashboard_claude.py --build --output executive.pptx
"""
    )

    parser.add_argument('--source', help='Source PowerPoint file')
    parser.add_argument('--output', help='Output PowerPoint file (default: source_executive.pptx)')
    parser.add_argument('--prepare', action='store_true',
                       help='Prepare slides for Claude analysis (Step 1 only)')
    parser.add_argument('--build', action='store_true',
                       help='Build final presentation from Claude insights (Step 3 only)')
    parser.add_argument('--insights', default='temp/claude_insights.json',
                       help='Path to Claude insights JSON (default: temp/claude_insights.json)')

    args = parser.parse_args()

    # ========================================================================
    # SINGLE-COMMAND WORKFLOW: Orchestrate all 3 steps automatically
    # ========================================================================
    if args.source and not args.prepare and not args.build:
        # Auto-generate output filename if not provided
        output_path = args.output or generate_output_filename(args.source)
        print("\n" + "=" * 70)
        print("AUTOMATED CONVERSION WORKFLOW")
        print("=" * 70)
        print(f"\nSource: {args.source}")
        print(f"Output: {output_path}")
        print("\nThis will run all 3 steps automatically:")
        print("  1. Extract dashboard images")
        print("  2. Claude analyzes and generates insights")
        print("  3. Build executive presentation")

        # STEP 1: Prepare slides for analysis
        print("\n" + "=" * 70)
        print("STEP 1: EXTRACTING DASHBOARDS")
        print("=" * 70)
        request_file = prepare_for_claude_analysis(args.source)

        # STEP 2: Trigger Claude analysis
        trigger_claude_analysis(request_file)

        # Wait for user confirmation after Claude generates insights
        print("\n" + "=" * 70)
        input("Press ENTER after Claude has generated insights...")
        print("=" * 70)

        # STEP 3: Build final presentation
        print("\n" + "=" * 70)
        print("STEP 3: BUILDING PRESENTATION")
        print("=" * 70)
        build_presentation_from_insights(args.source, output_path, args.insights)

        print("\n" + "=" * 70)
        print("✓ CONVERSION COMPLETE!")
        print("=" * 70)
        print(f"\nExecutive presentation created: {output_path}")
        return 0

    # ========================================================================
    # MANUAL WORKFLOW: Individual steps
    # ========================================================================
    if args.prepare or (args.source and not args.build):
        # Step 1: Prepare slides for Claude
        if not args.source:
            print("Error: --source required")
            return 1

        request_file = prepare_for_claude_analysis(args.source)
        show_claude_instructions(request_file)

    elif args.build:
        # Step 3: Build final presentation
        # Get source from request file
        with open('temp/analysis_request.json', 'r') as f:
            request = json.load(f)
            source_path = request['source_file']

        # Auto-generate output filename if not provided
        output_path = args.output or generate_output_filename(source_path)

        build_presentation_from_insights(source_path, output_path, args.insights)

    else:
        parser.print_help()
        return 1

    return 0


if __name__ == '__main__':
    sys.exit(main())
