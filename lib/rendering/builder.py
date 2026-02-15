"""
Professional Slide Builder - Layer 4A

Builds executive-ready PowerPoint slides using python-pptx with Analytics template styling.
"""

from typing import Dict, List, Optional
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from lib.analysis.insights import Insight
from PIL import Image
import io


class AnalyticsStyleGuide:
    """Color palette and typography from Analytics template"""

    # Colors
    DARK_BLUE = RGBColor(0, 32, 96)
    ACCENT_BLUE = RGBColor(0, 114, 198)
    DARK_GRAY = RGBColor(64, 64, 64)
    BEIGE_BG = RGBColor(245, 239, 231)
    WHITE = RGBColor(255, 255, 255)

    # Typography
    FONT_NAME = "Segoe UI"
    HEADLINE_SIZE = Pt(24)
    INSIGHT_SIZE = Pt(14)
    LABEL_SIZE = Pt(12)

    # Layout measurements (in inches) - 16:9 widescreen format
    SLIDE_WIDTH = Inches(13.333)  # 16:9 aspect ratio
    SLIDE_HEIGHT = Inches(7.5)

    MARGIN = Inches(0.5)
    IMAGE_MAX_WIDTH = Inches(6.5)  # Wider for 16:9
    IMAGE_MAX_HEIGHT = Inches(5.5)
    TEXT_BOX_WIDTH = Inches(5.5)  # Wider text area
    SPACING = Inches(0.4)


class SlideBuilder:
    """Builds professional PowerPoint slides"""

    def __init__(self, source_prs: Presentation):
        """
        Initialize builder with source presentation.

        Args:
            source_prs: Source PowerPoint presentation to extract images from
        """
        self.source_prs = source_prs
        self.style = AnalyticsStyleGuide()

        # Create new presentation
        self.prs = Presentation()
        self.prs.slide_width = int(self.style.SLIDE_WIDTH)
        self.prs.slide_height = int(self.style.SLIDE_HEIGHT)

    def add_title_slide(self, title: str, subtitle: str = ""):
        """Add title slide with footer disclaimer"""
        slide_layout = self.prs.slide_layouts[0]  # Title layout
        slide = self.prs.slides.add_slide(slide_layout)

        # Define consistent positioning for title and subtitle
        title_left = self.style.MARGIN
        title_width = self.style.SLIDE_WIDTH - (2 * self.style.MARGIN)
        title_top = Inches(2.5)
        title_height = Inches(1.5)

        subtitle_top = Inches(4.2)
        subtitle_height = Inches(1.0)

        # Set title with explicit dimensions
        title_shape = slide.shapes.title
        title_shape.left = title_left
        title_shape.top = title_top
        title_shape.width = title_width
        title_shape.height = title_height
        title_shape.text = title
        self._style_text(title_shape.text_frame, font_size=Pt(32), bold=True, color=self.style.DARK_BLUE)

        # Set subtitle if provided (aligned with title, explicit dimensions)
        if subtitle and len(slide.placeholders) > 1:
            subtitle_shape = slide.placeholders[1]
            subtitle_shape.left = title_left
            subtitle_shape.top = subtitle_top
            subtitle_shape.width = title_width
            subtitle_shape.height = subtitle_height
            subtitle_shape.text = subtitle
            self._style_text(subtitle_shape.text_frame, font_size=Pt(18), color=self.style.DARK_GRAY)

        # Add footer disclaimer
        footer_left = self.style.MARGIN
        footer_top = self.style.SLIDE_HEIGHT - Inches(1)
        footer_width = self.style.SLIDE_WIDTH - (2 * self.style.MARGIN)
        footer_height = Inches(0.5)

        footer_box = slide.shapes.add_textbox(footer_left, footer_top, footer_width, footer_height)
        footer_frame = footer_box.text_frame
        footer_frame.text = "AI generated. Verify before sharing"
        self._style_text(
            footer_frame,
            font_size=Pt(10),
            color=self.style.DARK_GRAY,
            alignment=PP_ALIGN.CENTER
        )

        return slide

    def add_section_slide(self, section_title: str):
        """Add section divider slide"""
        slide_layout = self.prs.slide_layouts[6]  # Blank layout
        slide = self.prs.slides.add_slide(slide_layout)

        # Add section title centered
        left = self.style.MARGIN
        top = Inches(3)
        width = self.style.SLIDE_WIDTH - (2 * self.style.MARGIN)
        height = Inches(1.5)

        textbox = slide.shapes.add_textbox(left, top, width, height)
        text_frame = textbox.text_frame
        text_frame.text = section_title

        self._style_text(
            text_frame,
            font_size=Pt(36),
            bold=True,
            color=self.style.ACCENT_BLUE,
            alignment=PP_ALIGN.CENTER
        )

        return slide

    def add_executive_summary_slide(self, summary_bullets: List[str]):
        """Add executive summary slide with 5 key insights"""
        slide_layout = self.prs.slide_layouts[6]  # Blank layout
        slide = self.prs.slides.add_slide(slide_layout)

        # Add title
        title_box = slide.shapes.add_textbox(
            self.style.MARGIN,
            self.style.MARGIN,
            self.style.SLIDE_WIDTH - (2 * self.style.MARGIN),
            Inches(1)
        )
        title_frame = title_box.text_frame
        title_frame.text = "Executive Summary"
        self._style_text(
            title_frame,
            font_size=Pt(32),
            bold=True,
            color=self.style.DARK_BLUE
        )

        # Add bullets
        bullets_top = Inches(1.8)
        bullets_height = self.style.SLIDE_HEIGHT - bullets_top - self.style.MARGIN

        bullets_box = slide.shapes.add_textbox(
            self.style.MARGIN,
            bullets_top,
            self.style.SLIDE_WIDTH - (2 * self.style.MARGIN),
            bullets_height
        )

        text_frame = bullets_box.text_frame
        text_frame.word_wrap = True
        text_frame.vertical_anchor = MSO_ANCHOR.TOP

        for i, bullet in enumerate(summary_bullets[:5]):  # Max 5 bullets
            if i > 0:
                text_frame.add_paragraph()

            p = text_frame.paragraphs[i]
            p.text = f"• {bullet}"
            p.level = 0

            self._style_paragraph(
                p,
                font_size=Pt(16),
                color=self.style.DARK_GRAY,
                space_after=Pt(18)
            )

        return slide

    def add_recommendations_slide(self, recommendations: List[str]):
        """Add next steps/recommendations slide with 3-5 actions"""
        slide_layout = self.prs.slide_layouts[6]  # Blank layout
        slide = self.prs.slides.add_slide(slide_layout)

        # Add title
        title_box = slide.shapes.add_textbox(
            self.style.MARGIN,
            self.style.MARGIN,
            self.style.SLIDE_WIDTH - (2 * self.style.MARGIN),
            Inches(1)
        )
        title_frame = title_box.text_frame
        title_frame.text = "Next Steps & Recommendations"
        self._style_text(
            title_frame,
            font_size=Pt(32),
            bold=True,
            color=self.style.ACCENT_BLUE
        )

        # Add recommendations
        recs_top = Inches(1.8)
        recs_height = self.style.SLIDE_HEIGHT - recs_top - self.style.MARGIN

        recs_box = slide.shapes.add_textbox(
            self.style.MARGIN,
            recs_top,
            self.style.SLIDE_WIDTH - (2 * self.style.MARGIN),
            recs_height
        )

        text_frame = recs_box.text_frame
        text_frame.word_wrap = True
        text_frame.vertical_anchor = MSO_ANCHOR.TOP

        for i, rec in enumerate(recommendations):  # 3-5 recommendations
            if i > 0:
                text_frame.add_paragraph()

            p = text_frame.paragraphs[i]
            p.text = f"{i+1}. {rec}"
            p.level = 0

            self._style_paragraph(
                p,
                font_size=Pt(16),
                color=self.style.DARK_GRAY,
                space_after=Pt(18)
            )

        return slide

    def add_insight_slide(
        self,
        slide_number: int,
        headline: str,
        insights: List[str],
        source_image: Optional[Image.Image] = None
    ):
        """
        Add content slide with headline, insights, and optional screenshot.

        Args:
            slide_number: Slide number from source
            headline: Insight-driven headline
            insights: List of 2-3 bullet points
            source_image: Optional screenshot from source slide
        """
        slide_layout = self.prs.slide_layouts[6]  # Blank layout
        slide = self.prs.slides.add_slide(slide_layout)

        # Calculate layout positions
        image_left = self.style.MARGIN
        image_top = Inches(1.5)

        text_left = image_left + self.style.IMAGE_MAX_WIDTH + self.style.SPACING
        text_width = self.style.SLIDE_WIDTH - text_left - self.style.MARGIN

        # Add headline
        headline_box = slide.shapes.add_textbox(
            self.style.MARGIN,
            self.style.MARGIN,
            self.style.SLIDE_WIDTH - (2 * self.style.MARGIN),
            Inches(1)
        )
        headline_frame = headline_box.text_frame
        headline_frame.text = headline
        headline_frame.word_wrap = True
        self._style_text(
            headline_frame,
            font_size=self.style.HEADLINE_SIZE,
            bold=True,
            color=self.style.ACCENT_BLUE
        )

        # Add source image if provided
        if source_image:
            self._add_image_with_border(
                slide,
                source_image,
                image_left,
                image_top,
                self.style.IMAGE_MAX_WIDTH,
                self.style.IMAGE_MAX_HEIGHT
            )

        # Add insights as bullet points
        insights_top = image_top
        insights_height = self.style.IMAGE_MAX_HEIGHT

        insights_box = slide.shapes.add_textbox(
            text_left,
            insights_top,
            text_width,
            insights_height
        )

        text_frame = insights_box.text_frame
        text_frame.word_wrap = True
        text_frame.vertical_anchor = MSO_ANCHOR.TOP

        for i, insight in enumerate(insights[:3]):  # Max 3 insights
            if i > 0:
                text_frame.add_paragraph()

            p = text_frame.paragraphs[i]
            p.text = f"• {insight}"
            p.level = 0

            self._style_paragraph(
                p,
                font_size=self.style.INSIGHT_SIZE,
                color=self.style.DARK_GRAY,
                space_after=Pt(12)
            )

        return slide

    def _add_image_with_border(
        self,
        slide,
        image: Image.Image,
        left: float,
        top: float,
        max_width: float,
        max_height: float
    ):
        """Add image with white background and blue border"""
        # Calculate scaled dimensions maintaining aspect ratio
        img_width, img_height = image.size
        aspect_ratio = img_width / img_height

        if aspect_ratio > 1:  # Landscape
            width = min(max_width, Inches(img_width / 96))  # Assume 96 DPI
            height = width / aspect_ratio
        else:  # Portrait
            height = min(max_height, Inches(img_height / 96))
            width = height * aspect_ratio

        # Ensure doesn't exceed max dimensions
        if width > max_width:
            width = max_width
            height = width / aspect_ratio

        if height > max_height:
            height = max_height
            width = height * aspect_ratio

        # Save image to BytesIO
        img_stream = io.BytesIO()
        image.save(img_stream, format='PNG')
        img_stream.seek(0)

        # Add image
        pic = slide.shapes.add_picture(img_stream, left, top, width, height)

        # Add border (blue line)
        line = pic.line
        line.color.rgb = self.style.ACCENT_BLUE
        line.width = Pt(2)

        return pic

    def _style_text(
        self,
        text_frame,
        font_size: Pt,
        bold: bool = False,
        color: RGBColor = None,
        alignment: PP_ALIGN = PP_ALIGN.LEFT
    ):
        """Apply text styling"""
        for paragraph in text_frame.paragraphs:
            self._style_paragraph(paragraph, font_size, bold, color, alignment)

    def _style_paragraph(
        self,
        paragraph,
        font_size: Pt,
        bold: bool = False,
        color: RGBColor = None,
        alignment: PP_ALIGN = PP_ALIGN.LEFT,
        space_after: Pt = None
    ):
        """Apply paragraph styling"""
        paragraph.alignment = alignment

        if space_after:
            paragraph.space_after = space_after

        for run in paragraph.runs:
            font = run.font
            font.name = self.style.FONT_NAME
            font.size = font_size
            font.bold = bold

            if color:
                font.color.rgb = color

    def save(self, output_path: str):
        """Save presentation to file"""
        self.prs.save(output_path)


def _get_source_images_from_temp(source_path: str) -> Dict[int, str]:
    """
    Get already-extracted images from temp/ directory.
    Used when source is PDF (images already extracted during prepare phase).

    Args:
        source_path: Path to source file (used for validation)

    Returns:
        Dictionary mapping slide_number → image_path
    """
    import json

    try:
        with open('temp/analysis_request.json', 'r', encoding='utf-8') as f:
            request = json.load(f)

        # Build mapping of slide number to image path
        image_map = {}
        for slide_info in request.get('slides', []):
            slide_num = slide_info['slide_number']
            image_path = slide_info['image_path']
            image_map[slide_num] = image_path

        return image_map

    except (FileNotFoundError, json.JSONDecodeError, KeyError) as e:
        print(f"  WARNING: Could not load temp images: {e}")
        return {}


def extract_slide_image(source_prs: Presentation, slide_number: int) -> Optional[Image.Image]:
    """
    Extract screenshot from source slide.

    Args:
        source_prs: Source presentation
        slide_number: Slide number (0-indexed)

    Returns:
        PIL Image or None if extraction fails
    """
    try:
        if slide_number >= len(source_prs.slides):
            return None

        slide = source_prs.slides[slide_number]

        # Look for images in slide
        for shape in slide.shapes:
            if shape.shape_type == 13:  # Picture
                image_stream = io.BytesIO(shape.image.blob)
                return Image.open(image_stream)

        return None

    except Exception:
        return None


def render_presentation(
    source_path: str,
    insights: Dict[str, Insight],
    output_path: str
):
    """
    Main entry point: Render executive presentation.
    Supports both .pptx and .pdf source files.

    Args:
        source_path: Path to source file (.pptx or .pdf)
        insights: Dictionary of slide titles to Insights
        output_path: Path for output PowerPoint
    """
    file_type = Path(source_path).suffix.lower()

    # For PDF sources, images are already in temp/ directory
    # For PPTX sources, we load the presentation directly
    if file_type == '.pdf':
        # Create a dummy presentation (not used for image extraction)
        from pptx import Presentation as PrsClass
        source_prs = PrsClass()
        source_images_map = _get_source_images_from_temp(source_path)
    else:
        # Load source presentation for PPTX
        source_prs = Presentation(source_path)
        source_images_map = None

    # Create builder
    builder = SlideBuilder(source_prs)

    # Add title slide
    builder.add_title_slide(
        "Executive Analytics Dashboard",
        "Insights & Recommendations"
    )

    # Add executive summary slide (if provided)
    if '__executive_summary__' in insights:
        executive_summary = insights['__executive_summary__']
        if isinstance(executive_summary, list) and executive_summary:
            builder.add_executive_summary_slide(executive_summary)

    # Add content slides
    slide_mapping = {}  # Track which source slides we've processed

    # For PDF sources, iterate through insights directly (no source slides to iterate)
    # For PPTX sources, iterate through source slides as before
    if source_images_map is not None:
        # PDF workflow: Use temp/ images and insights mapping
        for title, insight in insights.items():
            # Find slide number from analysis request
            # We'll iterate through the map to find the matching title
            source_image = None
            slide_num = None

            # Load analysis request to get slide numbers
            import json
            try:
                with open('temp/analysis_request.json', 'r', encoding='utf-8') as f:
                    request = json.load(f)
                    for slide_info in request.get('slides', []):
                        if slide_info['title'] == title:
                            slide_num = slide_info['slide_number']
                            image_path = source_images_map.get(slide_num)
                            if image_path and Path(image_path).exists():
                                source_image = Image.open(image_path)
                            break
            except Exception:
                pass

            if source_image and slide_num:
                builder.add_insight_slide(
                    slide_number=slide_num,
                    headline=insight.headline,
                    insights=insight.bullet_points,
                    source_image=source_image
                )
    else:
        # PPTX workflow: Original logic
        for slide_idx, slide in enumerate(source_prs.slides):
            # Extract slide title (first shape with text, typically)
            slide_title = ""
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text.strip():
                    slide_title = shape.text.strip()
                    break

            # Remove emoji characters to match parsed titles
            import re
            slide_title_clean = re.sub(r'[\U00010000-\U0010ffff]', '', slide_title).strip()

            # Find matching insight
            insight = insights.get(slide_title_clean)

            if insight:
                # Extract image from source slide
                source_image = extract_slide_image(source_prs, slide_idx)

                # Add insight slide
                builder.add_insight_slide(
                    slide_number=slide_idx + 1,
                    headline=insight.headline,
                    insights=insight.bullet_points,
                    source_image=source_image
                )

    # Add recommendations slide (if provided)
    if '__recommendations__' in insights:
        recommendations = insights['__recommendations__']
        if isinstance(recommendations, list) and recommendations:
            builder.add_recommendations_slide(recommendations)

    # Save presentation
    builder.save(output_path)
