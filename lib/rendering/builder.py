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
        """Add title slide"""
        slide_layout = self.prs.slide_layouts[0]  # Title layout
        slide = self.prs.slides.add_slide(slide_layout)

        # Set title
        title_shape = slide.shapes.title
        title_shape.text = title
        self._style_text(title_shape.text_frame, font_size=Pt(32), bold=True, color=self.style.DARK_BLUE)

        # Set subtitle if provided
        if subtitle and len(slide.placeholders) > 1:
            subtitle_shape = slide.placeholders[1]
            subtitle_shape.text = subtitle
            self._style_text(subtitle_shape.text_frame, font_size=Pt(18), color=self.style.DARK_GRAY)

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

    Args:
        source_path: Path to source PowerPoint
        insights: Dictionary of slide titles to Insights
        output_path: Path for output PowerPoint
    """
    # Load source presentation
    source_prs = Presentation(source_path)

    # Create builder
    builder = SlideBuilder(source_prs)

    # Add title slide
    builder.add_title_slide(
        "Executive Analytics Dashboard",
        "Insights & Recommendations"
    )

    # Add content slides
    slide_mapping = {}  # Track which source slides we've processed

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

    # Save presentation
    builder.save(output_path)
