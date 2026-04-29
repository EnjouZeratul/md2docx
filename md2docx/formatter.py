"""Formatting utilities for DOCX elements.

Pure functions that operate on run/paragraph objects without holding state.
"""

from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def set_run_font(run, font_name: str = '宋体', bold: bool = False,
                 size: int = 12, italic: bool = False):
    """Set font properties on a run.

    Handles both Chinese and English fonts properly.

    Args:
        run: docx Run object
        font_name: Font family name
        bold: Bold style
        size: Font size in pt
        italic: Italic style
    """
    run.font.name = font_name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic

    # Set East Asian font for Chinese support
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)


def set_para_spacing(para, space_before: float = 0, space_after: float = 0,
                     line_spacing: float = 1.5, indent_cm: float = None,
                     align: str = 'left'):
    """Set paragraph spacing and alignment.

    Args:
        para: docx Paragraph object
        space_before: Space before paragraph in pt
        space_after: Space after paragraph in pt
        line_spacing: Line spacing multiplier
        indent_cm: First line indent in cm (None for no indent)
        align: Text alignment ('left', 'center', 'right', 'justify')
    """
    para.paragraph_format.space_before = Pt(space_before)
    para.paragraph_format.space_after = Pt(space_after)
    para.paragraph_format.line_spacing = line_spacing

    if indent_cm is not None:
        para.paragraph_format.first_line_indent = Cm(indent_cm)

    # Set alignment
    align_map = {
        'left': WD_ALIGN_PARAGRAPH.LEFT,
        'center': WD_ALIGN_PARAGRAPH.CENTER,
        'right': WD_ALIGN_PARAGRAPH.RIGHT,
        'justify': WD_ALIGN_PARAGRAPH.JUSTIFY,
    }
    para.paragraph_format.alignment = align_map.get(align, WD_ALIGN_PARAGRAPH.LEFT)


def set_outline_level(para, level: int):
    """Set outline level for a paragraph.

    Outline levels are used for TOC generation.
    Level 0 = 章 (Chapter)
    Level 1 = 节 (Section)
    Level 2 = 小节 (Subsection)

    Args:
        para: docx Paragraph object
        level: Outline level (0-8)
    """
    # Get or create paragraph properties
    pPr = para._element.get_or_add_pPr()

    # Create outline level element
    outlineLvl = OxmlElement('w:outlineLvl')
    outlineLvl.set(qn('w:val'), str(level))

    # Add to paragraph properties
    pPr.append(outlineLvl)


def add_header_border(para):
    """Add bottom border to header paragraph.

    Creates a single line border at the bottom of the paragraph.

    Args:
        para: docx Paragraph object (typically in header)
    """
    pPr = para._element.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')

    # Create bottom border
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '4')  # Border width in 1/8 pt
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '000000')

    pBdr.append(bottom)
    pPr.append(pBdr)
