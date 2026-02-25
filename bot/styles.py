"""GOST 7.32-2017 style constants and helpers for python-docx."""

from docx import Document
from docx.shared import Pt, Mm, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn

# --- Page setup ---
PAGE_WIDTH = Mm(210)
PAGE_HEIGHT = Mm(297)
MARGIN_LEFT = Mm(30)
MARGIN_RIGHT = Mm(15)
MARGIN_TOP = Mm(20)
MARGIN_BOTTOM = Mm(20)

# --- Font ---
FONT_NAME = "Times New Roman"
FONT_SIZE = Pt(14)
CODE_FONT_NAME = "Courier New"
CODE_FONT_SIZE = Pt(12)
TABLE_FONT_SIZE = Pt(12)

# --- Spacing ---
LINE_SPACING = 1.5  # multiplier
PARAGRAPH_INDENT = Cm(1.25)

# --- Work type labels ---
WORK_TYPES = {
    "lab": "лабораторной работе",
    "coursework": "курсовой работе",
    "practice": "практике",
    "report": "работе",
}

# --- Default title page values (ГУМОРФ-style) ---
DEFAULT_UNIVERSITY = (
    "Федеральное агентство морского и речного транспорта\n"
    "Федеральное государственное бюджетное образовательное учреждение\n"
    "высшего образования\n"
    "«Государственный университет морского и речного флота\n"
    "имени адмирала С.О. Макарова»"
)
DEFAULT_INSTITUTE = "ИНСТИТУТ ВОДНОГО ТРАНСПОРТА"
DEFAULT_DEPARTMENT = "КАФЕДРА «КОМПЛЕКСНОЕ ОБЕСПЕЧЕНИЕ ИНФОРМАЦИОННОЙ БЕЗОПАСНОСТИ»"
DEFAULT_CITY = "Санкт – Петербург"
DEFAULT_GROUP = "ИЗ–41"


def setup_page(doc: Document) -> None:
    """Configure page size and margins for GOST 7.32-2017."""
    section = doc.sections[0]
    section.page_width = PAGE_WIDTH
    section.page_height = PAGE_HEIGHT
    section.left_margin = MARGIN_LEFT
    section.right_margin = MARGIN_RIGHT
    section.top_margin = MARGIN_TOP
    section.bottom_margin = MARGIN_BOTTOM
    section.orientation = WD_ORIENT.PORTRAIT


def set_run_font(run, font_name=FONT_NAME, font_size=FONT_SIZE, bold=False, italic=False):
    """Apply font settings to a run."""
    run.font.name = font_name
    run.font.size = font_size
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor(0, 0, 0)
    # Force font for Cyrillic characters
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = r.makeelement(qn("w:rFonts"), {})
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:eastAsia"), font_name)
    rFonts.set(qn("w:cs"), font_name)


def set_paragraph_format(paragraph, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                         first_line_indent=PARAGRAPH_INDENT,
                         space_before=Pt(0), space_after=Pt(0),
                         line_spacing=LINE_SPACING):
    """Apply GOST paragraph formatting."""
    fmt = paragraph.paragraph_format
    fmt.alignment = alignment
    fmt.first_line_indent = first_line_indent
    fmt.space_before = space_before
    fmt.space_after = space_after
    fmt.line_spacing = line_spacing


def setup_default_style(doc: Document) -> None:
    """Set the Normal style to GOST defaults."""
    style = doc.styles["Normal"]
    style.font.name = FONT_NAME
    style.font.size = FONT_SIZE
    style.font.color.rgb = RGBColor(0, 0, 0)
    fmt = style.paragraph_format
    fmt.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    fmt.first_line_indent = PARAGRAPH_INDENT
    fmt.line_spacing = LINE_SPACING
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)
    # Cyrillic font fallback
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = style.element.makeelement(qn("w:rFonts"), {})
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:eastAsia"), FONT_NAME)
    rFonts.set(qn("w:cs"), FONT_NAME)
