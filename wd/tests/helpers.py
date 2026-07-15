from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt


def _set_style_font(style, east_asia_font: str, size_pt: float, bold: bool) -> None:
    style.font.name = east_asia_font
    style.font.size = Pt(size_pt)
    style.font.bold = bold
    style.element.rPr.rFonts.set(qn("w:eastAsia"), east_asia_font)


def make_docx_with_styles(path: Path) -> Path:
    document = Document()
    _set_style_font(document.styles["Normal"], "宋体", 12, False)
    _set_style_font(document.styles["Heading 1"], "黑体", 16, True)
    _set_style_font(document.styles["Heading 2"], "黑体", 14, True)
    _set_style_font(document.styles["Heading 3"], "黑体", 12, True)
    document.save(path)
    return path
