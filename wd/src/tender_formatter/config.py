from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

from tender_formatter.domain import FormatProfile, TextStyleRules


def save_profile(profile: FormatProfile, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = path.with_suffix(path.suffix + ".tmp")
    temporary_path.write_text(profile.model_dump_json(indent=2), encoding="utf-8")
    temporary_path.replace(path)


def load_profile(path: Path) -> FormatProfile:
    return FormatProfile.model_validate_json(path.read_text(encoding="utf-8"))


def _style_rules(style) -> TextStyleRules:
    fonts = style.element.rPr.rFonts if style.element.rPr is not None else None
    east_asia_font = fonts.get(qn("w:eastAsia")) if fonts is not None else None
    paragraph_format = style.paragraph_format
    alignment = {
        WD_ALIGN_PARAGRAPH.LEFT: "left",
        WD_ALIGN_PARAGRAPH.CENTER: "center",
        WD_ALIGN_PARAGRAPH.RIGHT: "right",
        WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
    }.get(paragraph_format.alignment, "justify")
    line_spacing = paragraph_format.line_spacing
    if not isinstance(line_spacing, (int, float)):
        line_spacing = 1.5
    return TextStyleRules(
        east_asia_font=east_asia_font or style.font.name or "宋体",
        latin_font=style.font.name or "Times New Roman",
        size_pt=style.font.size.pt if style.font.size is not None else 12,
        bold=bool(style.font.bold),
        alignment=alignment,
        line_spacing=float(line_spacing),
        space_before_pt=(
            paragraph_format.space_before.pt
            if paragraph_format.space_before is not None
            else 0
        ),
        space_after_pt=(
            paragraph_format.space_after.pt
            if paragraph_format.space_after is not None
            else 0
        ),
    )


def read_template_profile(path: Path, name: str) -> FormatProfile:
    document = Document(path)
    headings = {
        level: _style_rules(document.styles[f"Heading {level}"])
        for level in (1, 2, 3)
    }
    section = document.sections[0]
    profile = FormatProfile(
        name=name,
        template_path=path,
        body=_style_rules(document.styles["Normal"]),
        headings=headings,
    )
    profile.page.top_cm = section.top_margin.cm
    profile.page.bottom_cm = section.bottom_margin.cm
    profile.page.left_cm = section.left_margin.cm
    profile.page.right_cm = section.right_margin.cm
    profile.page.gutter_cm = section.gutter.cm
    return profile
