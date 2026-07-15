from pathlib import Path

from docx import Document
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
    return TextStyleRules(
        east_asia_font=east_asia_font or style.font.name or "宋体",
        latin_font=style.font.name or "Times New Roman",
        size_pt=style.font.size.pt if style.font.size is not None else 12,
        bold=bool(style.font.bold),
    )


def read_template_profile(path: Path, name: str) -> FormatProfile:
    document = Document(path)
    headings = {
        level: _style_rules(document.styles[f"Heading {level}"])
        for level in (1, 2, 3)
    }
    return FormatProfile(
        name=name,
        template_path=path,
        body=_style_rules(document.styles["Normal"]),
        headings=headings,
    )
