from pathlib import Path
from zipfile import BadZipFile

from docx import Document
from docx.opc.exceptions import PackageNotFoundError

from tender_formatter.domain import DocumentAnalysis, ParagraphEvidence


def _integer_property(element, child_name: str) -> int | None:
    properties = element.pPr
    if properties is None:
        return None
    child = getattr(properties, child_name, None)
    if child is None or child.val is None:
        return None
    return int(child.val)


def _numbering_level(paragraph) -> int | None:
    properties = paragraph._p.pPr
    if properties is None or properties.numPr is None:
        return None
    level = properties.numPr.ilvl
    if level is None or level.val is None:
        return None
    return int(level.val)


def _bold_ratio(paragraph) -> float:
    visible_runs = [(run, len(run.text)) for run in paragraph.runs if run.text]
    total = sum(length for _, length in visible_runs)
    if total == 0:
        return 0.0
    bold = sum(length for run, length in visible_runs if run.bold is True)
    return bold / total


def _extract_paragraph(paragraph) -> ParagraphEvidence:
    alignment = paragraph.alignment
    return ParagraphEvidence(
        text=paragraph.text,
        style_name=paragraph.style.name if paragraph.style is not None else "",
        outline_level=_integer_property(paragraph._p, "outlineLvl"),
        numbering_level=_numbering_level(paragraph),
        bold_ratio=_bold_ratio(paragraph),
        alignment=str(alignment) if alignment is not None else None,
    )


def analyze_docx(path: Path) -> DocumentAnalysis:
    if path.suffix.lower() != ".docx":
        raise ValueError("第一版仅支持 .docx 文件")
    try:
        document = Document(path)
    except (PackageNotFoundError, BadZipFile, KeyError) as exc:
        raise ValueError(f"无法打开 Word 文件：{path.name}") from exc

    paragraphs = [_extract_paragraph(paragraph) for paragraph in document.paragraphs]
    image_count = len(document.part.element.xpath(".//a:blip"))
    return DocumentAnalysis(
        source=path,
        paragraphs=paragraphs,
        decisions=[],
        table_count=len(document.tables),
        image_count=image_count,
        section_count=len(document.sections),
    )
