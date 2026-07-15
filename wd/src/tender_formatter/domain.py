from enum import StrEnum
from pathlib import Path

from pydantic import BaseModel, Field, model_validator


class BlockKind(StrEnum):
    HEADING = "heading"
    BODY = "body"
    CAPTION = "caption"
    EMPTY = "empty"


class RiskLevel(StrEnum):
    INFO = "info"
    REVIEW = "review"
    HIGH = "high"


class ParagraphEvidence(BaseModel):
    text: str
    style_name: str = ""
    outline_level: int | None = None
    numbering_level: int | None = None
    bold_ratio: float = 0.0
    alignment: str | None = None


class BlockDecision(BaseModel):
    index: int
    kind: BlockKind
    level: int | None = None
    confidence: float = Field(ge=0, le=1)
    reasons: list[str] = []
    risk: RiskLevel = RiskLevel.INFO

    @model_validator(mode="after")
    def validate_level(self):
        if self.kind == BlockKind.HEADING and self.level not in (1, 2, 3):
            raise ValueError("heading level must be 1, 2, or 3")
        if self.kind != BlockKind.HEADING and self.level is not None:
            raise ValueError("level is only valid for heading")
        return self


class TextStyleRules(BaseModel):
    east_asia_font: str = "宋体"
    latin_font: str = "Times New Roman"
    size_pt: float = 12
    bold: bool = False
    alignment: str = "justify"
    line_spacing: float = 1.5
    space_before_pt: float = 0
    space_after_pt: float = 0


class PageRules(BaseModel):
    top_cm: float = 2.5
    bottom_cm: float = 2.5
    left_cm: float = 3.0
    right_cm: float = 2.5
    gutter_cm: float = 0
    body_page_start: int = 1
    first_page_different: bool = True
    odd_even_pages: bool = False


class TableRules(BaseModel):
    width_percent: int = Field(default=100, ge=10, le=100)
    repeat_header: bool = True
    alignment: str = "center"


class ImageRules(BaseModel):
    max_width_cm: float = Field(default=15.5, gt=0)
    alignment: str = "center"
    caption_label: str = "图"


class FormatProfile(BaseModel):
    name: str
    template_path: Path | None = None
    high_confidence: float = Field(default=0.9, ge=0, le=1)
    review_confidence: float = Field(default=0.65, ge=0, le=1)
    body: TextStyleRules = TextStyleRules()
    headings: dict[int, TextStyleRules] = {}
    page: PageRules = PageRules()
    table: TableRules = TableRules()
    image: ImageRules = ImageRules()
    toc_levels: int = Field(default=3, ge=1, le=3)

    @model_validator(mode="after")
    def validate_thresholds(self):
        if self.high_confidence <= self.review_confidence:
            raise ValueError("high_confidence must exceed review_confidence")
        return self


class DocumentAnalysis(BaseModel):
    source: Path
    paragraphs: list[ParagraphEvidence]
    decisions: list[BlockDecision]
    table_count: int = 0
    image_count: int = 0
    section_count: int = 0
    structure_warnings: list[str] = []
    risky_table_indexes: list[int] = []


class FormatOperation(BaseModel):
    kind: str
    target: int | str
    parameters: dict = {}


class FormatPlan(BaseModel):
    source: Path
    output: Path
    operations: list[FormatOperation]
    warnings: list[str] = []


class ProcessingResult(BaseModel):
    output: Path
    report: Path
    operation_count: int
    warnings: list[str] = []
