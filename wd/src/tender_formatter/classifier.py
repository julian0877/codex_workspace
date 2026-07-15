import re

from tender_formatter.domain import (
    BlockDecision,
    BlockKind,
    FormatProfile,
    ParagraphEvidence,
    RiskLevel,
)


_CHAPTER = re.compile(r"^第[一二三四五六七八九十百]+章\s*")
_LEVEL_THREE = re.compile(r"^\d+\.\d+\.\d+(?:[、.]|\s)+\S")
_LEVEL_TWO = re.compile(r"^\d+\.\d+(?:[、.]|\s)+\S")
_LEVEL_ONE = re.compile(r"^\d+[、.]\s*\S")
_CAPTION = re.compile(r"^图\s*\d+(?:[.-]\d+)*\s+")
_HEADING_STYLE = re.compile(r"(?:Heading|标题)\s*([123])", re.IGNORECASE)


def _explicit_level(text: str) -> int | None:
    if _CHAPTER.match(text):
        return 1
    if _LEVEL_THREE.match(text):
        return 3
    if _LEVEL_TWO.match(text):
        return 2
    if _LEVEL_ONE.match(text):
        return 1
    return None


def _style_level(paragraph: ParagraphEvidence) -> int | None:
    match = _HEADING_STYLE.search(paragraph.style_name)
    if match:
        return int(match.group(1))
    if paragraph.outline_level in (0, 1, 2):
        return paragraph.outline_level + 1
    if paragraph.numbering_level in (0, 1, 2):
        return paragraph.numbering_level + 1
    return None


def _risk_for_confidence(confidence: float, profile: FormatProfile) -> RiskLevel:
    if confidence >= profile.high_confidence:
        return RiskLevel.INFO
    return RiskLevel.REVIEW


def _classify_one(
    index: int, paragraph: ParagraphEvidence, profile: FormatProfile
) -> BlockDecision:
    text = paragraph.text.strip()
    if not text:
        return BlockDecision(index=index, kind=BlockKind.EMPTY, confidence=1.0)
    if _CAPTION.match(text):
        return BlockDecision(
            index=index,
            kind=BlockKind.CAPTION,
            confidence=0.95,
            reasons=["匹配图片题注格式"],
        )

    explicit_level = _explicit_level(text)
    style_level = _style_level(paragraph)
    level = explicit_level or style_level
    reasons: list[str] = []
    confidence = 0.0
    if explicit_level is not None:
        confidence = 0.85
        reasons.append("匹配标题编号格式")
    if style_level is not None:
        if level is None:
            level = style_level
        confidence += 0.15 if explicit_level is not None else 0.75
        reasons.append("Word 标题样式或大纲层级")
    if paragraph.numbering_level is not None and level == paragraph.numbering_level + 1:
        confidence += 0.15
        reasons.append("Word 自动编号层级一致")
    if len(text) <= 40 and paragraph.bold_ratio >= 0.8:
        confidence += 0.05
        reasons.append("短段落且主要为粗体")

    if level is not None:
        confidence = min(confidence, 1.0)
        return BlockDecision(
            index=index,
            kind=BlockKind.HEADING,
            level=level,
            confidence=confidence,
            reasons=reasons,
            risk=_risk_for_confidence(confidence, profile),
        )
    if len(text) <= 40 and paragraph.bold_ratio >= 0.8:
        return BlockDecision(
            index=index,
            kind=BlockKind.HEADING,
            level=1,
            confidence=0.7,
            reasons=["无编号短粗体段落，层级需要确认"],
            risk=RiskLevel.REVIEW,
        )
    return BlockDecision(index=index, kind=BlockKind.BODY, confidence=0.95)


def classify_paragraphs(
    paragraphs: list[ParagraphEvidence], profile: FormatProfile
) -> list[BlockDecision]:
    decisions = [
        _classify_one(index, paragraph, profile)
        for index, paragraph in enumerate(paragraphs)
    ]
    seen_levels: set[int] = set()
    for decision in decisions:
        if decision.kind != BlockKind.HEADING or decision.level is None:
            continue
        if decision.level > 1 and decision.level - 1 not in seen_levels:
            decision.risk = RiskLevel.HIGH
            decision.reasons.append("缺少父级标题")
        seen_levels.discard(decision.level)
        seen_levels.update({decision.level})
        seen_levels.difference_update(
            level for level in tuple(seen_levels) if level > decision.level
        )
    return decisions
