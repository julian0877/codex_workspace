from tender_formatter.classifier import classify_paragraphs
from tender_formatter.domain import (
    BlockKind,
    FormatProfile,
    ParagraphEvidence,
    RiskLevel,
)


def classify(*items):
    return classify_paragraphs(list(items), FormatProfile(name="test"))


def test_chinese_chapter_and_decimal_levels_are_recognized():
    result = classify(
        ParagraphEvidence(text="第一章 总体施工部署"),
        ParagraphEvidence(text="1.1 编制依据"),
        ParagraphEvidence(text="1.1.1 规范标准"),
    )

    assert [(item.kind, item.level) for item in result] == [
        (BlockKind.HEADING, 1),
        (BlockKind.HEADING, 2),
        (BlockKind.HEADING, 3),
    ]


def test_word_numbering_and_heading_style_strengthen_evidence():
    result = classify(
        ParagraphEvidence(
            text="施工部署", style_name="Heading 1", numbering_level=0
        )
    )

    assert result[0].level == 1
    assert result[0].confidence >= 0.9
    assert result[0].risk == RiskLevel.INFO


def test_unnumbered_bold_short_paragraph_requires_review():
    result = classify(ParagraphEvidence(text="施工部署", bold_ratio=1.0))

    assert result[0].kind == BlockKind.HEADING
    assert result[0].risk == RiskLevel.REVIEW


def test_heading_level_jump_is_high_risk():
    result = classify(
        ParagraphEvidence(text="1.1 内容"),
        ParagraphEvidence(text="1.1.1 内容"),
    )

    assert result[0].risk == RiskLevel.HIGH
    assert "缺少父级标题" in result[0].reasons


def test_plain_and_empty_paragraphs_are_not_headings():
    result = classify(
        ParagraphEvidence(text="这是普通正文。"),
        ParagraphEvidence(text="   "),
    )

    assert result[0].kind == BlockKind.BODY
    assert result[1].kind == BlockKind.EMPTY


def test_duplicate_and_skipped_sibling_numbers_are_high_risk():
    result = classify(
        ParagraphEvidence(text="第一章 总则"),
        ParagraphEvidence(text="1.1 内容"),
        ParagraphEvidence(text="1.1 重复"),
        ParagraphEvidence(text="1.3 跳号"),
    )

    assert result[2].risk == RiskLevel.HIGH
    assert "标题编号重复" in result[2].reasons
    assert result[3].risk == RiskLevel.HIGH
    assert "同级标题编号不连续" in result[3].reasons
