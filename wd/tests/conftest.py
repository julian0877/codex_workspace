from pathlib import Path

import pytest

from tender_formatter.domain import (
    BlockDecision,
    BlockKind,
    DocumentAnalysis,
    FormatProfile,
    ParagraphEvidence,
    RiskLevel,
)


@pytest.fixture
def profile():
    return FormatProfile(name="公司标准")


@pytest.fixture
def analysis(tmp_path):
    source = tmp_path / "input.docx"
    source.write_bytes(b"source")
    paragraphs = [ParagraphEvidence(text="第一章 总则")]
    decisions = [
        BlockDecision(
            index=0, kind=BlockKind.HEADING, level=1, confidence=0.95
        )
    ]
    return DocumentAnalysis(
        source=source,
        paragraphs=paragraphs,
        decisions=decisions,
        table_count=1,
        image_count=1,
        section_count=1,
    )


@pytest.fixture
def analysis_with_review(analysis):
    return analysis.model_copy(
        update={
            "decisions": [
                BlockDecision(
                    index=0,
                    kind=BlockKind.HEADING,
                    level=1,
                    confidence=0.7,
                    risk=RiskLevel.REVIEW,
                )
            ]
        }
    )
