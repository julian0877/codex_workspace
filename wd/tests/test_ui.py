from pathlib import Path

from tender_formatter.domain import (
    BlockDecision,
    BlockKind,
    DocumentAnalysis,
    ParagraphEvidence,
    RiskLevel,
)
from tender_formatter.ui.main_window import MainWindow


class FakeService:
    pass


def risk_analysis() -> DocumentAnalysis:
    paragraphs = [
        ParagraphEvidence(text="第一章 总则"),
        ParagraphEvidence(text="施工部署"),
        ParagraphEvidence(text="1.1.1 缺父级"),
    ]
    decisions = [
        BlockDecision(
            index=0, kind=BlockKind.HEADING, level=1, confidence=0.95
        ),
        BlockDecision(
            index=1,
            kind=BlockKind.HEADING,
            level=1,
            confidence=0.7,
            risk=RiskLevel.REVIEW,
        ),
        BlockDecision(
            index=2,
            kind=BlockKind.HEADING,
            level=3,
            confidence=0.85,
            risk=RiskLevel.HIGH,
        ),
    ]
    return DocumentAnalysis(
        source=Path("input.docx"),
        paragraphs=paragraphs,
        decisions=decisions,
    )


def test_next_is_disabled_until_source_and_template_exist(qtbot, tmp_path):
    window = MainWindow(FakeService())
    qtbot.addWidget(window)

    assert not window.next_button.isEnabled()
    window.source_edit.setText(str(tmp_path / "missing.docx"))
    assert not window.next_button.isEnabled()

    source = tmp_path / "source.docx"
    template = tmp_path / "template.docx"
    source.touch()
    template.touch()
    window.source_edit.setText(str(source))
    window.template_edit.setText(str(template))
    assert window.next_button.isEnabled()


def test_review_page_lists_only_review_and_high_risks(qtbot):
    window = MainWindow(FakeService())
    qtbot.addWidget(window)

    window.set_analysis(risk_analysis())

    assert window.review_model.rowCount() == 2


def test_generate_is_blocked_until_all_risks_confirmed(qtbot):
    window = MainWindow(FakeService())
    qtbot.addWidget(window)
    window.set_analysis(risk_analysis())

    assert not window.generate_button.isEnabled()
    window.review_model.confirm_all_as_suggested()

    assert window.generate_button.isEnabled()
