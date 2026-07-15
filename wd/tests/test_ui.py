from pathlib import Path

from PySide6.QtCore import Qt

from tender_formatter.domain import (
    BlockDecision,
    BlockKind,
    DocumentAnalysis,
    ParagraphEvidence,
    RiskLevel,
)
from tender_formatter.ui.main_window import MainWindow
from tests.helpers import make_cover_docx, make_docx_with_styles


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


def test_generate_is_blocked_until_each_risk_is_confirmed(qtbot):
    window = MainWindow(FakeService())
    qtbot.addWidget(window)
    window.set_analysis(risk_analysis())

    assert not window.generate_button.isEnabled()
    window.review_model.confirm_row(0)

    assert not window.generate_button.isEnabled()
    window.review_model.confirm_row(1)

    assert window.generate_button.isEnabled()


def test_review_model_allows_user_to_correct_kind_and_level(qtbot):
    window = MainWindow(FakeService())
    qtbot.addWidget(window)
    window.set_analysis(risk_analysis())
    model = window.review_model

    assert model.setData(model.index(0, 3), "2", Qt.EditRole)
    model.confirm_row(0)

    assert model.overrides()[1].level == 2


def test_template_cover_fields_are_exposed_for_input(qtbot, tmp_path):
    source = make_docx_with_styles(tmp_path / "source.docx")
    template = make_cover_docx(
        tmp_path / "template.docx", "{{项目名称}}\n{{投标单位}}\n{{目录}}"
    )
    window = MainWindow(FakeService())
    qtbot.addWidget(window)
    window.source_edit.setText(str(source))
    window.template_edit.setText(str(template))
    window.output_edit.setText(str(tmp_path / "output.docx"))

    window._next()

    assert set(window.cover_edits) == {"项目名称", "投标单位"}


def test_structure_warnings_require_explicit_acknowledgement(qtbot):
    analysis = risk_analysis().model_copy(
        update={"structure_warnings": ["存在浮动图片，需要人工复核"]}
    )
    window = MainWindow(FakeService())
    qtbot.addWidget(window)
    window.set_analysis(analysis)
    window.review_model.confirm_row(0)
    window.review_model.confirm_row(1)

    assert not window.generate_button.isEnabled()
    window._acknowledge_structure_warnings()

    assert window.generate_button.isEnabled()
