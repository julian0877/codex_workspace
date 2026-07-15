import pytest

from tender_formatter.domain import BlockDecision, BlockKind
from tender_formatter.planner import build_plan, default_output_path


def test_plan_refuses_source_as_output(analysis, profile):
    with pytest.raises(ValueError, match="不能覆盖源文件"):
        build_plan(analysis, profile, {}, analysis.source)


def test_plan_requires_resolution_for_review_items(
    analysis_with_review, profile, tmp_path
):
    with pytest.raises(ValueError, match="尚有未确认"):
        build_plan(
            analysis_with_review,
            profile,
            {},
            tmp_path / "out.docx",
        )


def test_confirmed_override_allows_plan(analysis_with_review, profile, tmp_path):
    confirmed = BlockDecision(
        index=0, kind=BlockKind.HEADING, level=1, confidence=1.0
    )

    plan = build_plan(
        analysis_with_review,
        profile,
        {0: confirmed},
        tmp_path / "out.docx",
    )

    kinds = {operation.kind for operation in plan.operations}
    assert {
        "apply_paragraph",
        "apply_page",
        "format_tables",
        "format_images",
        "replace_cover",
        "word_finalize",
    } <= kinds


def test_default_output_path_adds_suffix(analysis):
    assert default_output_path(analysis.source).name == "input_已格式化.docx"
