import json

from tender_formatter.report import write_report


def test_report_contains_counts_but_not_body_text(tmp_path):
    report = write_report(
        tmp_path / "report.json",
        operation_count=8,
        warnings=["编号跳级"],
        paragraph_count=100,
    )

    data = json.loads(report.read_text(encoding="utf-8"))

    assert data["operation_count"] == 8
    assert data["paragraph_count"] == 100
    assert data["warnings"] == ["编号跳级"]
    assert "document_text" not in data
