import pytest

from tender_formatter.analyzer import analyze_docx
from tests.helpers import make_mixed_docx


def test_analyzer_extracts_structure_without_modifying_source(tmp_path):
    source = make_mixed_docx(tmp_path / "mixed.docx")
    before = source.read_bytes()

    result = analyze_docx(source)

    assert [p.text for p in result.paragraphs[:3]] == [
        "第一章 总体方案",
        "1.1 编制依据",
        "正文内容",
    ]
    assert result.paragraphs[0].style_name == "Heading 1"
    assert result.table_count == 1
    assert result.image_count == 1
    assert result.section_count == 1
    assert source.read_bytes() == before


def test_analyzer_rejects_non_docx(tmp_path):
    path = tmp_path / "input.doc"
    path.write_bytes(b"not a docx")

    with pytest.raises(ValueError, match=r"仅支持 \.docx"):
        analyze_docx(path)
