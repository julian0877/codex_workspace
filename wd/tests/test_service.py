from docx import Document

from tender_formatter.domain import FormatProfile
from tender_formatter.service import FormatterService
from tests.helpers import make_mixed_docx


class FakeWordAdapter:
    def __init__(self):
        self.finalized = []

    def finalize(self, path, _settings):
        self.finalized.append(path)


def test_service_analyze_then_format_calls_word_and_preserves_source(tmp_path):
    source = make_mixed_docx(tmp_path / "input.docx")
    before = source.read_bytes()
    word = FakeWordAdapter()
    service = FormatterService(word=word)
    profile = FormatProfile(name="公司标准")

    analysis = service.analyze(source, profile)
    result = service.format(
        analysis,
        profile,
        overrides={},
        output=tmp_path / "input_已格式化.docx",
        cover_values={},
    )

    assert result.output.exists() and result.report.exists()
    assert source.read_bytes() == before
    assert word.finalized == [result.output]
    Document(result.output)
