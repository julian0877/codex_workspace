import shutil

import pytest
from docx import Document

from tender_formatter.domain import FormatProfile
from tender_formatter.service import FormatterService
from tests.helpers import make_cover_docx, make_mixed_docx


class FakeWordAdapter:
    def __init__(self, fail=False, drop_objects=False):
        self.assembled = []
        self.fail = fail
        self.drop_objects = drop_objects

    def assemble(
        self, content_path, output_path, template_path, _settings, cover_values
    ):
        self.assembled.append(
            (content_path, output_path, template_path, cover_values)
        )
        if self.fail:
            raise RuntimeError("Word failed")
        if self.drop_objects:
            Document().save(output_path)
        else:
            shutil.copy2(content_path, output_path)


def test_service_analyze_then_format_calls_word_and_preserves_source(tmp_path):
    source = make_mixed_docx(tmp_path / "input.docx")
    before = source.read_bytes()
    word = FakeWordAdapter()
    service = FormatterService(word=word)
    template = make_cover_docx(tmp_path / "template.docx", "{{目录}}")
    profile = FormatProfile(name="公司标准", template_path=template)

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
    assert len(word.assembled) == 1
    assert word.assembled[0][2] == template
    assert word.assembled[0][1] != result.output
    Document(result.output)


def test_service_failure_never_publishes_half_finished_output(tmp_path):
    source = make_mixed_docx(tmp_path / "input.docx")
    template = make_cover_docx(tmp_path / "template.docx", "{{目录}}")
    output = tmp_path / "input_已格式化.docx"
    service = FormatterService(word=FakeWordAdapter(fail=True))
    profile = FormatProfile(name="公司标准", template_path=template)
    analysis = service.analyze(source, profile)

    with pytest.raises(RuntimeError, match="Word failed"):
        service.format(analysis, profile, {}, output, {})

    assert not output.exists()


def test_service_refuses_to_overwrite_existing_output(tmp_path):
    source = make_mixed_docx(tmp_path / "input.docx")
    template = make_cover_docx(tmp_path / "template.docx", "{{目录}}")
    output = tmp_path / "input_已格式化.docx"
    output.write_bytes(b"existing")
    service = FormatterService(word=FakeWordAdapter())
    profile = FormatProfile(name="公司标准", template_path=template)
    analysis = service.analyze(source, profile)

    with pytest.raises(ValueError, match="输出文件已存在"):
        service.format(analysis, profile, {}, output, {})

    assert output.read_bytes() == b"existing"


def test_service_stops_publish_when_word_loses_tables_or_images(tmp_path):
    source = make_mixed_docx(tmp_path / "input.docx")
    template = make_cover_docx(tmp_path / "template.docx", "{{目录}}")
    output = tmp_path / "input_已格式化.docx"
    service = FormatterService(word=FakeWordAdapter(drop_objects=True))
    profile = FormatProfile(name="公司标准", template_path=template)
    analysis = service.analyze(source, profile)

    with pytest.raises(RuntimeError, match="对象数量减少"):
        service.format(analysis, profile, {}, output, {})

    assert not output.exists()
