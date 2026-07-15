from pathlib import Path

from docx import Document

from tender_formatter.analyzer import analyze_docx
from tender_formatter.classifier import classify_paragraphs
from tender_formatter.cover import replace_cover_fields
from tender_formatter.domain import (
    BlockDecision,
    DocumentAnalysis,
    FormatProfile,
    ProcessingResult,
)
from tender_formatter.formatter import execute_docx_plan
from tender_formatter.planner import build_plan
from tender_formatter.report import write_report
from tender_formatter.word_adapter import WordAdapter, WordSettings


class FormatterService:
    def __init__(self, word: WordAdapter | None = None):
        self._word = word or WordAdapter()

    def analyze(
        self, source: Path, profile: FormatProfile
    ) -> DocumentAnalysis:
        analysis = analyze_docx(source)
        decisions = classify_paragraphs(analysis.paragraphs, profile)
        return analysis.model_copy(update={"decisions": decisions})

    def format(
        self,
        analysis: DocumentAnalysis,
        profile: FormatProfile,
        overrides: dict[int, BlockDecision],
        output: Path,
        cover_values: dict[str, str],
    ) -> ProcessingResult:
        plan = build_plan(analysis, profile, overrides, output)
        execute_docx_plan(plan, profile)

        document = Document(plan.output)
        missing_fields = replace_cover_fields(document, cover_values)
        if missing_fields:
            fields = "、".join(missing_fields)
            raise ValueError(f"封面必填字段缺少值：{fields}")
        document.save(plan.output)

        settings = WordSettings(body_page_start=profile.page.body_page_start)
        self._word.finalize(plan.output, settings)
        Document(plan.output)

        report_path = plan.output.with_suffix(".report.json")
        warnings = list(plan.warnings)
        write_report(
            report_path,
            operation_count=len(plan.operations),
            warnings=warnings,
            paragraph_count=len(analysis.paragraphs),
        )
        return ProcessingResult(
            output=plan.output,
            report=report_path,
            operation_count=len(plan.operations),
            warnings=warnings,
        )
