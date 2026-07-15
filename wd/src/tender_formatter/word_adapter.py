from collections.abc import Callable
from pathlib import Path
import shutil

from pydantic import BaseModel, Field

from docx import Document

from tender_formatter.cover import find_cover_fields, replace_cover_fields


class WordAutomationError(RuntimeError):
    """Raised when Microsoft Word cannot finalize a document."""


class WordSettings(BaseModel):
    update_toc: bool = True
    body_page_start: int = Field(default=1, ge=1)
    body_section_index: int = Field(default=1, ge=1)
    first_page_different: bool = True
    odd_even_pages: bool = False


def _default_dispatch_ex(name: str):
    from win32com.client import DispatchEx

    return DispatchEx(name)


def _iter_collection(collection):
    for index in range(1, collection.Count + 1):
        yield index, collection(index)


class WordAdapter:
    def __init__(self, dispatch_ex: Callable = _default_dispatch_ex):
        self._dispatch_ex = dispatch_ex

    def _configure_sections(self, document, settings: WordSettings) -> None:
        for section_index, section in _iter_collection(document.Sections):
            section.PageSetup.DifferentFirstPageHeaderFooter = (
                settings.first_page_different
            )
            section.PageSetup.OddAndEvenPagesHeaderFooter = settings.odd_even_pages
            if section_index > 1:
                for _, header in _iter_collection(section.Headers):
                    header.LinkToPrevious = False
                for _, footer in _iter_collection(section.Footers):
                    footer.LinkToPrevious = False
            if section_index == settings.body_section_index:
                _, primary_footer = next(_iter_collection(section.Footers))
                primary_footer.PageNumbers.RestartNumberingAtSection = True
                primary_footer.PageNumbers.StartingNumber = settings.body_page_start

    def finalize(self, path: Path, settings: WordSettings) -> None:
        import pythoncom

        pythoncom.CoInitialize()
        application = None
        document = None
        try:
            application = self._dispatch_ex("Word.Application")
            application.Visible = False
            application.DisplayAlerts = 0
            document = application.Documents.Open(str(path.resolve()))
            self._configure_sections(document, settings)
            if settings.update_toc:
                for _, table_of_contents in _iter_collection(
                    document.TablesOfContents
                ):
                    table_of_contents.Update()
            document.Fields.Update()
            document.Repaginate()
            document.Save()
        except Exception as exc:
            raise WordAutomationError(f"Word 自动化失败：{exc}") from exc
        finally:
            if document is not None:
                try:
                    document.Close(False)
                except Exception:
                    pass
            if application is not None:
                try:
                    application.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()

    def assemble(
        self,
        content_path: Path,
        output_path: Path,
        template_path: Path,
        settings: WordSettings,
        cover_values: dict[str, str],
    ) -> None:
        import pythoncom

        required_fields = find_cover_fields(template_path) - {"目录"}
        missing = sorted(
            field for field in required_fields if not cover_values.get(field)
        )
        if missing:
            raise ValueError(f"封面必填字段缺少值：{'、'.join(missing)}")

        shutil.copy2(template_path, output_path)
        template_document = Document(output_path)
        replace_cover_fields(template_document, cover_values)
        template_document.save(output_path)

        pythoncom.CoInitialize()
        application = None
        document = None
        try:
            application = self._dispatch_ex("Word.Application")
            application.Visible = False
            application.DisplayAlerts = 0
            document = application.Documents.Open(str(output_path.resolve()))

            insertion = document.Content
            insertion.Collapse(0)
            insertion.InsertBreak(2)
            insertion.Collapse(0)
            insertion.InsertFile(str(content_path.resolve()))
            settings.body_section_index = document.Sections.Count

            toc_range = document.Content.Duplicate
            found_marker = toc_range.Find.Execute(FindText="{{目录}}")
            if found_marker:
                toc_range.Text = ""
                document.TablesOfContents.Add(
                    Range=toc_range,
                    UseHeadingStyles=True,
                    UpperHeadingLevel=1,
                    LowerHeadingLevel=3,
                    IncludePageNumbers=True,
                    RightAlignPageNumbers=True,
                )
            elif document.TablesOfContents.Count == 0:
                raise ValueError("企业样板缺少 {{目录}} 占位符或已有自动目录")

            self._configure_sections(document, settings)
            body_section = document.Sections(settings.body_section_index)
            primary_footer = body_section.Footers(1)
            if primary_footer.PageNumbers.Count == 0:
                primary_footer.PageNumbers.Add(1, True)
            for _, table_of_contents in _iter_collection(document.TablesOfContents):
                table_of_contents.Update()
            document.Fields.Update()
            document.Repaginate()
            document.Save()
        except Exception as exc:
            if isinstance(exc, (ValueError, WordAutomationError)):
                raise
            raise WordAutomationError(f"Word 自动化失败：{exc}") from exc
        finally:
            if document is not None:
                try:
                    document.Close(False)
                except Exception:
                    pass
            if application is not None:
                try:
                    application.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()
