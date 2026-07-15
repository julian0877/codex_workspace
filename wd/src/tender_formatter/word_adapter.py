from collections.abc import Callable
from pathlib import Path

from pydantic import BaseModel, Field


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
