"""Template-mode DOCX rendering.

This is deliberately a placeholder renderer, not a full report writer.  It
copies the original template (layout, styles, tables, headers/footers, RTL and
numbering) and replaces only ``{{...}}`` placeholders with validated values
from :class:`~local_office_ai.reporting.specification.ReportContent`.
"""

from __future__ import annotations

import os
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from local_office_ai.core.exceptions import ReportFormatError, ReportTemplateError
from local_office_ai.reporting.formats import OutputFormat, ensure_extension_matches
from local_office_ai.reporting.specification import ReportContent
from local_office_ai.reporting.template import PLACEHOLDER_PATTERN, ReportTemplate


@dataclass
class TemplateRenderResult:
    output_path: Path
    replaced: int = 0
    unresolved_placeholders: tuple[str, ...] = ()
    warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "output_path": str(self.output_path),
            "replaced": self.replaced,
            "unresolved_placeholders": list(self.unresolved_placeholders),
            "warnings": list(self.warnings),
        }


class DocxTemplateRenderer:
    """Replace template placeholders while keeping the template's formatting."""

    def render(
        self,
        template: ReportTemplate,
        content: ReportContent,
        output_path: Path | str,
    ) -> TemplateRenderResult:
        content.ensure_valid()
        if not template.valid:
            raise ReportTemplateError(
                "قالب Word معتبر نیست: " + "؛ ".join(template.errors or ["دلیل نامشخص"])
            )
        output = ensure_extension_matches(Path(output_path), OutputFormat.DOCX)
        if output.resolve() == Path(template.path).resolve():
            raise ReportFormatError("مسیر خروجی با مسیر قالب Word یکسان است؛ قالب باید دست‌نخورده بماند.")
        values = content.placeholder_values()
        Document = self._document_class()
        try:
            document = Document(template.path)
        except Exception as exc:  # noqa: BLE001 - parser errors vary by python-docx version
            raise ReportTemplateError(f"قالب Word قابل خواندن نیست: {template.path} ({exc})") from exc

        state = _RenderState(values=values)
        self._replace_in_paragraphs(document.paragraphs, state)
        for table in document.tables:
            self._replace_in_table(table, state)
        for section in document.sections:
            for container in (
                section.header,
                section.first_page_header,
                section.even_page_header,
                section.footer,
                section.first_page_footer,
                section.even_page_footer,
            ):
                if _is_linked(container):
                    continue
                self._replace_in_paragraphs(container.paragraphs, state)
                for table in container.tables:
                    self._replace_in_table(table, state)

        self._save_atomic(document, output)
        warnings = list(template.warnings)
        if state.unresolved:
            warnings.append(
                "برای این جایگاه‌ها مقدار معتبر پیدا نشد و متن اصلی حفظ شد: "
                + "، ".join(sorted(state.unresolved))
            )
        return TemplateRenderResult(
            output_path=output,
            replaced=state.replaced,
            unresolved_placeholders=tuple(sorted(state.unresolved)),
            warnings=warnings,
        )

    @staticmethod
    def _document_class() -> Any:
        try:
            from docx import Document
        except ImportError as exc:  # pragma: no cover - optional dependency branch
            raise ReportTemplateError(
                "برای تولید گزارش Word، کتابخانه python-docx نصب نیست."
            ) from exc
        return Document

    def _replace_in_paragraphs(self, paragraphs: Any, state: "_RenderState") -> None:
        for paragraph in paragraphs:
            self._replace_in_paragraph(paragraph, state)

    def _replace_in_table(self, table: Any, state: "_RenderState") -> None:
        for row in table.rows:
            seen_cells: set[int] = set()
            for cell in row.cells:
                identity = id(cell._tc)
                if identity in seen_cells:
                    continue
                seen_cells.add(identity)
                self._replace_in_paragraphs(cell.paragraphs, state)
                for nested in cell.tables:
                    self._replace_in_table(nested, state)

    def _replace_in_paragraph(self, paragraph: Any, state: "_RenderState") -> None:
        runs = list(paragraph.runs)
        if not runs:
            return
        search_from = 0
        while True:
            combined = "".join(run.text or "" for run in runs)
            match = PLACEHOLDER_PATTERN.search(combined, search_from)
            if match is None:
                return
            key = match.group(1).strip()
            value = state.values.get(key)
            if value is None:
                state.unresolved.add(key)
                search_from = match.end()
                continue
            spans = _run_spans(runs)
            start_run, start_offset = _locate(spans, match.start())
            end_run, end_offset = _locate(spans, match.end() - 1)
            end_offset += 1
            if start_run is None or end_run is None:
                state.unresolved.add(key)
                search_from = match.end()
                continue
            involved = runs[start_run : end_run + 1]
            if not all(_run_is_text_only(run) for run in involved):
                state.unresolved.add(key)
                search_from = match.end()
                continue
            if start_run == end_run:
                text = runs[start_run].text or ""
                runs[start_run].text = text[:start_offset] + value + text[end_offset:]
            else:
                first_text = runs[start_run].text or ""
                runs[start_run].text = first_text[:start_offset] + value
                for run in runs[start_run + 1 : end_run]:
                    run.text = ""
                last_text = runs[end_run].text or ""
                runs[end_run].text = last_text[end_offset:]
            state.replaced += 1
            search_from = match.start() + len(value)

    @staticmethod
    def _save_atomic(document: Any, output: Path) -> None:
        output.parent.mkdir(parents=True, exist_ok=True)
        temp = output.with_name(output.stem + ".tmp" + output.suffix)
        try:
            document.save(str(temp))
            if not zipfile.is_zipfile(temp):
                raise ReportFormatError("خروجی Word معتبر ساخته نشد (فایل ZIP نیست).")
            with zipfile.ZipFile(temp) as archive:
                if "word/document.xml" not in archive.namelist():
                    raise ReportFormatError("خروجی Word معتبر ساخته نشد (word/document.xml ندارد).")
            os.replace(temp, output)
        except Exception:
            temp.unlink(missing_ok=True)
            raise


@dataclass
class _RenderState:
    values: dict[str, str]
    replaced: int = 0
    unresolved: set[str] = field(default_factory=set)


def _is_linked(container: Any) -> bool:
    try:
        return bool(container.is_linked_to_previous)
    except (AttributeError, ValueError):
        return False


def _run_spans(runs: list[Any]) -> list[tuple[int, int]]:
    spans: list[tuple[int, int]] = []
    cursor = 0
    for run in runs:
        length = len(run.text or "")
        spans.append((cursor, cursor + length))
        cursor += length
    return spans


def _locate(spans: list[tuple[int, int]], index: int) -> tuple[int | None, int]:
    for run_index, (start, end) in enumerate(spans):
        if start <= index < end:
            return run_index, index - start
    return None, 0


def _run_is_text_only(run: Any) -> bool:
    try:
        from docx.oxml.ns import qn
    except ImportError:  # pragma: no cover - optional dependency branch
        return True
    allowed = {qn("w:rPr"), qn("w:t")}
    return all(child.tag in allowed for child in run._r)
