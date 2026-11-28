"""Reopen a generated DOCX and verify its structure, RTL and figures.

A file named ``.docx`` is not a Word document.  This module checks the real
package (ZIP members), reopens it with python-docx, verifies that the expected
sections and tables exist, confirms RTL properties, page numbering, and that
every number in the document is present in the validated report content.
"""

from __future__ import annotations

import zipfile
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Iterable

from local_office_ai.core.exceptions import ReportFormatError
from local_office_ai.excel.normalization import (
    canonical_number,
    extract_numeric_claims,
    parse_number,
)
from local_office_ai.reporting.docx_rtl import qn
from local_office_ai.reporting.template import PLACEHOLDER_PATTERN

REQUIRED_DOCX_MEMBERS = ("[Content_Types].xml", "_rels/.rels", "word/document.xml")


@dataclass
class DocxValidation:
    path: str
    valid: bool
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    paragraph_count: int = 0
    table_count: int = 0
    heading_titles: tuple[str, ...] = ()
    sections_with_rtl: int = 0
    rtl_runs: int = 0
    rtl_tables: int = 0
    has_header: bool = False
    has_footer: bool = False
    has_page_number_field: bool = False
    numeric_tokens: int = 0
    unsupported_numbers: tuple[str, ...] = ()
    unresolved_placeholders: tuple[str, ...] = ()

    def to_dict(self) -> dict[str, Any]:
        return {
            "path": self.path,
            "valid": self.valid,
            "errors": list(self.errors),
            "warnings": list(self.warnings),
            "paragraph_count": self.paragraph_count,
            "table_count": self.table_count,
            "heading_titles": list(self.heading_titles),
            "sections_with_rtl": self.sections_with_rtl,
            "rtl_runs": self.rtl_runs,
            "rtl_tables": self.rtl_tables,
            "has_header": self.has_header,
            "has_footer": self.has_footer,
            "has_page_number_field": self.has_page_number_field,
            "numeric_tokens": self.numeric_tokens,
            "unsupported_numbers": list(self.unsupported_numbers),
            "unresolved_placeholders": list(self.unresolved_placeholders),
        }


def validate_docx_output(
    path: Path | str,
    content: Any = None,
    *,
    required_sections: Iterable[str] = (),
    minimum_tables: int = 0,
    require_rtl: bool = True,
    require_page_numbers: bool = True,
    allow_unresolved_placeholders: bool = False,
    allow_benign_numbers: bool = False,
    extra_allowed_numbers: Iterable[str] = (),
) -> DocxValidation:
    """Validate a delivered DOCX and return every finding, not just a boolean."""

    target = Path(path)
    validation = DocxValidation(path=str(target), valid=False)
    if not target.exists():
        validation.errors.append(f"فایل Word ساخته نشد: {target}")
        return validation
    if not target.is_file() or target.stat().st_size <= 0:
        validation.errors.append(f"فایل Word خالی یا غیرعادی است: {target}")
        return validation
    if not zipfile.is_zipfile(target):
        validation.errors.append("خروجی Word یک بسته ZIP معتبر نیست.")
        return validation
    try:
        with zipfile.ZipFile(target) as archive:
            names = set(archive.namelist())
            document_xml = archive.read("word/document.xml") if "word/document.xml" in names else b""
    except (OSError, zipfile.BadZipFile) as exc:
        validation.errors.append(f"بسته Word قابل خواندن نیست: {exc}")
        return validation
    missing = [member for member in REQUIRED_DOCX_MEMBERS if member not in names]
    if missing:
        validation.errors.append("اجزای لازم DOCX غایب است: " + "، ".join(missing))
        return validation
    try:
        from lxml import etree

        etree.fromstring(document_xml)
    except Exception as exc:  # noqa: BLE001 - lxml raises XMLSyntaxError and friends
        validation.errors.append(f"XML سند Word معتبر نیست: {exc}")
        return validation

    try:
        from docx import Document

        document = Document(str(target))
    except Exception as exc:  # noqa: BLE001 - python-docx raises several parser errors
        validation.errors.append(f"فایل Word با python-docx باز نشد: {exc}")
        return validation

    validation.paragraph_count = len(document.paragraphs)
    validation.table_count = len(document.tables)
    if validation.paragraph_count == 0 and validation.table_count == 0:
        validation.errors.append("سند Word هیچ پاراگراف یا جدولی ندارد.")

    headings = [
        paragraph.text.strip()
        for paragraph in document.paragraphs
        if _heading_level(paragraph) > 0 and paragraph.text.strip()
    ]
    validation.heading_titles = tuple(headings)

    normalized_headings = {_normalize(title) for title in headings}
    for title in required_sections:
        if _normalize(title) not in normalized_headings:
            validation.errors.append(f"بخش مورد انتظار در سند پیدا نشد: {title}")

    if validation.table_count < minimum_tables:
        validation.errors.append(
            f"تعداد جدول‌های سند ({validation.table_count}) کمتر از حد انتظار ({minimum_tables}) است."
        )

    validation.sections_with_rtl = sum(
        1 for section in document.sections if section._sectPr.find(qn("w:bidi")) is not None
    )
    validation.rtl_runs = _count_tag(document, "w:rtl")
    validation.rtl_tables = _count_tag(document, "w:bidiVisual")
    if require_rtl:
        if validation.sections_with_rtl == 0:
            validation.errors.append("هیچ بخشی از سند جهت راست‌به‌چپ (w:bidi) ندارد.")
        if validation.rtl_runs == 0:
            validation.errors.append("هیچ اجرای متنی با ویژگی راست‌به‌چپ (w:rtl) ثبت نشده است.")
        if validation.table_count and validation.rtl_tables == 0:
            validation.errors.append("جدول‌های سند جهت راست‌به‌چپ (w:bidiVisual) ندارند.")

    header_texts: list[str] = []
    footer_texts: list[str] = []
    header_numeric_texts: list[str] = []
    footer_numeric_texts: list[str] = []
    for section in document.sections:
        for container, bucket in (
            (section.header, header_texts),
            (section.first_page_header, header_texts),
            (section.even_page_header, header_texts),
            (section.footer, footer_texts),
            (section.first_page_footer, footer_texts),
            (section.even_page_footer, footer_texts),
        ):
            if _is_linked(container):
                continue
            numeric_bucket = (
                header_numeric_texts if bucket is header_texts else footer_numeric_texts
            )
            for paragraph in container.paragraphs:
                bucket.append(paragraph.text)
                if not _paragraph_has_field(paragraph):
                    numeric_bucket.append(paragraph.text)
            for table in container.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            bucket.append(paragraph.text)
                            if not _paragraph_has_field(paragraph):
                                numeric_bucket.append(paragraph.text)
    validation.has_header = any(text.strip() for text in header_texts)
    validation.has_footer = any(text.strip() for text in footer_texts)
    validation.has_page_number_field = _has_page_field(document)
    if require_page_numbers and not validation.has_page_number_field:
        validation.errors.append("شماره صفحه (فیلد PAGE) در پاصفحه پیدا نشد.")

    body_texts = _numeric_safe_texts(document)
    joined = "\n".join(body_texts + header_texts + footer_texts)
    placeholders = sorted({match.group(1).strip() for match in PLACEHOLDER_PATTERN.finditer(joined)})
    validation.unresolved_placeholders = tuple(placeholders)
    if placeholders and not allow_unresolved_placeholders:
        validation.errors.append(
            "جایگاه‌های پرنشده در سند باقی مانده‌اند: " + "، ".join(placeholders)
        )

    if content is not None:
        approved = set(content.approved_numbers())
        for token in extra_allowed_numbers:
            canonical_extra = _canonical(token)
            if canonical_extra is not None:
                approved.add(canonical_extra)
        tokens: list[str] = []
        unsupported: list[str] = []
        for text in body_texts + header_numeric_texts + footer_numeric_texts:
            for token in extract_numeric_claims(text):
                tokens.append(token)
                canonical = _canonical(token)
                if canonical is None:
                    continue
                if canonical in approved:
                    continue
                if allow_benign_numbers and _benign(canonical):
                    continue
                if token not in unsupported:
                    unsupported.append(token)
        validation.numeric_tokens = len(tokens)
        validation.unsupported_numbers = tuple(unsupported)
        if unsupported:
            validation.errors.append(
                "اعداد بدون منبع معتبر در سند پیدا شد: " + "، ".join(unsupported[:10])
            )

    validation.valid = not validation.errors
    return validation


def ensure_valid_docx(path: Path | str, content: Any = None, **kwargs: Any) -> DocxValidation:
    validation = validate_docx_output(path, content, **kwargs)
    if not validation.valid:
        raise ReportFormatError(
            "خروجی Word اعتبارسنجی نشد: " + "؛ ".join(validation.errors)
        )
    return validation


def _heading_level(paragraph: Any) -> int:
    name = (getattr(paragraph.style, "name", "") or "").strip()
    if name.lower().startswith("heading"):
        suffix = name.split()[-1]
        return int(suffix) if suffix.isdigit() else 1
    if name.lower() in {"title", "subtitle"}:
        return 1
    try:
        outline = paragraph._p.pPr.outlineLvl
    except AttributeError:
        return 0
    return int(outline.val) + 1 if outline is not None else 0


def _count_tag(document: Any, tag: str) -> int:
    return sum(1 for _ in document.element.body.iter(qn(tag)))


def _has_page_field(document: Any) -> bool:
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
            for element in container._element.iter(qn("w:instrText")):
                if "PAGE" in (element.text or "").upper():
                    return True
    return False


def _is_linked(container: Any) -> bool:
    try:
        return bool(container.is_linked_to_previous)
    except (AttributeError, ValueError):
        return False


def _numeric_safe_texts(document: Any) -> list[str]:
    """Document text excluding Word field paragraphs (page numbers etc.)."""

    texts: list[str] = []
    for paragraph in document.paragraphs:
        if _paragraph_has_field(paragraph):
            continue
        texts.append(paragraph.text)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if _paragraph_has_field(paragraph):
                        continue
                    texts.append(paragraph.text)
    return texts


def _paragraph_has_field(paragraph: Any) -> bool:
    return any(element.tag == qn("w:fldChar") for element in paragraph._p.iter())


def _normalize(text: str) -> str:
    return " ".join((text or "").replace("\u200c", " ").split())


def _canonical(token: str) -> str | None:
    parsed = parse_number(token, numeric_context=True)
    if not isinstance(parsed.value, Decimal):
        return None
    try:
        return canonical_number(parsed.value)
    except (InvalidOperation, ValueError):
        return None


def _benign(canonical: str) -> bool:
    """Small counts and years are template numbering, not report figures."""

    try:
        value = Decimal(canonical)
    except (InvalidOperation, ValueError):
        return False
    if value == value.to_integral_value():
        integer = int(value)
        return 0 <= integer <= 100 or 1200 <= integer <= 2200
    return False
