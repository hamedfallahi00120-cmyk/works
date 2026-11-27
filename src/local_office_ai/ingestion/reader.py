"""Read common administrative files into a shared, bounded document model.

Supported: XLSX/XLSM (handled by the existing Excel workflow), CSV, DOCX,
PPTX, PDF, TXT and common raster images via optional Tesseract OCR.
Old binary DOC/PPT formats are intentionally not accepted.
"""

from __future__ import annotations

import csv
import os
import shutil
import zipfile
from pathlib import Path
from typing import Any

from local_office_ai.core.exceptions import EngineError
from local_office_ai.core.settings import Settings
from local_office_ai.ingestion.models import ExtractedDocument, ExtractedTable

SUPPORTED_DOCUMENT_SUFFIXES = {
    ".xlsx", ".xlsm", ".csv", ".docx", ".pptx", ".pdf", ".txt",
    ".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp",
}
EXCEL_SUFFIXES = {".xlsx", ".xlsm"}
IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp"}


class DocumentReader:
    def __init__(self, settings: Settings, *, max_rows_per_table: int = 1500, max_pages: int = 150) -> None:
        self.settings = settings
        self.max_rows_per_table = max(50, int(max_rows_per_table))
        self.max_pages = max(1, int(max_pages))

    def read(self, path: Path | str) -> ExtractedDocument:
        source = Path(path).expanduser().resolve()
        self._validate(source)
        suffix = source.suffix.lower()
        if suffix in EXCEL_SUFFIXES:
            raise EngineError("فایل Excel باید از موتور تخصصی Excel خوانده شود.")
        readers = {
            ".csv": self._read_csv,
            ".docx": self._read_docx,
            ".pptx": self._read_pptx,
            ".pdf": self._read_pdf,
            ".txt": self._read_text,
        }
        if suffix in IMAGE_SUFFIXES:
            return self._read_image(source)
        reader = readers.get(suffix)
        if reader is None:
            raise EngineError(f"این قالب پشتیبانی نمی‌شود: {suffix or 'بدون پسوند'}")
        return reader(source)

    def _validate(self, source: Path) -> None:
        if not source.exists() or not source.is_file():
            raise EngineError(f"فایل پیدا نشد یا فایل معمولی نیست: {source}")
        if source.suffix.lower() not in SUPPORTED_DOCUMENT_SUFFIXES:
            raise EngineError(f"قالب فایل پشتیبانی نمی‌شود: {source.suffix}")
        if source.name.startswith("~$"):
            raise EngineError("فایل موقت قفل‌شده Office قابل پردازش نیست.")
        if source.stat().st_size <= 0:
            raise EngineError("فایل خالی است.")
        if source.stat().st_size > self.settings.max_file_mb * 1024 * 1024:
            raise EngineError(f"حجم فایل از سقف {self.settings.max_file_mb} مگابایت بیشتر است.")

    @staticmethod
    def _decode_text(source: Path) -> tuple[str, str]:
        raw = source.read_bytes()
        for encoding in ("utf-8-sig", "utf-8", "cp1256", "utf-16", "cp1252"):
            try:
                return raw.decode(encoding), encoding
            except UnicodeDecodeError:
                continue
        return raw.decode("utf-8", errors="replace"), "utf-8 (replacement)"

    def _read_text(self, source: Path) -> ExtractedDocument:
        text, encoding = self._decode_text(source)
        return ExtractedDocument(
            source=source,
            kind="text",
            title=source.stem,
            text_blocks=[("متن", text[:120_000])],
            warnings=["متن برای جلوگیری از مصرف زیاد حافظه به ۱۲۰ هزار نویسه محدود شد."] if len(text) > 120_000 else [],
            metadata={"encoding": encoding},
        )

    def _read_csv(self, source: Path) -> ExtractedDocument:
        text, encoding = self._decode_text(source)
        sample = text[:8192]
        delimiter = ","
        try:
            delimiter = csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
        except csv.Error:
            # Persian CSV files are often semicolon or tab separated.
            counts = {candidate: sample.count(candidate) for candidate in (",", ";", "\t", "|")}
            delimiter = max(counts, key=counts.get) if max(counts.values(), default=0) else ","
        rows: list[list[str]] = []
        try:
            for row in csv.reader(text.splitlines(), delimiter=delimiter):
                if not row or not any(str(cell).strip() for cell in row):
                    continue
                rows.append([str(cell).strip() for cell in row])
                if len(rows) > self.max_rows_per_table + 1:
                    break
        except csv.Error as exc:
            raise EngineError(f"خواندن CSV ناموفق بود: {exc}") from exc
        if not rows:
            raise EngineError("فایل CSV داده‌ای ندارد.")
        width = max(len(row) for row in rows)
        padded = [row + [""] * (width - len(row)) for row in rows]
        first = padded[0]
        # A header is accepted when it is mostly non-numeric text.
        from local_office_ai.excel.normalization import parse_number
        numeric_header_cells = sum(1 for value in first if parse_number(value, numeric_context=True).value is not None)
        header_present = bool(first) and numeric_header_cells <= max(0, len(first) // 3)
        headers = [value or f"ستون {index + 1}" for index, value in enumerate(first)] if header_present else [f"ستون {index + 1}" for index in range(width)]
        data_rows = padded[1:] if header_present else padded
        truncated = len(data_rows) > self.max_rows_per_table
        data_rows = data_rows[: self.max_rows_per_table]
        table = ExtractedTable(
            source=source,
            location="CSV",
            title=source.stem,
            headers=headers,
            rows=data_rows,
            truncated=truncated,
            warnings=["ردیف‌های CSV به سقف استخراج محدود شدند."] if truncated else [],
        )
        return ExtractedDocument(
            source=source,
            kind="csv",
            title=source.stem,
            tables=[table],
            metadata={"encoding": encoding, "delimiter": delimiter, "header_present": header_present},
            warnings=list(table.warnings),
        )

    def _read_docx(self, source: Path) -> ExtractedDocument:
        try:
            from docx import Document
        except ImportError as exc:
            raise EngineError("برای خواندن Word، کتابخانه python-docx نصب نیست.") from exc
        try:
            document = Document(str(source))
        except Exception as exc:  # noqa: BLE001
            raise EngineError(f"خواندن Word ناموفق بود: {exc}") from exc
        blocks: list[tuple[str, str]] = []
        for index, paragraph in enumerate(document.paragraphs, start=1):
            text = paragraph.text.strip()
            if text:
                style = getattr(paragraph.style, "name", "") or "متن"
                blocks.append((f"بند {index} / {style}", text))
        tables: list[ExtractedTable] = []
        for index, table in enumerate(document.tables, start=1):
            raw = [[cell.text.strip() for cell in row.cells] for row in table.rows]
            normalized = _normalize_table_rows(raw, self.max_rows_per_table)
            if normalized:
                headers, rows, truncated = normalized
                tables.append(ExtractedTable(source, f"جدول Word {index}", f"جدول {index}", headers, rows, truncated))
        warnings: list[str] = []
        if not blocks and not tables:
            warnings.append("متن یا جدول قابل استخراج در سند Word پیدا نشد.")
        return ExtractedDocument(source, "word", source.stem, blocks, tables, warnings=warnings)

    def _read_pptx(self, source: Path) -> ExtractedDocument:
        try:
            from pptx import Presentation
        except ImportError as exc:
            raise EngineError("برای خواندن PowerPoint، کتابخانه python-pptx نصب نیست.") from exc
        try:
            presentation = Presentation(str(source))
        except Exception as exc:  # noqa: BLE001
            raise EngineError(f"خواندن PowerPoint ناموفق بود: {exc}") from exc
        blocks: list[tuple[str, str]] = []
        tables: list[ExtractedTable] = []
        warnings: list[str] = []
        slides = list(presentation.slides)
        for slide_number, slide in enumerate(slides[: self.max_pages], start=1):
            title = ""
            for shape in slide.shapes:
                if getattr(shape, "has_text_frame", False):
                    text = shape.text.strip()
                    if text:
                        if getattr(shape, "is_placeholder", False) and getattr(shape.placeholder_format, "type", None) == 1:
                            title = text
                        blocks.append((f"اسلاید {slide_number}" + (f" / {title}" if title else ""), text))
                if getattr(shape, "has_table", False):
                    raw = [[cell.text.strip() for cell in row.cells] for row in shape.table.rows]
                    normalized = _normalize_table_rows(raw, self.max_rows_per_table)
                    if normalized:
                        headers, rows, truncated = normalized
                        tables.append(ExtractedTable(source, f"اسلاید {slide_number}", f"جدول اسلاید {slide_number}", headers, rows, truncated))
            try:
                notes = slide.notes_slide.notes_text_frame.text.strip()
                if notes:
                    blocks.append((f"یادداشت سخنران اسلاید {slide_number}", notes))
            except Exception:  # noqa: BLE001
                pass
        if len(slides) > self.max_pages:
            warnings.append(f"تعداد اسلایدها از سقف {self.max_pages} بیشتر بود؛ ادامه بررسی نشد.")
        if not blocks and not tables:
            warnings.append("متن یا جدول قابل استخراج در اسلایدها پیدا نشد.")
        return ExtractedDocument(source, "powerpoint", source.stem, blocks, tables, slide_count=len(slides), warnings=warnings)

    def _read_pdf(self, source: Path) -> ExtractedDocument:
        try:
            import pdfplumber
        except ImportError as exc:
            raise EngineError("برای خواندن PDF، کتابخانه pdfplumber نصب نیست.") from exc
        blocks: list[tuple[str, str]] = []
        tables: list[ExtractedTable] = []
        warnings: list[str] = []
        total_pages = 0
        try:
            with pdfplumber.open(str(source)) as pdf:
                total_pages = len(pdf.pages)
                pages = pdf.pages[: self.max_pages]
                for page_number, page in enumerate(pages, start=1):
                    text = (page.extract_text(layout=False) or "").strip()
                    if text:
                        blocks.append((f"صفحه {page_number}", text[:20_000]))
                    try:
                        page_tables = page.extract_tables() or []
                    except Exception:  # noqa: BLE001
                        page_tables = []
                    for table_index, raw in enumerate(page_tables, start=1):
                        normalized = _normalize_table_rows(raw, self.max_rows_per_table)
                        if normalized:
                            headers, rows, truncated = normalized
                            tables.append(ExtractedTable(source, f"صفحه {page_number}", f"جدول صفحه {page_number} - {table_index}", headers, rows, truncated))
                    if not text and not page_tables:
                        # OCR only when the PDF page is image-based and the optional OCR stack is ready.
                        try:
                            image = page.to_image(resolution=150).original
                            ocr_text = self._ocr_pil_image(image)
                            if ocr_text.strip():
                                blocks.append((f"صفحه {page_number} / OCR", ocr_text[:20_000]))
                        except Exception as exc:  # noqa: BLE001
                            warnings.append(f"صفحه {page_number} متن قابل استخراج نداشت؛ OCR در دسترس نبود: {exc}")
                total_pages = len(pdf.pages)
        except Exception as exc:  # noqa: BLE001
            raise EngineError(f"خواندن PDF ناموفق بود: {exc}") from exc
        if total_pages > self.max_pages:
            warnings.append(f"فقط {self.max_pages} صفحه از {total_pages} صفحه بررسی شد.")
        if not blocks and not tables:
            warnings.append("متن قابل استخراج در PDF پیدا نشد؛ ممکن است سند اسکن‌شده یا رمزگذاری‌شده باشد.")
        return ExtractedDocument(source, "pdf", source.stem, blocks, tables, page_count=total_pages, warnings=warnings)

    def _read_image(self, source: Path) -> ExtractedDocument:
        try:
            from PIL import Image, ImageOps
        except ImportError as exc:
            raise EngineError("برای OCR تصویر، کتابخانه Pillow نصب نیست.") from exc
        try:
            with Image.open(source) as original:
                image = ImageOps.exif_transpose(original).convert("RGB")
                image.thumbnail((3000, 3000))
                text = self._ocr_pil_image(image)
        except EngineError:
            raise
        except Exception as exc:  # noqa: BLE001
            raise EngineError(f"بازکردن تصویر ناموفق بود: {exc}") from exc
        warnings = []
        if not text.strip():
            warnings.append("متنی از تصویر استخراج نشد؛ کیفیت تصویر یا زبان OCR را بررسی کنید.")
        return ExtractedDocument(source, "image_ocr", source.stem, [("OCR تصویر", text)], warnings=warnings)

    def _ocr_pil_image(self, image: Any) -> str:
        try:
            import pytesseract
        except ImportError as exc:
            raise EngineError("برای OCR، pytesseract نصب نیست.") from exc
        configured = os.environ.get("TESSERACT_CMD", "").strip()
        candidates = [
            configured,
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            str(self.settings.project_root / "tools" / "Tesseract-OCR" / "tesseract.exe"),
        ]
        executable = next((item for item in candidates if item and Path(item).is_file()), None)
        if executable:
            pytesseract.pytesseract.tesseract_cmd = executable
        elif not shutil.which("tesseract"):
            raise EngineError("موتور Tesseract در ویندوز نصب یا در PATH ثبت نشده است.")
        try:
            languages = set(pytesseract.get_languages(config=""))
        except Exception as exc:  # noqa: BLE001
            raise EngineError(f"Tesseract اجرا نشد: {exc}") from exc
        if "fas" not in languages:
            raise EngineError("زبان فارسی Tesseract (fas.traineddata) نصب نیست؛ OCR فارسی قابل اجرا نیست.")
        lang = "fas+eng" if "eng" in languages else "fas"
        try:
            return pytesseract.image_to_string(image, lang=lang, config="--psm 6")
        except Exception as exc:  # noqa: BLE001
            raise EngineError(f"استخراج OCR ناموفق بود: {exc}") from exc


def _normalize_table_rows(raw: list[list[Any]] | None, limit: int) -> tuple[list[str], list[list[Any]], bool] | None:
    if not raw:
        return None
    cleaned = [[str(value).strip() if value is not None else "" for value in row] for row in raw]
    cleaned = [row for row in cleaned if any(cell for cell in row)]
    if not cleaned:
        return None
    width = max(len(row) for row in cleaned)
    cleaned = [row + [""] * (width - len(row)) for row in cleaned]
    first = cleaned[0]
    headers = [cell or f"ستون {index + 1}" for index, cell in enumerate(first)]
    rows = cleaned[1:]
    truncated = len(rows) > limit
    return headers, rows[:limit], truncated
