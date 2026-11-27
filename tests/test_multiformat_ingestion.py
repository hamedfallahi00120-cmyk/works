from __future__ import annotations

from pathlib import Path

import pytest

from local_office_ai.core.settings import Settings
from local_office_ai.excel.mixed_report_writer import MixedReportWriter
from local_office_ai.ingestion.reader import DocumentReader
from local_office_ai.ingestion.statistics import profile_table
from local_office_ai.app.document_set_workflow import DocumentSetWorkflow


def _settings(tmp_path: Path) -> Settings:
    paths = {
        "project_root": tmp_path,
        "input_dir": tmp_path / "data" / "input",
        "processed_dir": tmp_path / "data" / "processed",
        "output_dir": tmp_path / "data" / "output",
        "temp_dir": tmp_path / "data" / "temp",
        "log_dir": tmp_path / "logs",
    }
    for path in paths.values():
        if isinstance(path, Path) and path.suffix == "":
            path.mkdir(parents=True, exist_ok=True)
    return Settings(**paths)


def test_csv_persian_numbers_and_identifier_date_exclusion(tmp_path: Path) -> None:
    settings = _settings(tmp_path)
    source = tmp_path / "daily.csv"
    source.write_text(
        "نام,تعداد,تاریخ,شناسه\nالف,۱۲,۱۴۰۴/۰۱/۰۱,۱۰۰۱\nب,۱۸,۱۴۰۴/۰۱/۰۲,۱۰۰۲\nج,۲۰,۱۴۰۴/۰۱/۰۳,۱۰۰۳\n",
        encoding="utf-8-sig",
    )
    document = DocumentReader(settings).read(source)
    assert document.kind == "csv"
    assert len(document.tables) == 1
    profiles = profile_table(
        source.name,
        "CSV",
        source.stem,
        document.tables[0].headers,
        document.tables[0].rows,
    )
    assert [item["column"] for item in profiles] == ["تعداد"]
    assert profiles[0]["sum"] == 50



def test_txt_reader_preserves_persian_text(tmp_path: Path) -> None:
    settings = _settings(tmp_path)
    source = tmp_path / "notes.txt"
    source.write_text("گزارش ماهانه\nتعداد فعالیت‌ها: ۱۲۵", encoding="utf-8")
    document = DocumentReader(settings).read(source)
    assert document.kind == "text"
    assert "۱۲۵" in document.compact_text()

def test_docx_and_pptx_text_and_tables(tmp_path: Path) -> None:
    pytest.importorskip("docx")
    pytest.importorskip("pptx")
    from docx import Document
    from pptx import Presentation
    from pptx.util import Inches

    settings = _settings(tmp_path)
    docx_path = tmp_path / "memo.docx"
    document = Document()
    document.add_heading("گزارش اداری", level=1)
    document.add_paragraph("خلاصه عملکرد واحدها")
    table = document.add_table(rows=1, cols=2)
    table.rows[0].cells[0].text = "واحد"
    table.rows[0].cells[1].text = "تعداد"
    row = table.add_row().cells
    row[0].text, row[1].text = "الف", "۱۲"
    document.save(docx_path)

    pptx_path = tmp_path / "brief.pptx"
    presentation = Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[5])
    slide.shapes.title.text = "گزارش ماهانه"
    text_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(5), Inches(1))
    text_box.text = "تغییرات عملکرد"
    table = slide.shapes.add_table(2, 2, Inches(1), Inches(3), Inches(5), Inches(1)).table
    table.cell(0, 0).text, table.cell(0, 1).text = "ماه", "تعداد"
    table.cell(1, 0).text, table.cell(1, 1).text = "اول", "۱۵"
    presentation.save(pptx_path)

    reader = DocumentReader(settings)
    word = reader.read(docx_path)
    powerpoint = reader.read(pptx_path)
    assert word.kind == "word" and word.tables and word.text_blocks
    assert powerpoint.kind == "powerpoint" and powerpoint.tables and powerpoint.text_blocks
    assert powerpoint.slide_count == 1


def test_pdf_text_and_recursive_scan(tmp_path: Path) -> None:
    pytest.importorskip("pdfplumber")
    reportlab = pytest.importorskip("reportlab.pdfgen.canvas")
    settings = _settings(tmp_path)
    nested = tmp_path / "input" / "nested"
    nested.mkdir(parents=True, exist_ok=True)
    pdf_path = nested / "note.pdf"
    canvas = reportlab.Canvas(str(pdf_path))
    canvas.drawString(72, 750, "Administrative performance report")
    canvas.save()
    (nested / "notes.txt").write_text("گزارش متنی", encoding="utf-8")
    (nested / "~$locked.docx").write_text("temporary", encoding="utf-8")
    (tmp_path / "input" / "reference").mkdir()
    (tmp_path / "input" / "reference" / "ignore.csv").write_text("a,b\n1,2", encoding="utf-8")

    reader = DocumentReader(settings)
    pdf = reader.read(pdf_path)
    assert pdf.kind == "pdf" and pdf.text_blocks
    scanned = DocumentSetWorkflow.scan(tmp_path / "input", recursive=True)
    assert pdf_path.resolve() in scanned
    assert (nested / "notes.txt").resolve() in scanned
    assert not any("reference" in str(path) or path.name.startswith("~$") for path in scanned)


def test_mixed_workflow_creates_clean_workbook(tmp_path: Path) -> None:
    pytest.importorskip("docx")
    pytest.importorskip("pptx")
    pytest.importorskip("pdfplumber")
    pytest.importorskip("reportlab.pdfgen.canvas")
    from docx import Document
    from pptx import Presentation
    from pptx.util import Inches
    from openpyxl import Workbook, load_workbook
    from reportlab.pdfgen import canvas

    settings = _settings(tmp_path)
    input_dir = settings.input_dir
    xlsx = input_dir / "counts.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "آمار"
    sheet.append(["شاخص", "تعداد"])
    sheet.append(["الف", 12])
    sheet.append(["ب", 18])
    workbook.save(xlsx)

    csv_path = input_dir / "daily.csv"
    csv_path.write_text("نام,تعداد\nالف,۱۲\nب,۱۸\nج,۲۰\n", encoding="utf-8-sig")
    word_path = input_dir / "memo.docx"
    word = Document()
    word.add_paragraph("گزارش عملکرد واحدها")
    word.save(word_path)
    ppt_path = input_dir / "brief.pptx"
    presentation = Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[5])
    slide.shapes.title.text = "گزارش"
    text_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(5), Inches(1))
    text_box.text = "نتایج فعالیت"
    presentation.save(ppt_path)
    pdf_path = input_dir / "note.pdf"
    pdf = canvas.Canvas(str(pdf_path))
    pdf.drawString(72, 750, "Total activities: 25")
    pdf.save()

    result = DocumentSetWorkflow(settings).run(
        paths=[xlsx, csv_path, word_path, ppt_path, pdf_path],
        request="منابع را خلاصه کن و شاخص‌های عددی معتبر را نشان بده.",
        use_ai=False,
    )
    assert result.status == "completed"
    assert len([item for item in result.files if item.status == "تحلیل شد"]) == 5
    assert Path(result.output_path).is_file()
    output = load_workbook(result.output_path, read_only=True)
    names = set(output.sheetnames)
    assert "خلاصه اجرایی" in names and "فهرست منابع" in names
    assert not any(token in name.lower() for name in names for token in ("gpu", "cpu", "ram", "test", "performance", "monitor"))
    output.close()


def test_scanned_pdf_uses_ocr_provider(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    pytest.importorskip("pdfplumber")
    pytest.importorskip("pypdfium2")
    pytest.importorskip("PIL")
    canvas_module = pytest.importorskip("reportlab.pdfgen.canvas")
    from PIL import Image
    from reportlab.lib.utils import ImageReader

    settings = _settings(tmp_path)
    image_path = tmp_path / "scan.png"
    Image.new("RGB", (500, 300), "white").save(image_path)
    pdf_path = tmp_path / "scanned.pdf"
    canvas = canvas_module.Canvas(str(pdf_path))
    canvas.drawImage(ImageReader(str(image_path)), 0, 0, width=500, height=300)
    canvas.save()

    reader = DocumentReader(settings)
    monkeypatch.setattr(reader, "_ocr_pil_image", lambda image: "متن OCR آزمایشی")
    result = reader.read(pdf_path)
    assert any("OCR" in location and "متن OCR آزمایشی" in text for location, text in result.text_blocks)
