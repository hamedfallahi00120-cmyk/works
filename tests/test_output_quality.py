"""Output-quality contract: clean delivery, monthly integration and folder mode.

These tests use real .xlsx files built by openpyxl; nothing is mocked.
"""

from __future__ import annotations

from hashlib import sha256

import openpyxl
import pytest

from local_office_ai.app.coordinator import Coordinator
from local_office_ai.app.folder_workflow import FolderWorkflow
from local_office_ai.app.workflow import ExcelIntelligenceWorkflow
from local_office_ai.core.settings import load_settings
from local_office_ai.excel.data_qa import answer_request
from local_office_ai.excel.structure_intelligence import expected_days

METRIC = "تعداد کل اعمال قانون"
INTERNAL_SHEETS = {
    "گزارش عملکرد",
    "نمونه‌های پایش",
    "نمونه‌های GPU",
    "داده خام",
    "داده استانداردشده",
    "بلوک‌های استخراج",
    "ردیابی منبع",
    "روش‌شناسی آماری",
    "نتایج محاسبات",
}


MONTH_NUMBERS = {
    "فروردین": 1, "اردیبهشت": 2, "خرداد": 3, "تیر": 4,
    "مرداد": 5, "شهریور": 6, "مهر": 7, "آبان": 8,
}


def _monthly_book(path, months, *, days=None, extra_year=None):
    """A workbook shaped like the real monthly report: two year columns."""

    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)
    for name, year, values in months:
        month_no = MONTH_NUMBERS[name]
        sheet = workbook.create_sheet(f"{name} {year}")
        sheet["B2"] = f"آمار روزانه مقایسه‌ای {name} {year}"
        sheet["B4"] = "ردیف"
        sheet["C4"] = "تاریخ"
        sheet["D4"] = METRIC
        sheet["E4"] = METRIC
        # The real reports merge the metric title over its two year columns;
        # the multi-row header is what makes the year band detectable.
        sheet.merge_cells("D4:E4")
        sheet["D5"] = str(int(year) - 1)
        sheet["E5"] = str(year)
        month_days = expected_days(month_no) or 30
        count = month_days if days is None else min(days, month_days)
        for offset in range(1, count + 1):
            value = values[(offset - 1) % len(values)]
            sheet.cell(5 + offset, 2, offset)
            sheet.cell(5 + offset, 3, offset)
            sheet.cell(5 + offset, 4, value)
            sheet.cell(5 + offset, 5, value + 100)
        total_row = 5 + count + 1
        sheet.cell(total_row, 2, "جمع")
        for column in (4, 5):
            letter = openpyxl.utils.get_column_letter(column)
            sheet.cell(total_row, column, f"=SUM({letter}6:{letter}{total_row - 1})")
    for name, rows in (extra_year or {}).items():
        sheet = workbook.create_sheet(name)
        for text in rows:
            sheet.append(text)
    workbook.save(path)
    return path


@pytest.fixture()
def monthly_workbook(tmp_path):
    return _monthly_book(
        tmp_path / "ماهانه.xlsx",
        [
            ("فروردین", "1404", [10]),
            ("اردیبهشت", "1404", [20]),
            ("خرداد", "1404", [30]),
        ],
    )


@pytest.fixture()
def folder_settings(tmp_path):
    root = tmp_path / "project"
    (root / "config").mkdir(parents=True)
    (root / "data").mkdir()
    (root / "logs").mkdir()
    return load_settings(project_root=root)


def test_clean_delivery_has_no_internal_sheets(monthly_workbook, folder_settings, tmp_path):
    output = tmp_path / "clean.xlsx"
    result = Coordinator(folder_settings, console=False).run(
        monthly_workbook,
        use_ai=False,
        output=output,
        request="جمع ماهانه و بیشترین ماه را بگو",
        profile="clean",
    )
    workbook = openpyxl.load_workbook(result.output_path)
    names = set(workbook.sheetnames)
    assert not names & INTERNAL_SHEETS
    assert "پاسخ درخواست" in names
    assert "جدول ماهانه شاخص‌ها" in names
    workbook.close()


def test_audit_profile_keeps_traceability(monthly_workbook, folder_settings, tmp_path):
    output = tmp_path / "audit.xlsx"
    result = Coordinator(folder_settings, console=False).run(
        monthly_workbook, use_ai=False, output=output, profile="audit"
    )
    workbook = openpyxl.load_workbook(result.output_path)
    assert {"داده خام", "ردیابی منبع", "گزارش عملکرد", "بلوک‌های استخراج"} <= set(
        workbook.sheetnames
    )
    workbook.close()


def test_monthly_totals_and_ranking_come_from_python(monthly_workbook, folder_settings):
    analysis = ExcelIntelligenceWorkflow(folder_settings).analyze_only(
        monthly_workbook, request="مقایسه 1404 با 1403 و بیشترین ماه"
    )
    integration = analysis.integration
    assert integration is not None
    series = next(item for item in integration.series if item.metric == METRIC)
    # 31 days per month: D is 1403 (10 + 20 + 30 per day), E adds 100 per row.
    assert series.total("1403") == 31 * 60
    assert series.total("1404") == 31 * 60 + 31 * 300
    best = series.best("1404")
    assert best.month == "خرداد" and best.total == 31 * 130

    answers = answer_request(analysis.request_intent, analysis, integration)
    ranking = next(item for item in answers if item.kind == "ranking" and METRIC in item.title)
    assert "خرداد" in ranking.text and "4,030" in ranking.text


def test_incomplete_month_is_not_counted_in_the_period_total(tmp_path, folder_settings):
    path = _monthly_book(tmp_path / "ناقص.xlsx", [("فروردین", "1404", [10])])
    workbook = openpyxl.load_workbook(path)
    sheet = workbook["فروردین 1404"]
    for row in (15, 16, 17):  # three days of the month have no measurement
        sheet.cell(row, 4).value = None
        sheet.cell(row, 5).value = None
    workbook.save(path)
    workbook.close()

    analysis = ExcelIntelligenceWorkflow(folder_settings).analyze_only(path)
    series = next(item for item in analysis.integration.series if item.metric == METRIC)
    cell = next(item for item in series.cells if item.period == "1404")
    assert cell.total == 28 * 110  # 28 measured days: kept, but flagged incomplete
    assert cell.complete is False
    assert series.total("1404") is None  # incomplete months never enter the total


def test_folder_mode_compares_files_without_touching_them(tmp_path, folder_settings):
    first = _monthly_book(
        tmp_path / "الف.xlsx",
        [("فروردین", "1404", [10]), ("اردیبهشت", "1404", [20])],
    )
    second = _monthly_book(
        tmp_path / "ب.xlsx",
        [("فروردین", "1403", [5]), ("اردیبهشت", "1403", [7])],
    )
    before = {path.name: sha256(path.read_bytes()).hexdigest() for path in (first, second)}

    workflows = FolderWorkflow(folder_settings, progress=lambda message: None)
    result = workflows.run(
        tmp_path, request="جمع شاخص‌های هر فایل را مقایسه کن", use_ai=False
    )

    assert result.status == "completed"
    assert {item.name for item in result.files} == {"الف.xlsx", "ب.xlsx"}
    assert METRIC in result.metrics
    workbook = openpyxl.load_workbook(result.output_path)
    assert {"خلاصه پوشه", "فهرست فایل‌ها", "مقایسه بین فایل‌ها"} <= set(workbook.sheetnames)
    comparison = workbook["مقایسه بین فایل‌ها"]
    values = [
        [cell.value for cell in row]
        for row in comparison.iter_rows(min_row=4, max_row=6)
    ]
    totals = {row[0]: row[1:4] for row in values if row and row[0] == METRIC}
    assert totals, "سربرگ شاخص در جدول مقایسه نیست"
    summary = {
        row[0].value: row[1].value
        for row in workbook["خلاصه پوشه"].iter_rows(min_row=4, max_row=20)
        if row and row[0].value
    }
    assert summary.get("درخواست کاربر") == "جمع شاخص‌های هر فایل را مقایسه کن"
    assert summary.get("تفسیر درخواست"), "تفسیر درخواست در خلاصه پوشه نوشته نشد"
    answer = summary.get("پاسخ درخواست") or ""
    assert METRIC in answer and "بیشترین" in answer and "کمترین" in answer
    workbook.close()
    after = {path.name: sha256(path.read_bytes()).hexdigest() for path in (first, second)}
    assert after == before


def test_clean_profile_keeps_model_narrative_out_of_a_technical_sheet(
    tmp_path, ollama_settings
):
    """A model run must answer inside the report, never through an AI/execution sheet."""
    from local_office_ai.app.coordinator import Coordinator

    book = _monthly_book(
        tmp_path / "دو سال.xlsx",
        [("فروردین", "1404", [10, 20]), ("فروردین", "1403", [7, 9])],
    )
    coordinator = Coordinator(ollama_settings, console=False)
    result = coordinator.run(
        book,
        model="qwen3.5:9b",
        use_ai=True,
        request="مقایسه 1404 با 1403 و بیشترین ماه",
        output=tmp_path / "خروجی.xlsx",
        profile="clean",
    )

    workbook = openpyxl.load_workbook(result.output_path)
    names = workbook.sheetnames
    assert "تحلیل هوش مصنوعی" not in names
    assert not [name for name in names if "پایش" in name or "GPU" in name or "عملکرد" in name]
    summary = workbook["خلاصه اجرایی"]
    rows = [
        (row[0].value, row[1].value)
        for row in summary.iter_rows(min_row=4, max_row=80)
        if row[0].value
    ]
    labels = {str(label): value for label, value in rows}
    assert labels.get("مدل") == "qwen3.5:9b"
    texts = [str(row[0].value) for row in summary.iter_rows(min_row=4, max_row=140)]
    assert any("خلاصه مدیریتی" in text for text in texts)
    workbook.close()


def test_scope_answer_never_reports_zero_months(make_workbook, folder_settings):
    path = make_workbook(
        "کارکنان.xlsx",
        {
            "کارکنان": [
                ["نام", "واحد", "تاریخ استخدام", "حقوق"],
                ["الف", "فناوری", "1403/01/05", 100],
                ["ب", "مالی", "1403/02/11", 200],
                ["پ", "فناوری", "1403/03/09", 300],
            ]
        },
    )
    analysis = ExcelIntelligenceWorkflow(folder_settings).analyze_only(
        path, request="تحلیل کلی و مقایسه دوره‌ها"
    )
    answers = answer_request(analysis.request_intent, analysis, analysis.integration)
    scope = next(item for item in answers if item.kind == "scope")
    assert "0 ماه" not in scope.text
    assert "0 ماه" not in scope.value
