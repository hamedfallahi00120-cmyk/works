from __future__ import annotations

from datetime import datetime
from pathlib import Path

import openpyxl
import pytest

from local_office_ai.ai.analyzer import AIAnalyzer
from local_office_ai.core.exceptions import EngineError
from local_office_ai.excel.reader import file_sha256
from local_office_ai.excel.report_writer import ExcelReportWriter, ReportContext
from local_office_ai.monitoring.performance_report import build_performance_report
from local_office_ai.monitoring.resource_monitor import ResourceMonitor


@pytest.fixture()
def report_context(settings, sample_path):
    from local_office_ai.app.workflow import ExcelIntelligenceWorkflow

    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(sample_path)
    ai = AIAnalyzer(settings).analyze(analysis, use_ai=False)
    monitor = ResourceMonitor(settings)
    monitor.limitations.append("محدودیت نمونه")
    report = build_performance_report(
        monitor,
        settings,
        started_at=datetime(2026, 1, 1, 9, 0, 0),
        finished_at=datetime(2026, 1, 1, 9, 0, 3),
        total_seconds=3.0,
        timings={"خواندن فایل": 0.2},
        ai_metadata=ai.metadata,
        input_sha256=analysis.meta.sha256,
    )
    return ReportContext(
        analysis=analysis,
        ai=ai,
        performance=report,
        output_path=Path(settings.output_dir) / "report.xlsx",
        monitor=monitor,
    )


def test_report_structure_and_formatting(settings, report_context, sample_path):
    before = file_sha256(sample_path)
    output = ExcelReportWriter(settings).build(report_context)
    assert output.exists()
    assert not output.with_suffix(".tmp.xlsx").exists()
    assert file_sha256(sample_path) == before

    workbook = openpyxl.load_workbook(output)
    expected = {
        "داشبورد",
        "خلاصه اجرایی",
        "داده استانداردشده",
        "نتایج محاسبات",
        "حسابرسی اعداد",
        "تطبیق تعهدات",
        "موارد نیازمند بررسی",
        "تحلیل هوش مصنوعی",
        "ردیابی منبع",
        "گزارش عملکرد",
        "محدودیت‌ها و خطاها",
        "ساختار فایل",
    }
    assert expected.issubset(set(workbook.sheetnames))
    for name in expected:
        worksheet = workbook[name]
        assert worksheet.sheet_view.rightToLeft is True
        assert worksheet.max_row >= 1
    dashboard = workbook["داشبورد"]
    assert dashboard.freeze_panes
    assert len(dashboard._charts) == 1
    workbook.close()


def test_dashboard_numbers_and_status(settings, report_context):
    output = ExcelReportWriter(settings).build(report_context)
    workbook = openpyxl.load_workbook(output, data_only=True)
    dashboard = workbook["داشبورد"]
    values = {
        str(dashboard.cell(row, 1).value): dashboard.cell(row, 2).value
        for row in range(6, 12)
        if dashboard.cell(row, 1).value
    }
    assert values["مجموع اصل وام‌ها"] == 325000000
    assert values["مجموع اقساط ماهانه وام‌ها"] == 14749000
    assert values["مجموع تعهدات ماهانه"] == 18999000
    assert values["مجموع بازپرداخت قراردادی"] == 347576000
    # Status cell for a verified audit is filled green.
    audit_row = next(
        row
        for row in range(1, dashboard.max_row + 1)
        if dashboard.cell(row, 1).value == "مجموع اصل وام‌ها" and row > 12
    )
    assert dashboard.cell(audit_row, 5).value == "تأیید شد"
    assert dashboard.cell(audit_row, 5).fill.fgColor.rgb.endswith("E2EFDA")
    workbook.close()


def test_normalized_sheet_has_numeric_values(settings, report_context):
    output = ExcelReportWriter(settings).build(report_context)
    workbook = openpyxl.load_workbook(output, data_only=True)
    sheet = workbook["داده استانداردشده"]
    headers = [sheet.cell(3, column).value for column in range(1, 11)]
    assert headers[0] == "شماره ردیف"
    principal_column = headers.index("مبلغ وام") + 1
    assert isinstance(sheet.cell(4, principal_column).value, (int, float))
    date_column = headers.index("تاریخ آخرین قسط") + 1
    assert str(sheet.cell(4, date_column).value).startswith("14")
    workbook.close()


def test_anomaly_and_ai_sheets(settings, report_context):
    output = ExcelReportWriter(settings).build(report_context)
    workbook = openpyxl.load_workbook(output, data_only=True)
    anomalies = workbook["موارد نیازمند بررسی"]
    text = " ".join(str(anomalies.cell(row, 3).value or "") for row in range(1, anomalies.max_row + 1))
    assert "خارج از جدول" in text
    ai_sheet = workbook["تحلیل هوش مصنوعی"]
    mode_value = None
    for row in range(1, ai_sheet.max_row + 1):
        if ai_sheet.cell(row, 1).value == "حالت تحلیل":
            mode_value = ai_sheet.cell(row, 2).value
    assert mode_value and "بدون مدل" in str(mode_value)
    workbook.close()


def test_writer_rejects_unwritable_target(settings, report_context):
    report_context.output_path = Path("/proc/definitely/not/writable/report.xlsx")
    with pytest.raises((EngineError, OSError)):
        ExcelReportWriter(settings).build(report_context)
