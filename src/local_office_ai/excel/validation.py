"""Data validation: an explicit pass/fail checklist for the analysed workbook.

The validator does not compute anything and never repairs a value.  It states,
for every quality rule, whether the analysed data satisfies it, how many rows or
cells are affected, and where.  The report prints this checklist so an operator
can see what was verified, not only what went wrong.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Any

from local_office_ai.core.logging_config import get_logger
from local_office_ai.core.models import AnalysisResult, ColumnRole, TableKind
from local_office_ai.core.settings import Settings
from local_office_ai.excel.sheet_analyzer import SheetGrid

LOGGER = get_logger(__name__)

PASSED = "passed"
WARNING = "warning"
FAILED = "failed"
SKIPPED = "skipped"

STATUS_LABELS = {
    PASSED: "قبول",
    WARNING: "هشدار",
    FAILED: "رد",
    SKIPPED: "بررسی نشد",
}


@dataclass
class ValidationCheck:
    code: str
    label: str
    scope: str
    status: str
    affected: int = 0
    detail: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "code": self.code,
            "label": self.label,
            "scope": self.scope,
            "status": self.status,
            "status_label": STATUS_LABELS.get(self.status, self.status),
            "affected": self.affected,
            "detail": self.detail,
        }


@dataclass
class ValidationSummary:
    checks: list[ValidationCheck] = field(default_factory=list)

    @property
    def status(self) -> str:
        statuses = {check.status for check in self.checks}
        if FAILED in statuses:
            return FAILED
        if WARNING in statuses:
            return WARNING
        return PASSED

    @property
    def failed(self) -> list[ValidationCheck]:
        return [check for check in self.checks if check.status is FAILED]

    @property
    def warnings(self) -> list[ValidationCheck]:
        return [check for check in self.checks if check.status is WARNING]

    def to_dict(self) -> dict[str, Any]:
        return {
            "status": self.status,
            "status_label": STATUS_LABELS.get(self.status, self.status),
            "checks": [check.to_dict() for check in self.checks],
        }


def overall_status(checks: list[ValidationCheck]) -> str:
    return ValidationSummary(checks).status


class DataValidator:
    def __init__(self, settings: Settings) -> None:
        self.settings = settings

    # ------------------------------------------------------------------ #
    def validate(
        self,
        analysis: AnalysisResult,
        grids: dict[str, SheetGrid] | None = None,
    ) -> list[ValidationCheck]:
        grids = grids or {}
        checks = [
            self._file_checks(analysis),
            self._structure_checks(analysis),
            self._column_checks(analysis),
            self._numeric_checks(analysis),
            self._cell_checks(analysis, grids),
            self._integrity_checks(analysis),
            self._audit_checks(analysis),
            self._loan_checks(analysis),
            self._reconciliation_checks(analysis),
        ]
        LOGGER.info(
            "validation: %d checks, overall=%s",
            len(checks),
            overall_status(checks),
        )
        return checks

    # ------------------------------------------------------------------ #
    @staticmethod
    def _file_checks(analysis: AnalysisResult) -> ValidationCheck:
        sheets = len(analysis.meta.sheets)
        return ValidationCheck(
            code="file-sheets",
            label="برگه‌های خوانده‌شده",
            scope=Path(analysis.meta.path).name,
            status=PASSED if sheets else FAILED,
            affected=sheets,
            detail=f"{sheets} برگه خوانده شد." if sheets else "هیچ برگه‌ای خوانده نشد.",
        )

    @staticmethod
    def _structure_checks(analysis: AnalysisResult) -> ValidationCheck:
        primary = analysis.primary_table
        if primary is None:
            return ValidationCheck(
                code="primary-table",
                label="جدول اصلی",
                scope=Path(analysis.meta.path).name,
                status=FAILED,
                detail="جدول داده‌ای با ساختار قابل اتکا پیدا نشد.",
            )
        confidence = analysis.structure_confidence
        if confidence >= 0.6:
            status, detail = PASSED, f"اطمینان تشخیص ساختار {confidence:.0%} است."
        elif confidence >= 0.4:
            status, detail = WARNING, f"اطمینان تشخیص ساختار فقط {confidence:.0%} است؛ بازبینی چشمی لازم است."
        else:
            status, detail = FAILED, f"اطمینان تشخیص ساختار {confidence:.0%} است؛ ساختار تأیید نشده است."
        return ValidationCheck(
            code="structure-confidence",
            label="اطمینان تشخیص ساختار",
            scope=f"{primary.sheet}!{primary.address}",
            status=status,
            affected=primary.row_count,
            detail=detail,
        )

    @staticmethod
    def _column_checks(analysis: AnalysisResult) -> ValidationCheck:
        profiles = analysis.column_profiles
        if not profiles:
            return ValidationCheck(
                code="column-profile",
                label="پروفایل ستون‌ها",
                scope=Path(analysis.meta.path).name,
                status=SKIPPED,
                detail="ستونی برای پروفایل‌سازی وجود نداشت.",
            )
        duplicates = [p.header for p in profiles if p.header and sum(1 for q in profiles if q.header == p.header) > 1]
        if duplicates:
            return ValidationCheck(
                code="header-uniqueness",
                label="یکتایی عنوان ستون‌ها",
                scope=analysis.primary_table.sheet if analysis.primary_table else "",
                status=WARNING,
                affected=len(set(duplicates)),
                detail="عنوان تکراری: " + "، ".join(sorted(set(duplicates))[:6]),
            )
        empty_headers = [p.letter for p in profiles if not p.header]
        if empty_headers:
            return ValidationCheck(
                code="header-uniqueness",
                label="یکتایی و وجود عنوان ستون‌ها",
                scope=analysis.primary_table.sheet if analysis.primary_table else "",
                status=WARNING,
                affected=len(empty_headers),
                detail="ستون بدون عنوان: " + "، ".join(empty_headers[:8]),
            )
        return ValidationCheck(
            code="header-uniqueness",
            label="یکتایی عنوان ستون‌ها",
            scope=analysis.primary_table.sheet if analysis.primary_table else "",
            status=PASSED,
            affected=len(profiles),
            detail=f"{len(profiles)} ستون با عنوان یکتا.",
        )

    @staticmethod
    def _numeric_checks(analysis: AnalysisResult) -> ValidationCheck:
        numeric_roles = {ColumnRole.PRINCIPAL, ColumnRole.INSTALLMENT, ColumnRole.AMOUNT}
        numeric = [profile for profile in analysis.column_profiles if profile.role in numeric_roles]
        if not numeric:
            return ValidationCheck(
                code="numeric-integrity",
                label="سلامت مقادیر عددی",
                scope=Path(analysis.meta.path).name,
                status=SKIPPED,
                detail="ستون عددی با نقش مالی شناسایی نشد.",
            )
        weak = [profile for profile in numeric if profile.numeric_ratio < 0.8]
        if weak:
            return ValidationCheck(
                code="numeric-integrity",
                label="سلامت مقادیر عددی",
                scope=Path(analysis.meta.path).name,
                status=WARNING,
                affected=len(weak),
                detail="ستون با نسبت عددی کمتر از ۸۰٪: "
                + "، ".join(f"{p.header or p.letter} ({p.numeric_ratio:.0%})" for p in weak[:5]),
            )
        return ValidationCheck(
            code="numeric-integrity",
            label="سلامت مقادیر عددی",
            scope=Path(analysis.meta.path).name,
            status=PASSED,
            affected=len(numeric),
            detail=f"{len(numeric)} ستون مالی با مقادیر عددی معتبر.",
        )

    @staticmethod
    def _integrity_checks(analysis: AnalysisResult) -> ValidationCheck:
        meta = analysis.meta
        issues: list[str] = []
        if meta.total_error_cells:
            issues.append(f"{meta.total_error_cells} سلول خطا")
        if meta.has_macros:
            issues.append("فایل دارای ماکرو است و ماکرو اجرا نشد")
        hidden_sheets = [sheet.title for sheet in meta.sheets if sheet.state != "visible"]
        if hidden_sheets:
            issues.append(f"{len(hidden_sheets)} برگه مخفی")
        if meta.external_links:
            issues.append(f"{meta.external_links} پیوند خارجی")
        if issues:
            return ValidationCheck(
                code="workbook-integrity",
                label="سلامت فایل و محتوای مخفی",
                scope=Path(analysis.meta.path).name,
                status=WARNING,
                affected=len(issues),
                detail="؛ ".join(issues),
            )
        return ValidationCheck(
            code="workbook-integrity",
            label="سلامت فایل و محتوای مخفی",
            scope=Path(analysis.meta.path).name,
            status=PASSED,
            affected=meta.total_formulas,
            detail=f"بدون سلول خطا، بدون ماکرو و بدون برگه مخفی؛ {meta.total_formulas} سلول فرمول ثبت شد.",
        )

    def _cell_checks(
        self, analysis: AnalysisResult, grids: dict[str, SheetGrid]
    ) -> ValidationCheck:
        primary = analysis.primary_table
        if primary is None:
            return ValidationCheck(
                code="data-cells",
                label="سلامت سلول‌های داده",
                scope=Path(analysis.meta.path).name,
                status=SKIPPED,
                detail="جدول اصلی برای بررسی سلول‌ها وجود ندارد.",
            )
        grid = grids.get(primary.sheet)
        if grid is None:
            return ValidationCheck(
                code="data-cells",
                label="سلامت سلول‌های داده",
                scope=primary.sheet,
                status=SKIPPED,
                detail="شبکه سلول‌ها در دسترس نبود.",
            )
        merged = 0
        for row, col in grid.merged_covered:
            if (
                primary.first_data_row <= row <= primary.last_data_row
                and primary.first_col <= col <= primary.last_col
            ):
                merged += 1
        if merged:
            return ValidationCheck(
                code="merged-in-data",
                label="سلول ادغام‌شده در محدوده داده",
                scope=f"{primary.sheet}!{primary.address}",
                status=WARNING,
                affected=merged,
                detail=f"{merged} سلول پوشیده‌شده با ادغام در محدوده داده؛ مقدار فقط در لنگر ادغام است.",
            )
        return ValidationCheck(
            code="merged-in-data",
            label="سلول ادغام‌شده در محدوده داده",
            scope=f"{primary.sheet}!{primary.address}",
            status=PASSED,
            detail="هیچ ادغامی داخل محدوده داده نبود.",
        )

    @staticmethod
    def _audit_checks(analysis: AnalysisResult) -> ValidationCheck:
        audits = list(analysis.audits)
        intelligence = analysis.intelligence
        if not audits and intelligence is not None:
            audits = [audit for audit in intelligence.audits if audit.kind == "total"]
        if not audits:
            return ValidationCheck(
                code="declared-totals",
                label="حسابرسی جمع‌های اعلامی فایل",
                scope=Path(analysis.meta.path).name,
                status=SKIPPED,
                detail="فایل جمع اعلامی قابل حسابرسی نداشت.",
            )
        mismatches = [audit for audit in audits if audit.status.value == "mismatch"]
        unverifiable = [audit for audit in audits if audit.status.value == "not_verifiable"]
        if mismatches:
            # A few wrong declared cells do not invalidate the whole file; when
            # the share of mismatching totals stays small this is reported as a
            # warning so the overall status reflects the data, not one cell.
            share = len(mismatches) / max(1, len(audits))
            return ValidationCheck(
                code="declared-totals",
                label="حسابرسی جمع‌های اعلامی فایل",
                scope=Path(analysis.meta.path).name,
                status=FAILED if share >= 0.1 else WARNING,
                affected=len(mismatches),
                detail=(
                    f"{len(mismatches)} از {len(audits)} جمع اعلامی ({share:.1%}) با محاسبه مستقل "
                    "Python اختلاف دارد؛ مقدار محاسبه‌شده مبنای گزارش است. "
                    + "، ".join(
                        f"{getattr(audit, 'label', getattr(audit, 'metric', 'جمع'))} "
                        f"{getattr(audit, 'address', '')} (اختلاف {audit.difference})"
                        for audit in mismatches[:5]
                    )
                ),
            )
        if unverifiable:
            return ValidationCheck(
                code="declared-totals",
                label="حسابرسی جمع‌های اعلامی فایل",
                scope=Path(analysis.meta.path).name,
                status=WARNING,
                affected=len(unverifiable),
                detail=f"{len(unverifiable)} جمع اعلامی قابل بازتولید مستقل نبود.",
            )
        return ValidationCheck(
            code="declared-totals",
            label="حسابرسی جمع‌های اعلامی فایل",
            scope=Path(analysis.meta.path).name,
            status=PASSED,
            affected=len(audits),
            detail=f"{len(audits)} جمع اعلامی با محاسبه مستقل Python مطابق است.",
        )

    def _loan_checks(self, analysis: AnalysisResult) -> ValidationCheck:
        if analysis.primary_table is None or analysis.primary_table.kind is not TableKind.LOAN:
            return ValidationCheck(
                code="loan-fields",
                label="کامل بودن فیلدهای وام",
                scope=Path(analysis.meta.path).name,
                status=SKIPPED,
                detail="جدول اصلی معنای وام ندارد.",
            )
        if not analysis.loans:
            return ValidationCheck(
                code="loan-fields",
                label="کامل بودن فیلدهای وام",
                scope=analysis.primary_table.sheet,
                status=FAILED,
                detail="هیچ رکورد وامی از جدول استخراج نشد.",
            )
        missing = [
            loan for loan in analysis.loans
            if loan.principal is None or loan.installment is None or loan.installment_count is None
        ]
        bad_dates = [
            loan for loan in analysis.loans
            if loan.last_installment_date is not None
            and not (date(1990, 1, 1) <= loan.last_installment_date <= date(2100, 12, 31))
        ]
        overpaid = [
            loan for loan in analysis.loans
            if loan.paid_installments is not None
            and loan.installment_count is not None
            and loan.paid_installments > loan.installment_count
        ]
        if overpaid:
            return ValidationCheck(
                code="loan-fields",
                label="کامل بودن و سازگاری فیلدهای وام",
                scope=analysis.primary_table.sheet,
                status=FAILED,
                affected=len(overpaid),
                detail="تعداد اقساط پرداخت‌شده از کل اقساط بیشتر است: "
                + "، ".join(loan.institution_raw for loan in overpaid[:5]),
            )
        if missing:
            return ValidationCheck(
                code="loan-fields",
                label="کامل بودن فیلدهای وام",
                scope=analysis.primary_table.sheet,
                status=FAILED,
                affected=len(missing),
                detail="فیلد ضروری ناقص: " + "، ".join(loan.institution_raw for loan in missing[:5]),
            )
        if bad_dates:
            return ValidationCheck(
                code="loan-fields",
                label="بازه معتبر تاریخ اقساط",
                scope=analysis.primary_table.sheet,
                status=WARNING,
                affected=len(bad_dates),
                detail="تاریخ خارج از بازه ۱۹۹۰ تا ۲۱۰۰: "
                + "، ".join(loan.institution_raw for loan in bad_dates[:5]),
            )
        return ValidationCheck(
            code="loan-fields",
            label="کامل بودن فیلدهای وام",
            scope=analysis.primary_table.sheet,
            status=PASSED,
            affected=len(analysis.loans),
            detail=f"همه {len(analysis.loans)} رکورد وام فیلدهای ضروری دارند.",
        )

    @staticmethod
    def _reconciliation_checks(analysis: AnalysisResult) -> ValidationCheck:
        if not analysis.reconciliation:
            return ValidationCheck(
                code="reconciliation",
                label="تطبیق تعهدات",
                scope=Path(analysis.meta.path).name,
                status=SKIPPED,
                detail="داده‌ای برای تطبیق وجود نداشت.",
            )
        review = [
            item for item in analysis.reconciliation
            if item.status.value in {"probable", "outside_loan_table"}
        ]
        matched = [item for item in analysis.reconciliation if item.status.value == "match"]
        if review:
            return ValidationCheck(
                code="reconciliation",
                label="تطبیق تعهدات",
                scope=Path(analysis.meta.path).name,
                status=WARNING,
                affected=len(review),
                detail=f"{len(matched)} تطبیق قطعی، {len(review)} مورد نیازمند بررسی.",
            )
        return ValidationCheck(
            code="reconciliation",
            label="تطبیق تعهدات",
            scope=Path(analysis.meta.path).name,
            status=PASSED,
            affected=len(matched),
            detail=f"همه {len(matched)} تعهد با جدول وام‌ها تطبیق قطعی یافت.",
        )
