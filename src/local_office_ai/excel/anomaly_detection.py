"""Anomaly detection: duplicates, outliers, impossible values, risky structure.

Findings are advisory and always carry severity, location and detail.  Nothing
here changes a value; the operator decides.
"""

from __future__ import annotations

from decimal import Decimal
from statistics import median

from local_office_ai.core.logging_config import get_logger
from local_office_ai.core.models import (
    Anomaly,
    AuditStatus,
    Calculation,
    LoanRecord,
    MatchStatus,
    ObligationRecord,
    ReconciliationItem,
    Severity,
    TableCandidate,
    TableKind,
    ValueKind,
    WorkbookMeta,
)
from local_office_ai.core.settings import Settings

LOGGER = get_logger(__name__)


class AnomalyDetector:
    def __init__(self, settings: Settings) -> None:
        self.settings = settings

    def detect(
        self,
        meta: WorkbookMeta,
        tables: list[TableCandidate],
        primary: TableCandidate | None,
        loans: list[LoanRecord],
        obligations: list[ObligationRecord],
        calculations: list[Calculation],
        audits,
        reconciliation: list[ReconciliationItem],
        normalized_rows: list[dict] | None = None,
    ) -> list[Anomaly]:
        anomalies: list[Anomaly] = []
        anomalies.extend(self._workbook_anomalies(meta))
        anomalies.extend(self._structure_anomalies(tables, primary))
        anomalies.extend(self._loan_anomalies(loans))
        anomalies.extend(self._obligation_anomalies(obligations))
        anomalies.extend(self._calculation_anomalies(calculations))
        anomalies.extend(self._audit_anomalies(audits))
        anomalies.extend(self._reconciliation_anomalies(reconciliation))
        if normalized_rows:
            anomalies.extend(self._duplicate_row_anomalies(normalized_rows))
        order = {Severity.CRITICAL: 0, Severity.WARNING: 1, Severity.INFO: 2}
        anomalies.sort(key=lambda item: order.get(item.severity, 3))
        LOGGER.info("anomalies: %d findings", len(anomalies))
        return anomalies

    @staticmethod
    def _duplicate_row_anomalies(rows: list[dict]) -> list[Anomaly]:
        """Identical data rows in a generic table are worth a look."""
        seen: dict[tuple[str, ...], int] = {}
        anomalies: list[Anomaly] = []
        for row in rows:
            values = tuple(
                str(cell["value"])
                for cell in row["cells"]
                if cell["value"] not in (None, "") and cell.get("role") != "id"
            )
            if len(values) < 2:
                continue
            if values in seen:
                anomalies.append(
                    Anomaly(
                        Severity.WARNING,
                        "duplicate-row",
                        f"ردیف تکراری در برگه «{row['sheet']}»",
                        row["sheet"],
                        f"A{row['row']}",
                        f"ردیف مشابه: A{seen[values]}",
                    )
                )
            else:
                seen[values] = row["row"]
        return anomalies

    # ------------------------------------------------------------------ #
    def _workbook_anomalies(self, meta: WorkbookMeta) -> list[Anomaly]:
        anomalies: list[Anomaly] = []
        for sheet in meta.sheets:
            for address in sheet.error_cells[:50]:
                anomalies.append(
                    Anomaly(
                        Severity.WARNING,
                        "formula-error",
                        f"سلول دارای خطای فرمول در برگه «{sheet.title}»",
                        sheet.title,
                        address,
                        "مقدار خطا در محاسبات استفاده نشد.",
                    )
                )
            if sheet.state != "visible":
                anomalies.append(
                    Anomaly(
                        Severity.WARNING,
                        "hidden-sheet",
                        f"برگه «{sheet.title}» مخفی است و در تحلیل لحاظ شد.",
                        sheet.title,
                    )
                )
            if sheet.hidden_cols:
                anomalies.append(
                    Anomaly(
                        Severity.INFO,
                        "hidden-columns",
                        f"{sheet.hidden_cols} ستون مخفی در برگه «{sheet.title}»",
                        sheet.title,
                        detail="ستون‌های مخفی نادیده گرفته نشدند.",
                    )
                )
            if sheet.protected:
                anomalies.append(
                    Anomaly(
                        Severity.INFO,
                        "protected-sheet",
                        f"برگه «{sheet.title}» محافظت‌شده است.",
                        sheet.title,
                        detail="فقط خواندن انجام شد؛ فایل تغییر نکرد.",
                    )
                )
        if meta.has_macros:
            anomalies.append(
                Anomaly(Severity.WARNING, "macros", "فایل دارای ماکرو است؛ ماکرو اجرا یا منتقل نشد.")
            )
        if meta.external_links:
            anomalies.append(
                Anomaly(
                    Severity.WARNING,
                    "external-links",
                    f"{meta.external_links} پیوند خارجی در فایل وجود دارد.",
                    detail="مقادیر ذخیره‌شده خوانده شدند، نه مقادیر پیوند.",
                )
            )
        return anomalies

    def _structure_anomalies(
        self, tables: list[TableCandidate], primary: TableCandidate | None
    ) -> list[Anomaly]:
        anomalies: list[Anomaly] = []
        if primary is None:
            anomalies.append(
                Anomaly(
                    Severity.CRITICAL,
                    "no-financial-table",
                    "هیچ جدول مالی قابل اتکایی در فایل شناسایی نشد.",
                    detail="ساختار فایل گزارش شد اما محاسبه مالی انجام نشد.",
                )
            )
            return anomalies
        if primary.confidence < 0.6:
            anomalies.append(
                Anomaly(
                    Severity.WARNING,
                    "low-structure-confidence",
                    f"اطمینان تشخیص ساختار جدول اصلی پایین است ({primary.confidence:.2f}).",
                    primary.sheet,
                    primary.address,
                    "ستون‌ها باید توسط کاربر تأیید شوند.",
                )
            )
        if primary.kind is TableKind.GENERIC:
            anomalies.append(
                Anomaly(
                    Severity.INFO,
                    "generic-table",
                    "جدول اصلی معنای وام/تعهد ندارد؛ تحلیل عمومی ستون‌ها انجام شد.",
                    primary.sheet,
                    primary.address,
                )
            )
        for table in tables:
            if table.kind is TableKind.TIME_SERIES:
                continue
            for column in table.columns:
                if column.data_type is ValueKind.EMPTY:
                    anomalies.append(
                        Anomaly(
                            Severity.INFO,
                            "empty-column",
                            f"ستون «{column.header or column.letter}» داده‌ای ندارد.",
                            table.sheet,
                            column.header_address,
                        )
                    )
                if column.data_type is ValueKind.ERROR:
                    anomalies.append(
                        Anomaly(
                            Severity.WARNING,
                            "error-column",
                            f"ستون «{column.header or column.letter}» بیشتر مقدار خطا دارد.",
                            table.sheet,
                            column.header_address,
                        )
                    )
                if column.numeric_ratio >= 0.5 and column.empty_ratio > 0.4:
                    anomalies.append(
                        Anomaly(
                            Severity.INFO,
                            "sparse-column",
                            f"{column.empty_ratio:.0%} از ستون «{column.header or column.letter}» خالی است.",
                            table.sheet,
                            column.header_address,
                        )
                    )
        return anomalies

    def _loan_anomalies(self, loans: list[LoanRecord]) -> list[Anomaly]:
        anomalies: list[Anomaly] = []
        seen: dict[tuple, LoanRecord] = {}
        for loan in loans:
            location = loan.provenance.address
            sheet = loan.provenance.sheet
            if loan.principal is None:
                anomalies.append(
                    Anomaly(Severity.WARNING, "missing-principal", "مبلغ وام خوانده نشد.", sheet, location)
                )
            if loan.installment is None:
                anomalies.append(
                    Anomaly(Severity.WARNING, "missing-installment", "مبلغ قسط خوانده نشد.", sheet, location)
                )
            if loan.installment_count is None:
                anomalies.append(
                    Anomaly(Severity.WARNING, "missing-count", "تعداد اقساط خوانده نشد.", sheet, location)
                )
            if loan.principal is not None and loan.principal < 0:
                anomalies.append(
                    Anomaly(Severity.CRITICAL, "negative-amount", "مبلغ وام منفی است.", sheet, location)
                )
            if loan.installment is not None and loan.installment < 0:
                anomalies.append(
                    Anomaly(Severity.CRITICAL, "negative-amount", "مبلغ قسط منفی است.", sheet, location)
                )
            if loan.installment_count is not None and loan.installment_count <= 0:
                anomalies.append(
                    Anomaly(Severity.CRITICAL, "invalid-count", "تعداد اقساط صفر یا منفی است.", sheet, location)
                )
            if loan.installment_count is not None and loan.installment_count != loan.installment_count.to_integral_value():
                anomalies.append(
                    Anomaly(Severity.WARNING, "fractional-count", "تعداد اقساط اعشاری است.", sheet, location)
                )
            key = (loan.institution_key, loan.principal, loan.installment, loan.installment_count)
            if key in seen and loan.principal is not None:
                anomalies.append(
                    Anomaly(
                        Severity.WARNING,
                        "duplicate-loan",
                        f"ردیف وام تکراری برای «{loan.institution_raw}»",
                        sheet,
                        location,
                        f"ردیف مشابه: {seen[key].provenance.address}",
                    )
                )
            else:
                seen[key] = loan
            for issue in loan.issues:
                anomalies.append(
                    Anomaly(Severity.INFO, "row-issue", f"{loan.institution_raw}: {issue}", sheet, location)
                )

        for field_name, label in (("principal", "اصل وام"), ("installment", "قسط ماهانه")):
            values = [
                getattr(loan, field_name) for loan in loans if getattr(loan, field_name) is not None
            ]
            anomalies.extend(self._outliers(values, loans, field_name, label))
        return anomalies

    def _outliers(
        self,
        values: list[Decimal],
        loans: list[LoanRecord],
        field_name: str,
        label: str,
    ) -> list[Anomaly]:
        if len(values) < 4:
            return []
        center = median(values)
        deviations = [abs(value - center) for value in values]
        mad = median(deviations)
        if mad == 0:
            return []
        anomalies: list[Anomaly] = []
        for loan in loans:
            value = getattr(loan, field_name)
            if value is None:
                continue
            score = Decimal("0.6745") * abs(value - center) / mad
            if score > Decimal(str(self.settings.outlier_mad_threshold)):
                anomalies.append(
                    Anomaly(
                        Severity.INFO,
                        "outlier",
                        f"{label} «{loan.institution_raw}» نسبت به سایر ردیف‌ها غیرعادی است.",
                        loan.provenance.sheet,
                        loan.provenance.address,
                        f"مقدار {value} در برابر میانه {center} (امتیاز MAD {score:.1f})",
                    )
                )
        return anomalies

    def _obligation_anomalies(self, obligations: list[ObligationRecord]) -> list[Anomaly]:
        anomalies: list[Anomaly] = []
        seen: dict[tuple, ObligationRecord] = {}
        for item in obligations:
            if item.amount <= 0:
                anomalies.append(
                    Anomaly(
                        Severity.WARNING,
                        "non-positive-obligation",
                        f"تعهد ماهانه «{item.institution_raw}» صفر یا منفی است.",
                        item.provenance.sheet,
                        item.provenance.address,
                    )
                )
            key = (item.institution_key, item.amount)
            if key in seen:
                anomalies.append(
                    Anomaly(
                        Severity.INFO,
                        "duplicate-obligation",
                        f"تعهد ماهانه تکراری برای «{item.institution_raw}» با مبلغ یکسان.",
                        item.provenance.sheet,
                        item.provenance.address,
                        f"ردیف مشابه: {seen[key].provenance.address}",
                    )
                )
            else:
                seen[key] = item
        return anomalies

    def _calculation_anomalies(self, calculations: list[Calculation]) -> list[Anomaly]:
        anomalies: list[Anomaly] = []
        for calculation in calculations:
            if calculation.verified is False:
                anomalies.append(
                    Anomaly(
                        Severity.CRITICAL,
                        "verification-failed",
                        f"کنترل دوباره محاسبه «{calculation.label}» برابر نشد.",
                        detail=(
                            f"مقدار اصلی {calculation.value} در برابر "
                            f"{calculation.verification_value} ({calculation.verification_method})"
                        ),
                    )
                )
        return anomalies

    def _audit_anomalies(self, audits) -> list[Anomaly]:
        anomalies: list[Anomaly] = []
        for audit in audits:
            if audit.status is AuditStatus.MISMATCH:
                anomalies.append(
                    Anomaly(
                        Severity.CRITICAL,
                        "total-mismatch",
                        f"مغایرت بین مقدار اعلامی و محاسبه‌شده برای «{audit.label}»",
                        audit.declared_source.sheet if audit.declared_source else "",
                        audit.declared_source.address if audit.declared_source else "",
                        f"اختلاف: {audit.difference}",
                    )
                )
            elif audit.status is AuditStatus.NOT_VERIFIABLE:
                anomalies.append(
                    Anomaly(
                        Severity.WARNING,
                        "total-not-verifiable",
                        f"مقدار اعلامی «{audit.label}» قابل بررسی نبود.",
                    )
                )
        return anomalies

    def _reconciliation_anomalies(self, reconciliation: list[ReconciliationItem]) -> list[Anomaly]:
        anomalies: list[Anomaly] = []
        for item in reconciliation:
            if item.status is MatchStatus.PROBABLE:
                anomalies.append(
                    Anomaly(
                        Severity.WARNING,
                        "probable-match",
                        f"تطبیق احتمالی «{item.institution}» نیازمند تأیید است.",
                        detail=f"شباهت {item.similarity:.2f}؛ اختلاف {item.difference}",
                    )
                )
            elif item.status is MatchStatus.OUTSIDE_LOAN_TABLE:
                anomalies.append(
                    Anomaly(
                        Severity.WARNING,
                        "obligation-outside-table",
                        f"تعهد ماهانه «{item.obligation_name}» خارج از جدول وام‌ها ثبت شده است.",
                        item.obligation_sources[0].sheet if item.obligation_sources else "",
                        item.obligation_sources[0].address if item.obligation_sources else "",
                        f"مبلغ {item.obligation_total}",
                    )
                )
            elif item.status is MatchStatus.NO_OBLIGATION:
                anomalies.append(
                    Anomaly(
                        Severity.INFO,
                        "loan-without-obligation",
                        f"برای وام «{item.loan_name}» ردیف تعهد ماهانه پیدا نشد.",
                    )
                )
        return anomalies
