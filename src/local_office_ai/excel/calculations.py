"""Deterministic financial calculations.

Python is the only authority for arithmetic.  Every figure is computed with
``Decimal``, recomputed once through an independent representation and, where
the workbook declares its own total, audited against that declaration.
"""

from __future__ import annotations

from decimal import Decimal, localcontext
from fractions import Fraction
from statistics import median
from typing import Any, Iterable

from local_office_ai.core.logging_config import get_logger
from local_office_ai.core.models import (
    AuditResult,
    AuditStatus,
    Calculation,
    ColumnRole,
    DataRole,
    DeclaredTotal,
    LoanRecord,
    ObligationRecord,
    Provenance,
    TableCandidate,
    ValueKind,
)
from local_office_ai.core.settings import Settings
from local_office_ai.excel.sheet_analyzer import SheetGrid, is_total_label

LOGGER = get_logger(__name__)

MEASURE_LABELS = {
    "principal": "مجموع اصل وام‌ها",
    "installments": "مجموع اقساط ماهانه وام‌ها",
    "contract_repayment": "مجموع بازپرداخت قراردادی",
    "obligations": "مجموع تعهدات ماهانه",
    "count": "تعداد وام‌ها",
    "amount": "مجموع مبالغ",
}


def _sum(values: Iterable[Decimal]) -> Decimal:
    total = Decimal("0")
    for value in values:
        total += value
    return total


def _fraction_sum(values: Iterable[Decimal]) -> Decimal:
    """Independent re-sum through exact rationals."""
    total = Fraction(0)
    for value in values:
        total += Fraction(value)
    return Decimal(total.numerator) / Decimal(total.denominator)


def _decimal_median(values: list[Decimal]) -> Decimal | None:
    if not values:
        return None
    with localcontext() as context:
        context.prec = 28
        return Decimal(median(values))


class FinancialCalculator:
    def __init__(self, settings: Settings) -> None:
        self.settings = settings

    # ------------------------------------------------------------------ #
    def compute(
        self,
        loans: list[LoanRecord],
        obligations: list[ObligationRecord],
        declared_totals: list[DeclaredTotal],
    ) -> list[Calculation]:
        calculations: list[Calculation] = []
        principal_values = [loan.principal for loan in loans if loan.principal is not None]
        installment_values = [loan.installment for loan in loans if loan.installment is not None]
        repayment_values = [
            loan.contract_repayment for loan in loans if loan.contract_repayment is not None
        ]
        obligation_values = [item.amount for item in obligations]

        principal_sum = _sum(principal_values)
        installments_sum = _sum(installment_values)
        repayment_sum = _sum(repayment_values)
        obligations_sum = _sum(obligation_values)

        calculations.append(
            self._calc(
                "loan_count",
                "تعداد وام‌های ثبت‌شده",
                len(loans),
                "count",
                "شمارش ردیف‌های معتبر جدول وام",
                [loan.provenance for loan in loans],
                verification_value=Decimal(len({loan.provenance.row for loan in loans})),
                verification_method="شمارش ردیف‌های یکتا",
            )
        )
        calculations.append(
            self._calc(
                "principal_sum",
                MEASURE_LABELS["principal"],
                principal_sum,
                "ریال/تومان",
                "جمع ستون اصل وام با Decimal",
                [loan.provenance for loan in loans if loan.principal is not None],
                verification_value=_fraction_sum(principal_values),
                verification_method="جمع دوباره با نمایش کسری دقیق (Fraction)",
            )
        )
        calculations.append(
            self._calc(
                "installments_sum",
                MEASURE_LABELS["installments"],
                installments_sum,
                "ریال/تومان",
                "جمع ستون مبلغ هر قسط با Decimal",
                [loan.provenance for loan in loans if loan.installment is not None],
                verification_value=_fraction_sum(installment_values),
                verification_method="جمع دوباره با نمایش کسری دقیق (Fraction)",
            )
        )
        calculations.append(
            self._calc(
                "obligations_sum",
                MEASURE_LABELS["obligations"],
                obligations_sum,
                "ریال/تومان",
                "جمع ردیف‌های تعهد ماهانه",
                [item.provenance for item in obligations],
                verification_value=_fraction_sum(obligation_values),
                verification_method="جمع دوباره با نمایش کسری دقیق (Fraction)",
            )
        )

        # Contract repayment: installment × count, cross-checked with the sheet column.
        computed_repayment = _sum(
            loan.installment * loan.installment_count
            for loan in loans
            if loan.installment is not None and loan.installment_count is not None
        )
        column_repayment = repayment_sum if repayment_values else None
        verified = None
        if column_repayment is not None:
            verified = abs(column_repayment - computed_repayment) <= self.settings.money_tolerance
        calculations.append(
            self._calc(
                "contract_repayment_sum",
                MEASURE_LABELS["contract_repayment"],
                computed_repayment,
                "ریال/تومان",
                "مبلغ هر قسط × تعداد اقساط برای هر وام، سپس جمع",
                [loan.provenance for loan in loans],
                verification_value=column_repayment,
                verification_method="مقایسه با ستون بازپرداخت قراردادی در فایل (در صورت وجود)",
                verified=verified,
            )
        )

        # Averages, median, extremes.
        if principal_values:
            calculations.append(
                self._calc(
                    "principal_avg",
                    "میانگین اصل وام‌ها",
                    (principal_sum / Decimal(len(principal_values))).quantize(Decimal("0.01")),
                    "ریال/تومان",
                    "جمع اصل ÷ تعداد وام دارای مبلغ",
                    [],
                )
            )
            calculations.append(
                self._calc(
                    "principal_median", "میانه اصل وام‌ها",
                    _decimal_median(principal_values), "ریال/تومان", "میانه آماری", [],
                )
            )
            calculations.append(
                self._calc(
                    "principal_max", "بیشترین اصل وام", max(principal_values), "ریال/تومان", "حداکثر", []
                )
            )
            calculations.append(
                self._calc(
                    "principal_min", "کمترین اصل وام", min(principal_values), "ریال/تومان", "حداقل", []
                )
            )
        if installment_values:
            calculations.append(
                self._calc(
                    "installments_avg",
                    "میانگین اقساط ماهانه",
                    (installments_sum / Decimal(len(installment_values))).quantize(Decimal("0.01")),
                    "ریال/تومان",
                    "جمع اقساط ÷ تعداد وام",
                    [],
                )
            )
            calculations.append(
                self._calc(
                    "installments_max", "بیشترین قسط ماهانه",
                    max(installment_values), "ریال/تومان", "حداکثر", [],
                )
            )
            calculations.append(
                self._calc(
                    "installments_min", "کمترین قسط ماهانه",
                    min(installment_values), "ریال/تومان", "حداقل", [],
                )
            )
        if principal_sum:
            ratio = (installments_sum / principal_sum).quantize(Decimal("0.000001"))
            calculations.append(
                self._calc(
                    "installment_to_principal_ratio",
                    "نسبت مجموع اقساط ماهانه به اصل وام",
                    ratio,
                    "نسبت",
                    "مجموع اقساط ÷ مجموع اصل",
                    [],
                )
            )
        if installments_sum:
            calculations.append(
                self._calc(
                    "obligation_gap",
                    "اختلاف تعهدات ماهانه با اقساط وام‌ها",
                    obligations_sum - installments_sum,
                    "ریال/تومان",
                    "مجموع تعهدات − مجموع اقساط",
                    [item.provenance for item in obligations],
                )
            )
            calculations.append(
                self._calc(
                    "obligation_ratio",
                    "نسبت تعهدات ماهانه به اقساط وام‌ها",
                    (obligations_sum / installments_sum).quantize(Decimal("0.000001")),
                    "نسبت",
                    "مجموع تعهدات ÷ مجموع اقساط",
                    [],
                )
            )

        # Remaining balance: only when the file states how many installments were paid.
        paid_loans = [
            loan for loan in loans
            if loan.paid_installments is not None and loan.installment_count is not None
        ]
        if paid_loans:
            remaining_counts = [loan.remaining_installments for loan in paid_loans]
            remaining_counts = [value for value in remaining_counts if value is not None]
            remaining_total = _sum(remaining_counts)
            calculations.append(
                self._calc(
                    "remaining_installments",
                    "مجموع اقساط باقی‌مانده",
                    remaining_total,
                    "قسط",
                    "برای هر وام: تعداد اقساط − اقساط پرداخت‌شده، سپس جمع",
                    [loan.provenance for loan in paid_loans],
                    verification_value=Decimal(len(remaining_counts)),
                    verification_method="شمارش وام‌های دارای تعداد پرداخت‌شده",
                )
            )
            remaining_amounts = [
                loan.remaining_amount for loan in paid_loans if loan.remaining_amount is not None
            ]
            if remaining_amounts:
                calculations.append(
                    self._calc(
                        "remaining_amount",
                        "مجموع مبلغ اقساط باقی‌مانده",
                        _sum(remaining_amounts),
                        "ریال/تومان",
                        "برای هر وام: اقساط باقی‌مانده × مبلغ هر قسط، سپس جمع",
                        [loan.provenance for loan in paid_loans],
                        verification_value=_fraction_sum(remaining_amounts),
                        verification_method="جمع دوباره با Fraction",
                    )
                )

        # Per-institution subtotals, verified against the row-level totals.
        by_institution: dict[str, dict[str, Any]] = {}
        for loan in loans:
            entry = by_institution.setdefault(
                loan.institution_key, {"name": loan.institution_raw, "count": 0, "principal": Decimal("0"), "installment": Decimal("0")}
            )
            entry["count"] += 1
            entry["principal"] += loan.principal or Decimal("0")
            entry["installment"] += loan.installment or Decimal("0")
        for item in obligations:
            entry = by_institution.setdefault(
                item.institution_key, {"name": item.institution_raw, "count": 0, "principal": Decimal("0"), "installment": Decimal("0")}
            )
            entry["obligation"] = entry.get("obligation", Decimal("0")) + item.amount

        for key, entry in sorted(by_institution.items()):
            entry["key"] = key
        calculations.append(
            Calculation(
                name="by_institution",
                label="تجمیع بر اساس مؤسسه",
                value=Decimal(len(by_institution)),
                unit="مؤسسه",
                method="گروه‌بندی ردیف‌ها با کلید یکسان‌سازی‌شده نام مؤسسه",
                inputs=[f"{entry['name']}: {entry['installment']}" for entry in by_institution.values()],
            )
        )
        subtotal = _sum(entry["principal"] for entry in by_institution.values())
        calculations.append(
            self._calc(
                "principal_sum_check",
                "کنترل جمع اصل با تجمیع مؤسسه‌ای",
                subtotal,
                "ریال/تومان",
                "جمع زیرمجموع‌های مؤسسه‌ای",
                [],
                verification_value=principal_sum,
                verification_method="مقایسه با جمع ردیف‌به‌ردیف",
                verified=abs(subtotal - principal_sum) <= self.settings.money_tolerance,
            )
        )
        LOGGER.info(
            "calculations: loans=%d principal=%s installments=%s obligations=%s repayment=%s",
            len(loans), principal_sum, installments_sum, obligations_sum, computed_repayment,
        )
        return calculations

    @staticmethod
    def _calc(
        name: str,
        label: str,
        value: Decimal | int | None,
        unit: str,
        method: str,
        sources: list[Provenance],
        verification_value: Decimal | None = None,
        verification_method: str = "",
        verified: bool | None = None,
    ) -> Calculation:
        if verified is None and verification_value is not None and isinstance(value, Decimal):
            verified = value == verification_value
        return Calculation(
            name=name,
            label=label,
            value=value,
            unit=unit,
            method=method,
            sources=list(sources),
            verification_value=verification_value,
            verification_method=verification_method,
            verified=verified,
        )

    # ------------------------------------------------------------------ #
    def build_audits(
        self,
        calculations: list[Calculation],
        declared_totals: list[DeclaredTotal],
    ) -> list[AuditResult]:
        computed = {item.name: item.value for item in calculations}
        measure_to_calc = {
            "principal": "principal_sum",
            "installments": "installments_sum",
            "contract_repayment": "contract_repayment_sum",
            "obligations": "obligations_sum",
            "count": "loan_count",
            "amount": "amount_sum",
        }
        audits: list[AuditResult] = []
        seen: set[str] = set()
        for declared in sorted(declared_totals, key=lambda item: -item.confidence):
            measure = declared.measure
            if measure in seen:
                continue
            seen.add(measure)
            calc_name = measure_to_calc.get(measure)
            computed_value = computed.get(calc_name) if calc_name else None
            if computed_value is None:
                audits.append(
                    AuditResult(
                        measure=measure,
                        label=MEASURE_LABELS.get(measure, measure),
                        declared=declared.value,
                        computed=None,
                        difference=None,
                        status=AuditStatus.NOT_VERIFIABLE,
                        declared_source=declared.provenance,
                        method="مقدار محاسبه‌شده متناظر پیدا نشد",
                    )
                )
                continue
            computed_decimal = Decimal(computed_value) if not isinstance(computed_value, Decimal) else computed_value
            difference = computed_decimal - declared.value
            status = (
                AuditStatus.VERIFIED
                if abs(difference) <= self.settings.money_tolerance
                else AuditStatus.MISMATCH
            )
            audits.append(
                AuditResult(
                    measure=measure,
                    label=MEASURE_LABELS.get(measure, measure),
                    declared=declared.value,
                    computed=computed_decimal,
                    difference=difference,
                    status=status,
                    declared_source=declared.provenance,
                    method=f"مقدار اعلامی ({declared.mapping}) در برابر محاسبه مستقل Python",
                )
            )

        for measure, calc_name in measure_to_calc.items():
            if measure in seen or calc_name not in computed:
                continue
            value = computed[calc_name]
            audits.append(
                AuditResult(
                    measure=measure,
                    label=MEASURE_LABELS.get(measure, measure),
                    declared=None,
                    computed=Decimal(value) if not isinstance(value, Decimal) else value,
                    difference=None,
                    status=AuditStatus.NO_DECLARED_VALUE,
                    method="در فایل مقدار اعلامی برای این شاخص پیدا نشد",
                )
            )
        return audits

    # ------------------------------------------------------------------ #
    def summarize_table(
        self,
        grid: SheetGrid,
        table: TableCandidate,
        limit_rows: int = 5000,
    ) -> list[Calculation]:
        """Column statistics for a table without loan semantics."""
        from local_office_ai.excel.normalization import parse_value

        calculations: list[Calculation] = []
        rows = range(table.first_data_row, min(table.last_data_row, table.first_data_row + limit_rows - 1) + 1)
        for column in table.columns:
            if column.data_type is ValueKind.NUMBER:
                if column.statistical_role in {DataRole.IDENTIFIER, DataRole.TEMPORAL_KEY}:
                    continue
                numbers: list[Decimal] = []
                for row in rows:
                    cell = grid.get(row, column.index)
                    if cell is None or cell.is_empty:
                        continue
                    parsed = parse_value(cell.value, kind_hint=ValueKind.NUMBER)
                    if isinstance(parsed.value, Decimal):
                        numbers.append(parsed.value)
                if not numbers:
                    continue
                total = _sum(numbers)
                prefix = f"{column.letter}:{column.header or 'بدون عنوان'}"
                calculations.append(
                    self._calc(
                        f"generic_sum::{prefix}", f"جمع {column.header or column.letter}", total,
                        "", "جمع ستون با Decimal", [],
                        verification_value=_fraction_sum(numbers),
                        verification_method="جمع دوباره با Fraction",
                    )
                )
                calculations.append(
                    self._calc(
                        f"generic_avg::{prefix}", f"میانگین {column.header or column.letter}",
                        (total / Decimal(len(numbers))).quantize(Decimal("0.01")), "", "میانگین", [],
                    )
                )
                calculations.append(
                    self._calc(
                        f"generic_median::{prefix}", f"میانه {column.header or column.letter}",
                        _decimal_median(numbers), "", "میانه", [],
                    )
                )
                calculations.append(
                    self._calc(
                        f"generic_min::{prefix}", f"کمینه {column.header or column.letter}",
                        min(numbers), "", "حداقل", [],
                    )
                )
                calculations.append(
                    self._calc(
                        f"generic_max::{prefix}", f"بیشینه {column.header or column.letter}",
                        max(numbers), "", "حداکثر", [],
                    )
                )
                calculations.append(
                    self._calc(
                        f"generic_count::{prefix}", f"تعداد مقادیر {column.header or column.letter}",
                        len(numbers), "ردیف", "شمارش مقادیر عددی", [],
                    )
                )
            elif column.data_type is ValueKind.DATE:
                dates = []
                for row in rows:
                    cell = grid.get(row, column.index)
                    if cell is None or cell.is_empty:
                        continue
                    parsed = parse_value(cell.value, kind_hint=ValueKind.DATE)
                    if parsed.kind is ValueKind.DATE:
                        dates.append(parsed.value)
                if dates:
                    calculations.append(
                        self._calc(
                            f"generic_date_range::{column.letter}",
                            f"بازه تاریخ {column.header or column.letter}",
                            f"{min(dates)} تا {max(dates)}",
                            "",
                            "کمینه و بیشینه تاریخ‌های معتبر (میلادی ISO)",
                            [],
                        )
                    )
        calculations.extend(self._group_sums(grid, table, rows))
        return calculations

    def _group_sums(
        self,
        grid: SheetGrid,
        table: TableCandidate,
        rows: range,
        max_groups: int = 12,
    ) -> list[Calculation]:
        """Group the table by its most descriptive text column and sum the measure.

        The grouping column is chosen from the data, never from a fixed letter: a
        text column with 2..12 distinct values that is not an identifier.
        """
        from local_office_ai.excel.normalization import parse_value

        measure = None
        for role in (ColumnRole.AMOUNT, ColumnRole.INSTALLMENT, ColumnRole.PRINCIPAL):
            measure = table.column_by_role(role)
            if measure is not None:
                break
        if measure is None:
            measure = next(
                (column for column in table.columns if column.data_type is ValueKind.NUMBER),
                None,
            )
        if measure is None:
            return []

        best: tuple[int, Any, dict[str, Decimal]] | None = None
        for column in table.columns:
            if column.index == measure.index or column.data_type is ValueKind.NUMBER:
                continue
            if column.role is ColumnRole.ID:
                continue
            groups: dict[str, Decimal] = {}
            for row in rows:
                key_cell = grid.get(row, column.index)
                value_cell = grid.get(row, measure.index)
                if key_cell is None or key_cell.is_empty or value_cell is None or value_cell.is_empty:
                    continue
                key = str(key_cell.value).strip()
                if not key or is_total_label(key):
                    continue
                parsed = parse_value(value_cell.value, kind_hint=ValueKind.NUMBER)
                if not isinstance(parsed.value, Decimal):
                    continue
                groups[key] = groups.get(key, Decimal("0")) + parsed.value
            if 2 <= len(groups) <= max_groups and (best is None or len(groups) < best[0]):
                best = (len(groups), column, groups)
        if best is None:
            return []

        _, group_column, groups = best
        total = _sum(groups.values())
        prefix = f"{group_column.letter}:{group_column.header or 'بدون عنوان'}"
        results: list[Calculation] = [
            self._calc(
                f"group_count::{prefix}",
                f"تعداد گروه‌های {group_column.header or group_column.letter}",
                len(groups),
                "گروه",
                f"گروه‌بندی ردیف‌ها بر اساس ستون «{group_column.header or group_column.letter}»",
                [],
                verification_value=Decimal(len(set(groups))),
                verification_method="شمارش کلیدهای یکتا",
            )
        ]
        for name, amount in sorted(groups.items(), key=lambda item: item[1], reverse=True):
            share = (
                (amount / total * Decimal("100")).quantize(Decimal("0.01")) if total else None
            )
            results.append(
                self._calc(
                    f"group_sum::{prefix}::{name}",
                    f"جمع {measure.header or measure.letter} برای «{name}»",
                    amount,
                    "",
                    f"فیلتر ردیف‌های «{group_column.header}» = «{name}» و جمع ستون "
                    f"«{measure.header or measure.letter}»",
                    [],
                )
            )
            if share is not None:
                results.append(
                    self._calc(
                        f"group_share::{prefix}::{name}",
                        f"سهم «{name}» از جمع {measure.header or measure.letter}",
                        share,
                        "%",
                        "مبلغ گروه ÷ جمع کل × ۱۰۰",
                        [],
                    )
                )
        return results
