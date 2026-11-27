"""Reconciliation between loan installments and separately recorded obligations.

Matching is never binary.  A pair is reported as a definite match, a probable
match that needs confirmation, or no match at all; obligations that exist only
outside the loan table are reported as such.  Names are compared on a
normalized key plus a similarity score, and the original spellings are always
shown so the operator can judge.
"""

from __future__ import annotations

from decimal import Decimal
from typing import Any

from local_office_ai.core.logging_config import get_logger
from local_office_ai.core.models import (
    LoanRecord,
    MatchStatus,
    ObligationRecord,
    ReconciliationItem,
)
from local_office_ai.core.settings import Settings
from local_office_ai.excel.normalization import institution_similarity

LOGGER = get_logger(__name__)


class Reconciler:
    def __init__(self, settings: Settings) -> None:
        self.settings = settings
        self.synonyms = dict(settings.synonyms)

    def reconcile(
        self,
        loans: list[LoanRecord],
        obligations: list[ObligationRecord],
    ) -> list[ReconciliationItem]:
        loan_groups = self._group_loans(loans)
        obligation_groups = self._group_obligations(obligations)

        unmatched_loans = dict(loan_groups)
        unmatched_obligations = dict(obligation_groups)
        items: list[ReconciliationItem] = []

        # 1) Exact key matches.
        for key in list(unmatched_loans):
            if key in unmatched_obligations:
                items.append(self._build_item(key, unmatched_loans.pop(key), unmatched_obligations.pop(key), 1.0, "کلید یکسان"))
        # 2) Probable matches by similarity, strongest first.
        candidates: list[tuple[float, str, str]] = []
        for loan_key in unmatched_loans:
            for obligation_key in unmatched_obligations:
                score = institution_similarity(loan_key, obligation_key, self.synonyms)
                if score >= self.settings.reconciliation_probable_threshold:
                    candidates.append((score, loan_key, obligation_key))
        candidates.sort(reverse=True)
        used_loans: set[str] = set()
        used_obligations: set[str] = set()
        for score, loan_key, obligation_key in candidates:
            if loan_key in used_loans or obligation_key in used_obligations:
                continue
            used_loans.add(loan_key)
            used_obligations.add(obligation_key)
            items.append(
                self._build_item(
                    loan_key,
                    unmatched_loans[loan_key],
                    unmatched_obligations[obligation_key],
                    score,
                    "تطبیق احتمالی بر اساس شباهت نام",
                )
            )
            unmatched_loans.pop(loan_key, None)
            unmatched_obligations.pop(obligation_key, None)
        # 3) Leftovers.
        for loan_key, group in unmatched_loans.items():
            items.append(
                ReconciliationItem(
                    institution=group["name"],
                    loan_total=group["total"],
                    obligation_total=Decimal("0"),
                    difference=None,
                    status=MatchStatus.NO_OBLIGATION,
                    loan_name=group["name"],
                    obligation_name="",
                    similarity=0.0,
                    note="برای این وام ردیف تعهد ماهانه‌ای پیدا نشد.",
                    loan_sources=group["sources"],
                )
            )
        for obligation_key, group in unmatched_obligations.items():
            items.append(
                ReconciliationItem(
                    institution=group["name"],
                    loan_total=Decimal("0"),
                    obligation_total=group["total"],
                    difference=None,
                    status=MatchStatus.OUTSIDE_LOAN_TABLE,
                    loan_name="",
                    obligation_name=group["name"],
                    similarity=0.0,
                    note="تعهد ماهانه‌ای ثبت شده که ردیف وام متناظر در جدول وام‌ها ندارد.",
                    obligation_sources=group["sources"],
                )
            )

        # 4) Same entity written differently is worth reporting even when matched.
        variants = self._spelling_variants(loan_groups, obligation_groups)
        if variants:
            for item in items:
                for key, note in variants.items():
                    if item.institution == loan_groups.get(key, {}).get("name") or item.institution == obligation_groups.get(key, {}).get("name"):
                        item.note = (item.note + " " + note).strip()

        items.sort(key=lambda item: (item.status is MatchStatus.OUTSIDE_LOAN_TABLE, item.institution))
        LOGGER.info(
            "reconciliation: %d items (%d matched, %d probable, %d outside)",
            len(items),
            sum(1 for item in items if item.status is MatchStatus.MATCH),
            sum(1 for item in items if item.status is MatchStatus.PROBABLE),
            sum(1 for item in items if item.status is MatchStatus.OUTSIDE_LOAN_TABLE),
        )
        return items

    # ------------------------------------------------------------------ #
    def _build_item(
        self,
        key: str,
        loan_group: dict[str, Any],
        obligation_group: dict[str, Any],
        similarity: float,
        reason: str,
    ) -> ReconciliationItem:
        difference = obligation_group["total"] - loan_group["total"]
        amount_match = abs(difference) <= self.settings.money_tolerance
        name_exact = similarity >= self.settings.reconciliation_exact_threshold

        if name_exact and amount_match:
            status = MatchStatus.MATCH
            note = "نام و مبلغ هر دو مطابقت دارند."
        elif name_exact and not amount_match:
            status = MatchStatus.PROBABLE
            note = "نام یکسان است اما مبلغ مغایرت دارد؛ نیازمند تأیید."
        elif amount_match:
            status = MatchStatus.PROBABLE
            note = "مبلغ برابر است اما نام فقط به‌صورت احتمالی منطبق شده؛ نیازمند تأیید."
        else:
            status = MatchStatus.PROBABLE
            note = "نام و مبلغ با اطمینان کامل منطبق نشده‌اند؛ نیازمند تأیید."
        if loan_group["name"] != obligation_group["name"]:
            note += (
                f" شکل‌های نوشتاری «{loan_group['name']}» و «{obligation_group['name']}» "
                "برای تطبیق یکسان‌سازی شدند؛ متن منبع تغییر نکرد."
            )
        return ReconciliationItem(
            institution=loan_group["name"] if name_exact else f"{loan_group['name']} / {obligation_group['name']}",
            loan_total=loan_group["total"],
            obligation_total=obligation_group["total"],
            difference=difference,
            status=status,
            loan_name=loan_group["name"],
            obligation_name=obligation_group["name"],
            similarity=similarity,
            note=f"{note} ({reason})",
            loan_sources=loan_group["sources"],
            obligation_sources=obligation_group["sources"],
        )

    @staticmethod
    def _group_loans(loans: list[LoanRecord]) -> dict[str, dict[str, Any]]:
        groups: dict[str, dict[str, Any]] = {}
        for loan in loans:
            group = groups.setdefault(
                loan.institution_key,
                {"name": loan.institution_raw, "total": Decimal("0"), "sources": [], "names": set()},
            )
            group["total"] += loan.installment or Decimal("0")
            group["sources"].append(loan.provenance)
            group["names"].add(loan.institution_raw)
        return groups

    @staticmethod
    def _group_obligations(obligations: list[ObligationRecord]) -> dict[str, dict[str, Any]]:
        groups: dict[str, dict[str, Any]] = {}
        for item in obligations:
            group = groups.setdefault(
                item.institution_key,
                {"name": item.institution_raw, "total": Decimal("0"), "sources": [], "names": set()},
            )
            group["total"] += item.amount
            group["sources"].append(item.provenance)
            group["names"].add(item.institution_raw)
        return groups

    @staticmethod
    def _spelling_variants(
        loan_groups: dict[str, dict[str, Any]],
        obligation_groups: dict[str, dict[str, Any]],
    ) -> dict[str, str]:
        variants: dict[str, str] = {}
        for groups in (loan_groups, obligation_groups):
            for key, group in groups.items():
                names = {name for name in group["names"] if name}
                if len(names) > 1:
                    variants[key] = (
                        "برای تطبیق، شکل‌های نوشتاری "
                        + " / ".join(sorted(names))
                        + " یکسان‌سازی شدند؛ متن منبع تغییر نکرد."
                    )
        return variants
