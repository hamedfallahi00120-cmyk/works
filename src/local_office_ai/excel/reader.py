"""Safe workbook loading.

Guarantees
----------
* The input file is opened read-only and its SHA-256 is recorded before any
  processing; :meth:`WorkbookBundle.verify_unchanged` proves it was not modified.
* Macros are never executed and never copied: a macro-enabled file is read as
  data and the output is a fresh ``.xlsx``.
* Locked, corrupt, legacy or oversized files fail with a precise exception.
"""

from __future__ import annotations

import hashlib
import time
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook

from local_office_ai.core.exceptions import (
    CorruptWorkbookError,
    UnsupportedWorkbookError,
    WorkbookAccessError,
)
from local_office_ai.core.logging_config import get_logger
from local_office_ai.core.models import WorkbookMeta
from local_office_ai.core.settings import Settings

LOGGER = get_logger(__name__)

SUPPORTED_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm"}
MACRO_SUFFIXES = {".xlsm", ".xltm"}
LEGACY_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def file_sha256(path: Path, chunk_size: int = 1 << 20) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(chunk_size), b""):
            digest.update(chunk)
    return digest.hexdigest()


@dataclass
class WorkbookBundle:
    """Both views of a workbook plus integrity metadata."""

    path: Path
    values: Workbook
    formulas: Workbook
    sha256: str
    size_bytes: int
    has_macros: bool
    load_seconds: float
    warnings: list[str] = field(default_factory=list)
    meta: WorkbookMeta | None = None
    _mtime_ns: int = 0

    def verify_unchanged(self) -> bool:
        """Return True when the source file is byte-identical to load time."""
        return (
            self.path.is_file()
            and file_sha256(self.path) == self.sha256
            and self.path.stat().st_size == self.size_bytes
        )

    def close(self) -> None:
        for workbook in (self.values, self.formulas):
            try:
                workbook.close()
            except Exception:  # noqa: BLE001 - closing must never mask a result
                pass


class WorkbookReader:
    def __init__(self, settings: Settings) -> None:
        self.settings = settings

    def load(self, path: Path | str) -> WorkbookBundle:
        source = Path(path).expanduser()
        warnings: list[str] = []

        if not source.exists():
            raise WorkbookAccessError(f"فایل پیدا نشد: {source}")
        if not source.is_file():
            raise WorkbookAccessError(f"مسیر یک فایل نیست: {source}")

        suffix = source.suffix.lower()
        if suffix not in SUPPORTED_SUFFIXES:
            if suffix == ".xls":
                raise UnsupportedWorkbookError(
                    "فایل‌های قدیمی .xls پشتیبانی نمی‌شوند. فایل را در Excel با فرمت .xlsx ذخیره کنید."
                )
            if suffix in {".csv", ".txt"}:
                raise UnsupportedWorkbookError(
                    "فایل متنی/CSV پشتیبانی نمی‌شود. داده را در یک فایل .xlsx وارد کنید."
                )
            raise UnsupportedWorkbookError(f"پسوند پشتیبانی‌نشده: {suffix or '(بدون پسوند)'}")

        size_bytes = source.stat().st_size
        if size_bytes == 0:
            raise CorruptWorkbookError("فایل خالی است (۰ بایت).")
        max_bytes = self.settings.max_file_mb * 1024 * 1024
        if size_bytes > max_bytes:
            raise UnsupportedWorkbookError(
                f"حجم فایل {size_bytes / (1024 ** 2):.1f} مگابایت است؛ سقف مجاز "
                f"{self.settings.max_file_mb} مگابایت."
            )

        with source.open("rb") as handle:
            head = handle.read(8)
        if head.startswith(LEGACY_MAGIC):
            raise UnsupportedWorkbookError(
                "فایل در قالب قدیمی OLE (xls) است. آن را به .xlsx تبدیل کنید."
            )
        if not zipfile.is_zipfile(source):
            raise CorruptWorkbookError("فایل یک بسته Excel معتبر نیست (zip سالم نیست).")

        lock_file = source.with_name(f"~${source.name}")
        if lock_file.exists():
            warnings.append(
                "فایل قفل موقت Excel (~$...) دیده شد؛ ممکن است فایل باز باشد. "
                "خواندن انجام شد اما برای اطمینان فایل را ببندید."
            )
        if not _is_readable(source):
            raise WorkbookAccessError(f"اجازه خواندن فایل وجود ندارد: {source}")

        has_macros = suffix in MACRO_SUFFIXES
        if has_macros:
            warnings.append("فایل ماکرو دارد؛ ماکروها اجرا و به خروجی منتقل نمی‌شوند.")

        digest = file_sha256(source)
        started = time.perf_counter()
        try:
            values = load_workbook(source, data_only=True, keep_links=False)
            formulas = load_workbook(source, data_only=False, keep_links=False)
        except PermissionError as exc:
            raise WorkbookAccessError(f"دسترسی به فایل ممکن نیست: {exc}") from exc
        except (OSError, zipfile.BadZipFile, KeyError, ValueError) as exc:
            raise CorruptWorkbookError(f"ساختار فایل Excel قابل خواندن نیست: {exc}") from exc
        load_seconds = time.perf_counter() - started

        if not values.sheetnames:
            raise CorruptWorkbookError("فایل هیچ Sheet قابل خواندنی ندارد.")
        if len(values.sheetnames) > self.settings.max_sheets:
            raise UnsupportedWorkbookError(
                f"تعداد Sheetها ({len(values.sheetnames)}) از سقف {self.settings.max_sheets} بیشتر است."
            )

        external_links = 0
        try:
            external_links = len(getattr(values, "_external_links", []) or [])
        except Exception:  # noqa: BLE001 - private attribute, best effort only
            external_links = 0
        if external_links:
            warnings.append(f"{external_links} پیوند خارجی شناسایی شد؛ مقادیر ذخیره‌شده خوانده می‌شوند.")

        bundle = WorkbookBundle(
            path=source,
            values=values,
            formulas=formulas,
            sha256=digest,
            size_bytes=size_bytes,
            has_macros=has_macros,
            load_seconds=load_seconds,
            warnings=warnings,
            _mtime_ns=source.stat().st_mtime_ns,
        )
        LOGGER.info("workbook loaded: %s (%.2fs, %d bytes)", source.name, load_seconds, size_bytes)
        return bundle


def _is_readable(path: Path) -> bool:
    try:
        with path.open("rb") as handle:
            handle.read(1)
        return True
    except OSError:
        return False
