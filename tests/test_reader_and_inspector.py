from __future__ import annotations

import shutil

import pytest

from local_office_ai.core.exceptions import (
    CorruptWorkbookError,
    UnsupportedWorkbookError,
    WorkbookAccessError,
)
from local_office_ai.excel.reader import WorkbookReader, file_sha256
from local_office_ai.excel.workbook_inspector import inspect_workbook, used_bounds


def test_sample_loads_and_hash_is_recorded(settings, sample_path):
    bundle = WorkbookReader(settings).load(sample_path)
    assert bundle.sha256 == file_sha256(sample_path)
    assert bundle.size_bytes == sample_path.stat().st_size
    assert bundle.verify_unchanged()
    assert "داشبورد" in bundle.values.sheetnames
    assert not bundle.has_macros
    bundle.close()


def test_inspector_separates_data_from_formatting(settings, sample_path):
    bundle = WorkbookReader(settings).load(sample_path)
    meta = inspect_workbook(bundle, settings)
    sheet1 = next(sheet for sheet in meta.sheets if sheet.title == "Sheet1")
    # Data really ends at H15 although formatting reaches BA511.
    assert sheet1.used_dimensions == "A1:H15"
    assert sheet1.declared_dimensions == "A1:BA511"
    assert sheet1.max_row == 15
    assert sheet1.max_col == 8
    assert sheet1.formula_count == 8
    assert sheet1.error_cells == []
    assert meta.total_formulas == 8
    assert any("بزرگ‌تر از داده واقعی" in note for note in sheet1.suspicious)
    bundle.close()


def test_inspector_reports_hidden_columns_and_errors(settings, generic_workbook):
    bundle = WorkbookReader(settings).load(generic_workbook)
    meta = inspect_workbook(bundle, settings)
    costs = next(sheet for sheet in meta.sheets if sheet.title == "هزینه‌ها")
    assert costs.hidden_cols == 1
    assert "F4" in costs.error_cells
    assert costs.merged_ranges == 1
    assert any("ستون مخفی" in note for note in costs.suspicious)
    bundle.close()


def test_used_bounds_ignores_empty_styled_cells(settings, make_workbook):
    path = make_workbook("styled.xlsx", {"Sheet1": [[1, 2], [3, 4]]})

    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill

    workbook = load_workbook(path)
    worksheet = workbook["Sheet1"]
    worksheet["Z50"].fill = PatternFill("solid", fgColor="FFFF00")
    workbook.save(path)
    workbook.close()

    from openpyxl import load_workbook as load

    reopened = load(path)
    assert used_bounds(reopened["Sheet1"], 10_000, 1_000) == (2, 2)
    reopened.close()


def test_missing_file(settings, tmp_path):
    with pytest.raises(WorkbookAccessError):
        WorkbookReader(settings).load(tmp_path / "nope.xlsx")


def test_unsupported_suffix(settings, tmp_path):
    path = tmp_path / "data.csv"
    path.write_text("a,b\n1,2\n", encoding="utf-8")
    with pytest.raises(UnsupportedWorkbookError):
        WorkbookReader(settings).load(path)


def test_legacy_xls_rejected_by_magic(settings, tmp_path):
    path = tmp_path / "old.xls"
    path.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 64)
    with pytest.raises(UnsupportedWorkbookError):
        WorkbookReader(settings).load(path)


def test_corrupt_zip_rejected(settings, tmp_path):
    path = tmp_path / "broken.xlsx"
    path.write_bytes(b"this is not a zip file")
    with pytest.raises(CorruptWorkbookError):
        WorkbookReader(settings).load(path)


def test_empty_file_rejected(settings, tmp_path):
    path = tmp_path / "zero.xlsx"
    path.write_bytes(b"")
    with pytest.raises(CorruptWorkbookError):
        WorkbookReader(settings).load(path)


def test_size_limit(settings, sample_path):
    settings.max_file_mb = 0
    with pytest.raises(UnsupportedWorkbookError):
        WorkbookReader(settings).load(sample_path)


def test_input_is_never_modified(settings, sample_path, tmp_path):
    copy = tmp_path / "copy.xlsx"
    shutil.copy2(sample_path, copy)
    before = file_sha256(copy)
    bundle = WorkbookReader(settings).load(copy)
    inspect_workbook(bundle, settings)
    bundle.close()
    assert file_sha256(copy) == before


def test_lock_file_is_reported(settings, make_workbook):
    path = make_workbook("locked.xlsx", {"Sheet1": [["a", "b"], [1, 2]]})
    (path.parent / f"~${path.name}").write_text("lock", encoding="utf-8")
    bundle = WorkbookReader(settings).load(path)
    assert any("قفل" in warning for warning in bundle.warnings)
    bundle.close()
