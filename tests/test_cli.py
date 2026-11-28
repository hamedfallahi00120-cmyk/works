from __future__ import annotations

import json
import zipfile

from local_office_ai.ui.cli import main


def test_missing_file_returns_usage_error(project_root, tmp_path):
    code = main(["--root", str(project_root), "--file", str(tmp_path / "missing.xlsx")])
    assert code == 2


def test_missing_argument_returns_usage_error(project_root):
    assert main(["--root", str(project_root)]) == 2


def test_list_models_uses_config(project_root, fake_ollama, capsys):
    base_url = fake_ollama.start()
    config = project_root / "config" / "settings.json"
    config.write_text(json.dumps({"ollama_base_url": base_url}), encoding="utf-8")
    code = main(["--root", str(project_root), "--list-models"])
    captured = capsys.readouterr()
    assert code == 0
    assert "qwen3.5:9b" in captured.out
    assert "*" in captured.out


def test_list_models_reports_connection_failure(project_root, capsys):
    config = project_root / "config" / "settings.json"
    config.write_text(json.dumps({"ollama_base_url": "http://127.0.0.1:9", "ollama_retries": 0}), encoding="utf-8")
    code = main(["--root", str(project_root), "--list-models"])
    captured = capsys.readouterr()
    assert code == 1
    assert "Ollama" in captured.err


def test_check_structure(project_root, sample_path, capsys):
    code = main(["--root", str(project_root), "--file", str(sample_path), "--check-structure"])
    captured = capsys.readouterr()
    assert code == 0
    assert "Sheet1" in captured.out
    assert "جدول وام" in captured.out
    assert "جدول اصلی" in captured.out


def test_full_headless_run(project_root, sample_path, tmp_path, capsys):
    output = tmp_path / "cli_out.xlsx"
    json_path = tmp_path / "cli_summary.json"
    code = main(
        [
            "--root",
            str(project_root),
            "--file",
            str(sample_path),
            "--no-ai",
            "--yes",
            "--quiet",
            "--output",
            str(output),
            "--json",
            str(json_path),
        ]
    )
    assert code == 0
    assert output.exists()
    assert json_path.exists()
    payload = json.loads(json_path.read_text(encoding="utf-8"))
    assert payload["status"] == "completed"
    values = {item["name"]: item["value"] for item in payload["analysis"]["calculations"]}
    assert values["principal_sum"] == "325000000"


def test_corrupt_file_returns_error(project_root, tmp_path, capsys):
    bad = tmp_path / "bad.xlsx"
    bad.write_bytes(b"not an excel file")
    code = main(["--root", str(project_root), "--file", str(bad), "--no-ai", "--yes"])
    captured = capsys.readouterr()
    assert code == 1
    assert "خطا" in captured.err


def test_cli_image_generation_uses_both_services(project_root, sample_path, fake_ollama, fake_comfy, tmp_path, capsys):
    config = project_root / "config" / "settings.json"
    config.write_text(
        json.dumps(
            {
                "ollama_base_url": fake_ollama.start(),
                "comfyui_base_url": fake_comfy.start(),
                "comfyui_generation_timeout_seconds": 10,
                "ollama_retries": 1,
            }
        ),
        encoding="utf-8",
    )
    output = tmp_path / "with_image.xlsx"
    code = main(
        [
            "--root",
            str(project_root),
            "--file",
            str(sample_path),
            "--no-ai",
            "--yes",
            "--quiet",
            "--output",
            str(output),
            "--image",
            "a tidy office desk",
            "--no-image-ai",
        ]
    )
    captured = capsys.readouterr()
    assert code == 0, captured.err
    assert output.exists()
    assert "تصویر:" in captured.out
    images = list((project_root / "data" / "output" / "images").glob("*.png"))
    assert images and images[0].read_bytes().startswith(b"\x89PNG")


def test_output_suffix_is_validated(make_workbook, project_root):
    """G11 / R8: a ``.docx`` output must become a real Word document.

    Prompt 04B lets a Word request build a complete report without a template,
    so the suffix now yields a genuine DOCX package (never XLSX content under a
    Word name).  An explicit Excel format with a ``.docx`` name is still
    rejected.
    """
    source = make_workbook(
        "costs.xlsx",
        {
            "هزینه‌ها": [
                ["جدول هزینه‌های جاری"],
                [],
                ["ردیف", "شرح", "مبلغ", "دسته"],
                [1, "اجاره", 1250000, "اجاره"],
                [2, "برق", 1250000, "انرژی"],
                [3, "اینترنت", 850000, "ارتباطات"],
            ]
        },
    )
    output = project_root / "report.docx"
    code = main(
        [
            "--root",
            str(project_root),
            "--file",
            str(source),
            "--no-ai",
            "--yes",
            "--quiet",
            "--output",
            str(output),
        ]
    )
    assert code == 0
    assert output.exists()
    assert zipfile.is_zipfile(output)
    with zipfile.ZipFile(output) as archive:
        names = archive.namelist()
    assert "word/document.xml" in names
    assert not any(name.startswith("xl/") for name in names)

    mismatched = project_root / "wrong.docx"
    assert (
        main(
            [
                "--root",
                str(project_root),
                "--file",
                str(source),
                "--no-ai",
                "--yes",
                "--quiet",
                "--format",
                "xlsx",
                "--output",
                str(mismatched),
            ]
        )
        == 2
    )
    assert not mismatched.exists()
