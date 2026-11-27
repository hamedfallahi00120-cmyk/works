"""Turn raw monitor samples and stage timings into a readable performance report."""

from __future__ import annotations

import os
import platform
import shutil
import subprocess
import sys
from dataclasses import dataclass, field
from datetime import datetime
from statistics import median
from typing import Any

from local_office_ai.core.settings import Settings
from local_office_ai.monitoring.resource_monitor import IS_WINDOWS, ResourceMonitor, psutil


def _numeric(values: list[Any]) -> list[float]:
    return [float(value) for value in values if isinstance(value, (int, float)) and not isinstance(value, bool)]


def stats(values: list[Any]) -> dict[str, float | int | None]:
    numbers = _numeric(values)
    if not numbers:
        return {"count": 0, "avg": None, "max": None, "min": None, "p95": None, "median": None}
    ordered = sorted(numbers)
    index = min(len(ordered) - 1, max(0, int(round(0.95 * (len(ordered) - 1)))))
    return {
        "count": len(numbers),
        "avg": sum(numbers) / len(numbers),
        "max": max(numbers),
        "min": min(numbers),
        "p95": ordered[index],
        "median": median(numbers),
    }


def _cpu_model() -> str:
    if IS_WINDOWS:
        try:
            completed = subprocess.run(
                ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command",
                 "Get-CimInstance Win32_Processor | Select-Object -ExpandProperty Name"],
                capture_output=True, text=True, timeout=8,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            if completed.stdout.strip():
                return completed.stdout.strip().splitlines()[0].strip()
        except (OSError, subprocess.SubprocessError):
            pass
    try:
        with open("/proc/cpuinfo", "r", encoding="utf-8", errors="ignore") as handle:
            for line in handle:
                if line.lower().startswith("model name"):
                    return line.split(":", 1)[1].strip()
    except OSError:
        pass
    return platform.processor() or "نامشخص"


def _gpu_names() -> list[str]:
    names: list[str] = []
    if shutil.which("nvidia-smi"):
        try:
            completed = subprocess.run(
                ["nvidia-smi", "--query-gpu=name", "--format=csv,noheader"],
                capture_output=True, text=True, timeout=8,
            )
            names = [line.strip() for line in completed.stdout.splitlines() if line.strip()]
        except (OSError, subprocess.SubprocessError):
            names = []
    if not names and IS_WINDOWS:
        try:
            completed = subprocess.run(
                ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command",
                 "Get-CimInstance Win32_VideoController | Select-Object -ExpandProperty Name"],
                capture_output=True, text=True, timeout=8,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            names = [line.strip() for line in completed.stdout.splitlines() if line.strip()]
        except (OSError, subprocess.SubprocessError):
            names = []
    return names


def _ollama_version() -> str:
    if not shutil.which("ollama"):
        return "در دسترس نبود"
    try:
        completed = subprocess.run(["ollama", "--version"], capture_output=True, text=True, timeout=8)
        return (completed.stdout or completed.stderr).strip().splitlines()[0][:120]
    except (OSError, subprocess.SubprocessError, IndexError):
        return "در دسترس نبود"


def environment_snapshot() -> dict[str, Any]:
    info: dict[str, Any] = {
        "سیستم‌عامل": f"{platform.system()} {platform.release()}",
        "نسخه Python": sys.version.split()[0],
        "اجرای Python": sys.executable,
        "مدل CPU": _cpu_model(),
        "هسته منطقی CPU": os.cpu_count(),
        "کارت‌های گرافیک": ", ".join(_gpu_names()) or "شناسایی نشد",
        "نسخه Ollama": _ollama_version(),
    }
    if psutil is not None:
        info["هسته فیزیکی CPU"] = psutil.cpu_count(logical=False)
        info["حافظه کل GB"] = round(psutil.virtual_memory().total / (1024 ** 3), 2)
    return info


@dataclass
class PerformanceReport:
    started_at: datetime
    finished_at: datetime
    total_seconds: float
    phases: dict[str, float] = field(default_factory=dict)
    timings: dict[str, float] = field(default_factory=dict)
    resources: dict[str, Any] = field(default_factory=dict)
    inference_resources: dict[str, Any] = field(default_factory=dict)
    gpu: dict[str, Any] = field(default_factory=dict)
    model: dict[str, Any] = field(default_factory=dict)
    environment: dict[str, Any] = field(default_factory=dict)
    limitations: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    sample_count: int = 0
    gpu_sample_count: int = 0
    gpu_success_count: int = 0
    input_sha256: str = ""
    output_path: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "started_at": self.started_at.isoformat(timespec="seconds"),
            "finished_at": self.finished_at.isoformat(timespec="seconds"),
            "total_seconds": round(self.total_seconds, 3),
            "phases": {key: round(value, 4) for key, value in self.phases.items()},
            "timings": {key: round(value, 4) for key, value in self.timings.items()},
            "resources": self.resources,
            "inference_resources": self.inference_resources,
            "gpu": self.gpu,
            "model": self.model,
            "environment": self.environment,
            "limitations": self.limitations,
            "errors": self.errors,
            "sample_count": self.sample_count,
            "gpu_sample_count": self.gpu_sample_count,
            "gpu_success_count": self.gpu_success_count,
            "input_sha256": self.input_sha256,
            "output_path": self.output_path,
        }


def build_performance_report(
    monitor: ResourceMonitor,
    settings: Settings,
    *,
    started_at: datetime,
    finished_at: datetime,
    total_seconds: float,
    timings: dict[str, float],
    ai_metadata: dict[str, Any] | None = None,
    input_sha256: str = "",
    output_path: str = "",
) -> PerformanceReport:
    samples = monitor.samples
    inference_samples = [sample for sample in samples if sample.get("phase") == "استنتاج مدل"]
    gpu_samples = monitor.gpu_samples
    successful_gpu = [sample for sample in gpu_samples if sample.available]

    resources = {
        "system_cpu_pct": stats([sample.get("system_cpu_pct") for sample in samples]),
        "system_ram_used_gb": stats([sample.get("ram_used_gb") for sample in samples]),
        "system_ram_available_gb": stats([sample.get("ram_available_gb") for sample in samples]),
        "python_rss_mb": stats([sample.get("python_rss_mb") for sample in samples]),
        "ollama_cpu_machine_pct": stats([sample.get("ollama_cpu_machine_pct") for sample in samples]),
        "ollama_rss_mb": stats([sample.get("ollama_rss_mb") for sample in samples]),
        "sample_count": len(samples),
    }
    inference_resources = {
        "system_cpu_pct": stats([sample.get("system_cpu_pct") for sample in inference_samples]),
        "system_ram_used_gb": stats([sample.get("ram_used_gb") for sample in inference_samples]),
        "ollama_cpu_machine_pct": stats([sample.get("ollama_cpu_machine_pct") for sample in inference_samples]),
        "ollama_rss_mb": stats([sample.get("ollama_rss_mb") for sample in inference_samples]),
        "sample_count": len(inference_samples),
    }
    gpu = {
        "available": bool(successful_gpu),
        "sample_count": len(gpu_samples),
        "success_count": len(successful_gpu),
        "ollama_engine_max_pct": stats([sample.ollama_engine_max_pct for sample in gpu_samples]),
        "all_engine_max_pct": stats([sample.all_engine_max_pct for sample in gpu_samples]),
        "ollama_dedicated_mb": stats([sample.ollama_dedicated_mb for sample in gpu_samples]),
        "ollama_shared_mb": stats([sample.ollama_shared_mb for sample in gpu_samples]),
        "all_process_dedicated_mb": stats([sample.all_process_dedicated_mb for sample in gpu_samples]),
        "sample_interval_sec": stats([sample.sample_interval_sec for sample in gpu_samples]),
        "query_elapsed_sec": stats([sample.query_elapsed_sec for sample in gpu_samples]),
        "engine_over_100_seen": any(sample.engine_over_100 for sample in gpu_samples),
        "last_error": next((sample.error for sample in reversed(gpu_samples) if sample.error), ""),
    }

    model = dict(ai_metadata or {})
    limitations = list(dict.fromkeys(list(settings.extra.get("limitations", [])) + monitor.limitations))
    limitations.extend(
        [
            "درصد GPU Engine درصد کل کارت گرافیک نیست؛ چند Engine مستقل روی یک کارت وجود دارد.",
            "انتساب VRAM به پردازش فقط وقتی انجام می‌شود که شمارنده ویندوز PID را در نام نمونه داشته باشد.",
            "مقدار «در دسترس نبود» به معنای صفر نیست؛ یعنی اندازه‌گیری ممکن نبوده است.",
            "تعداد توکن و سرعت تولید فقط وقتی گزارش می‌شود که Ollama این فیلدها را برگرداند.",
            "زمان بارگذاری مدل تنها زمانی معتبر است که مدل در همان درخواست از حافظه بارگذاری شده باشد.",
            "فاصله واقعی نمونه‌برداری ممکن است از بازه هدف بیشتر شود؛ مقدار واقعی گزارش می‌شود.",
        ]
    )
    if not IS_WINDOWS:
        limitations.append("این اجرا روی ویندوز نبود؛ شمارنده‌های GPU/VRAM در دسترس نبودند.")
    if monitor.detected_processes:
        limitations.append(
            "پردازش‌های Ollama بر اساس نام فرایند شناسایی شدند: "
            + "، ".join(f"{pid}:{name}" for pid, name in sorted(monitor.detected_processes.items()))
        )
    else:
        limitations.append("در طول اجرا پردازش Ollama/Runner شناسایی نشد؛ CPU/RAM آن اندازه‌گیری نشد.")

    return PerformanceReport(
        started_at=started_at,
        finished_at=finished_at,
        total_seconds=total_seconds,
        phases=dict(monitor.phase_durations),
        timings=dict(timings),
        resources=resources,
        inference_resources=inference_resources,
        gpu=gpu,
        model=model,
        environment=environment_snapshot(),
        limitations=limitations,
        errors=list(monitor.errors),
        sample_count=len(samples),
        gpu_sample_count=len(gpu_samples),
        gpu_success_count=len(successful_gpu),
        input_sha256=input_sha256,
        output_path=output_path,
    )


def _fmt(value: Any) -> Any:
    if value is None:
        return "در دسترس نبود"
    if isinstance(value, float):
        return round(value, 2)
    return value


def performance_rows(report: PerformanceReport) -> list[tuple[str, Any, str, str, str]]:
    """Rows for the Excel performance sheet: (label, value, unit, scope, note)."""
    rows: list[tuple[str, Any, str, str, str]] = []
    rows.append(("زمان شروع پردازش", report.started_at.strftime("%Y-%m-%d %H:%M:%S"), "", "کل اجرا", "پس از تأیید فایل و مدل"))
    rows.append(("زمان پایان پردازش", report.finished_at.strftime("%Y-%m-%d %H:%M:%S"), "", "کل اجرا", ""))
    rows.append(("زمان کل پردازش", round(report.total_seconds, 3), "ثانیه", "کل اجرا", "از خواندن فایل تا ذخیره خروجی"))
    for name, seconds in report.phases.items():
        rows.append((f"زمان مرحله: {name}", round(seconds, 4), "ثانیه", "برنامه", "اندازه‌گیری داخلی مرحله"))
    for name, seconds in report.timings.items():
        if name in report.phases:
            continue
        rows.append((name, round(seconds, 4), "ثانیه", "برنامه", "زمان‌سنجی مرحله"))

    model = report.model
    model_ran = model.get("mode") == "model"
    not_run = "انجام نشد (بدون مدل)"
    rows.append(("مدل انتخاب‌شده", model.get("model") or "بدون مدل", "", "مدل", model.get("mode", "")))
    rows.append(("زمان درخواست مدل (کلاینت)", _fmt(model.get("client_elapsed_seconds")) if model_ran else not_run, "ثانیه", "مدل", "رفت‌وبرگشت HTTP"))
    rows.append(("زمان کل Ollama", _fmt(model.get("total_duration_seconds")) if model_ran else not_run, "ثانیه", "مدل", "از پاسخ API"))
    rows.append(("زمان بارگذاری مدل", _fmt(model.get("load_duration_seconds")) if model_ran else not_run, "ثانیه", "مدل", "در صورت گزارش API"))
    rows.append(("زمان پردازش ورودی", _fmt(model.get("prompt_eval_duration_seconds")) if model_ran else not_run, "ثانیه", "مدل", ""))
    rows.append(("زمان تولید پاسخ", _fmt(model.get("eval_duration_seconds")) if model_ran else not_run, "ثانیه", "مدل", ""))
    rows.append(("توکن ورودی", _fmt(model.get("prompt_tokens")) if model_ran else not_run, "توکن", "مدل", "prompt_eval_count"))
    rows.append(("توکن خروجی", _fmt(model.get("output_tokens")) if model_ran else not_run, "توکن", "مدل", "eval_count"))
    rows.append(("سرعت تولید", _fmt(model.get("tokens_per_second")) if model_ran else not_run, "توکن/ثانیه", "مدل", "eval_count ÷ eval_duration"))
    rows.append(("تعداد تلاش درخواست", _fmt(model.get("attempts")) if model_ran else not_run, "بار", "مدل", "با احتساب تلاش مجدد"))
    if model.get("error"):
        rows.append(("خطای مدل", model.get("error"), "", "مدل", "تحلیل بدون مدل انجام شد"))

    for key, label, unit in (
        ("system_cpu_pct", "CPU سیستم", "%"),
        ("system_ram_used_gb", "RAM مصرفی سیستم", "GB"),
        ("system_ram_available_gb", "RAM آزاد سیستم", "GB"),
        ("python_rss_mb", "RAM برنامه Python", "MB"),
        ("ollama_cpu_machine_pct", "CPU پردازش Ollama", "% کل CPU"),
        ("ollama_rss_mb", "RAM پردازش Ollama", "MB"),
    ):
        block = report.resources.get(key, {})
        rows.append((f"میانگین {label}", _fmt(block.get("avg")), unit, "کل اجرا", f"{block.get('count', 0)} نمونه"))
        rows.append((f"حداکثر {label}", _fmt(block.get("max")), unit, "کل اجرا", ""))
        rows.append((f"صدک ۹۵ {label}", _fmt(block.get("p95")), unit, "کل اجرا", ""))

    inference = report.inference_resources
    for key, label, unit in (
        ("system_cpu_pct", "CPU سیستم", "%"),
        ("system_ram_used_gb", "RAM مصرفی سیستم", "GB"),
        ("ollama_cpu_machine_pct", "CPU پردازش Ollama", "% کل CPU"),
        ("ollama_rss_mb", "RAM پردازش Ollama", "MB"),
    ):
        block = inference.get(key, {})
        rows.append((f"میانگین {label} در زمان استنتاج", _fmt(block.get("avg")), unit, "فقط استنتاج", f"{block.get('count', 0)} نمونه"))

    gpu = report.gpu
    rows.append(("پایش GPU فعال بود", "بله" if gpu.get("available") else "خیر", "", "GPU", gpu.get("last_error") or ""))
    rows.append(("نمونه‌های موفق GPU", gpu.get("success_count"), "نمونه", "GPU", f"از {gpu.get('sample_count')} نمونه"))
    rows.append(("حداکثر GPU Engine منتسب به Ollama", _fmt(gpu.get("ollama_engine_max_pct", {}).get("max")), "% یک Engine", "GPU", "درصد کل کارت نیست"))
    rows.append(("حداکثر Dedicated VRAM منتسب به Ollama", _fmt(gpu.get("ollama_dedicated_mb", {}).get("max")), "MB", "GPU", "بر اساس PID"))
    rows.append(("حداکثر Shared GPU Memory منتسب به Ollama", _fmt(gpu.get("ollama_shared_mb", {}).get("max")), "MB", "GPU", "بر اساس PID"))
    rows.append(("حداکثر Dedicated Memory همه پردازش‌ها", _fmt(gpu.get("all_process_dedicated_mb", {}).get("max")), "MB", "GPU", "مختص Ollama نیست"))
    rows.append(("میانه فاصله نمونه‌های GPU", _fmt(gpu.get("sample_interval_sec", {}).get("median")), "ثانیه", "GPU", "فاصله واقعی"))
    rows.append(("میانه زمان پرس‌وجوی شمارنده", _fmt(gpu.get("query_elapsed_sec", {}).get("median")), "ثانیه", "GPU", ""))
    rows.append(("GPU Engine بالای ۱۰۰٪ دیده شد", "بله" if gpu.get("engine_over_100_seen") else "خیر", "", "GPU", "مجموع چند Engine"))
    return rows
