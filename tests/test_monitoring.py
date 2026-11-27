from __future__ import annotations

import json
import time
from datetime import datetime

from local_office_ai.monitoring.performance_report import (
    build_performance_report,
    performance_rows,
    stats,
)
from local_office_ai.monitoring.resource_monitor import IS_WINDOWS, GpuSample, ResourceMonitor


def test_stats_never_fakes_zero():
    assert stats([])["avg"] is None
    assert stats([None, None])["max"] is None
    result = stats([1, 2, 3, 4])
    assert result["avg"] == 2.5
    assert result["max"] == 4
    assert result["min"] == 1
    assert result["count"] == 4
    assert result["p95"] == 4


def test_system_sample_shape(settings):
    monitor = ResourceMonitor(settings)
    monitor.start()
    try:
        time.sleep(0.15)
        sample = monitor.sample_system_once()
    finally:
        monitor.stop()
    for key in (
        "timestamp",
        "phase",
        "system_cpu_pct",
        "ram_used_gb",
        "ram_available_gb",
        "ram_total_gb",
        "python_rss_mb",
        "ollama_rss_mb",
        "ollama_cpu_machine_pct",
    ):
        assert key in sample
    assert sample["ram_total_gb"] > 0
    assert sample["python_rss_mb"] is not None


def test_gpu_unavailable_is_none_not_zero(settings):
    monitor = ResourceMonitor(settings)
    sample = monitor.sample_gpu_once()
    if IS_WINDOWS:
        assert isinstance(sample, GpuSample)
    else:
        assert sample.available is False
        assert sample.ollama_engine_max_pct is None
        assert sample.ollama_dedicated_mb is None
        assert sample.error


def test_phase_scope_records_durations(settings):
    monitor = ResourceMonitor(settings)
    with monitor.phase_scope("مرحله الف"):
        time.sleep(0.02)
    with monitor.phase_scope("مرحله الف"):
        time.sleep(0.02)
    assert monitor.phase_durations["مرحله الف"] >= 0.03
    assert monitor.phase == "آماده"


def test_performance_report_marks_missing_gpu(settings):
    monitor = ResourceMonitor(settings)
    monitor.limitations.append("محدودیت آزمایشی")
    monitor.gpu_samples.append(
        GpuSample(timestamp=time.time(), phase="استنتاج مدل", available=False, error="not available")
    )
    report = build_performance_report(
        monitor,
        settings,
        started_at=datetime(2026, 1, 1, 10, 0, 0),
        finished_at=datetime(2026, 1, 1, 10, 0, 5),
        total_seconds=5.0,
        timings={"خواندن فایل": 0.2},
        ai_metadata={"model": "qwen3.5:9b", "output_tokens": 10, "tokens_per_second": 5.0},
        input_sha256="abc",
    )
    assert report.gpu["available"] is False
    assert report.gpu["ollama_dedicated_mb"]["max"] is None
    rows = performance_rows(report)
    labels = [row[0] for row in rows]
    assert "حداکثر Dedicated VRAM منتسب به Ollama" in labels
    vram_row = next(row for row in rows if row[0] == "حداکثر Dedicated VRAM منتسب به Ollama")
    assert vram_row[1] == "در دسترس نبود"
    assert "محدودیت آزمایشی" in report.limitations
    assert any("GPU Engine" in item for item in report.limitations)
    # The report must be JSON serializable.
    json.dumps(report.to_dict(), ensure_ascii=False)


def test_monitor_threads_collect_samples(settings):
    settings.monitor_interval_seconds = 0.05
    settings.gpu_interval_seconds = 0.05
    monitor = ResourceMonitor(settings)
    monitor.start()
    time.sleep(0.25)
    elapsed = monitor.stop()
    assert elapsed >= 0.1
    assert len(monitor.samples) >= 2
    assert monitor.started_at is not None and monitor.finished_at is not None
