"""Resource monitoring for the processing run.

Rules that differ from the old implementation:
* a value that cannot be measured is ``None``, never ``0``;
* a single GPU *engine* percentage is never labelled as total GPU usage;
* VRAM is only attributed to a process when the Windows counter names its PID;
* every limitation is written into the report instead of being hidden.
"""

from __future__ import annotations

import contextlib
import json
import os
import re
import subprocess
import threading
import time
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Iterator

try:  # psutil is optional at import time so the engine can explain itself
    import psutil
except ImportError:  # pragma: no cover - dependency is declared in requirements
    psutil = None  # type: ignore[assignment]

from local_office_ai.core.logging_config import get_logger
from local_office_ai.core.settings import Settings

LOGGER = get_logger(__name__)

OLLAMA_PROCESS_NAMES = {
    "ollama",
    "ollama.exe",
    "ollama app.exe",
    "ollama_llama_server.exe",
    "llama-server",
    "llama-server.exe",
    "llama_server.exe",
    "llama-runner",
    "llama_runner.exe",
    "ollama-runner.exe",
    "ollama_runner.exe",
}

IS_WINDOWS = os.name == "nt"


@dataclass
class GpuSample:
    timestamp: float
    phase: str
    available: bool
    error: str = ""
    query_elapsed_sec: float | None = None
    sample_interval_sec: float | None = None
    ollama_engine_max_pct: float | None = None
    ollama_engine_sum_pct: float | None = None
    all_engine_max_pct: float | None = None
    all_engine_sum_pct: float | None = None
    ollama_dedicated_mb: float | None = None
    ollama_shared_mb: float | None = None
    all_process_dedicated_mb: float | None = None
    all_process_shared_mb: float | None = None
    ollama_pids: str = ""
    adapter_count: int | None = None
    engine_instance_count: int | None = None
    engine_over_100: bool = False

    def to_dict(self) -> dict[str, Any]:
        return dict(self.__dict__)


@dataclass
class ResourceMonitor:
    settings: Settings
    samples: list[dict[str, Any]] = field(default_factory=list)
    gpu_samples: list[GpuSample] = field(default_factory=list)
    phase_durations: dict[str, float] = field(default_factory=dict)
    detected_processes: dict[int, str] = field(default_factory=dict)
    limitations: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    started_at: datetime | None = None
    finished_at: datetime | None = None

    def __post_init__(self) -> None:
        self._stop = threading.Event()
        self._threads: list[threading.Thread] = []
        self._phase_stack: list[str] = []
        self._phase_start: dict[str, float] = {}
        self._start_perf: float | None = None
        self._previous_cpu: dict[int, float] = {}
        self._previous_sample: float | None = None
        self._last_gpu_start: float | None = None
        self._gpu_supported = False
        self._self_pid = os.getpid()

    # ------------------------------------------------------------------ #
    @property
    def phase(self) -> str:
        return self._phase_stack[-1] if self._phase_stack else "آماده"

    def set_phase(self, name: str) -> None:
        self._phase_stack = [name]

    @contextlib.contextmanager
    def phase_scope(self, name: str) -> Iterator[None]:
        self._phase_stack.append(name)
        started = time.perf_counter()
        try:
            yield
        finally:
            elapsed = time.perf_counter() - started
            self.phase_durations[name] = self.phase_durations.get(name, 0.0) + elapsed
            if self._phase_stack and self._phase_stack[-1] == name:
                self._phase_stack.pop()

    # ------------------------------------------------------------------ #
    def start(self) -> None:
        if psutil is None:
            self.limitations.append("psutil نصب نیست؛ پایش CPU/RAM انجام نشد.")
            return
        self.started_at = datetime.now()
        self._start_perf = time.perf_counter()
        psutil.cpu_percent(interval=None)
        self._threads.append(
            threading.Thread(target=self._system_loop, name="monitor-system", daemon=True)
        )
        self._threads[0].start()
        if self.settings.enable_gpu_monitoring:
            self._threads.append(
                threading.Thread(target=self._gpu_loop, name="monitor-gpu", daemon=True)
            )
            self._threads[1].start()

    def stop(self) -> float:
        self._stop.set()
        for thread in self._threads:
            thread.join(timeout=12)
        self._threads.clear()
        self.finished_at = datetime.now()
        elapsed = time.perf_counter() - (self._start_perf or time.perf_counter())
        LOGGER.info("monitor stopped after %.2fs (%d samples)", elapsed, len(self.samples))
        return elapsed

    # ------------------------------------------------------------------ #
    def _system_loop(self) -> None:
        while not self._stop.is_set():
            try:
                self.samples.append(self.sample_system_once())
            except Exception as exc:  # noqa: BLE001 - monitoring must not kill the run
                self.errors.append(f"system sample failed: {exc}")
            self._stop.wait(self.settings.monitor_interval_seconds)

    def _gpu_loop(self) -> None:
        if not IS_WINDOWS:
            self.gpu_samples.append(
                GpuSample(
                    timestamp=time.time(),
                    phase=self.phase,
                    available=False,
                    error="شمارنده‌های GPU ویندوز در این سیستم‌عامل موجود نیست.",
                )
            )
            self.limitations.append(
                "پایش GPU فقط روی ویندوز و با Windows Performance Counters انجام می‌شود."
            )
            return
        while not self._stop.is_set():
            tick = time.perf_counter()
            try:
                sample = self.sample_gpu_once()
                sample.sample_interval_sec = (
                    tick - self._last_gpu_start if self._last_gpu_start is not None else None
                )
                self.gpu_samples.append(sample)
            except Exception as exc:  # noqa: BLE001
                self.gpu_samples.append(
                    GpuSample(
                        timestamp=time.time(),
                        phase=self.phase,
                        available=False,
                        error=str(exc)[:300],
                    )
                )
                self.errors.append(f"gpu sample failed: {exc}")
            self._last_gpu_start = tick
            remaining = self.settings.gpu_interval_seconds - (time.perf_counter() - tick)
            self._stop.wait(max(0.0, remaining))

    # ------------------------------------------------------------------ #
    def _ollama_processes(self) -> list[Any]:
        found = []
        if psutil is None:
            return found
        for process in psutil.process_iter(["pid", "name"]):
            try:
                pid = int(process.info["pid"])
                if pid == self._self_pid:
                    continue
                name = str(process.info.get("name") or "").strip()
                lowered = name.lower()
                if lowered in OLLAMA_PROCESS_NAMES or lowered.startswith("ollama") or lowered.startswith("llama"):
                    found.append(process)
                    self.detected_processes[pid] = name
            except (psutil.NoSuchProcess, psutil.AccessDenied, TypeError, ValueError):
                continue
        return found

    def sample_system_once(self) -> dict[str, Any]:
        if psutil is None:
            return {"timestamp": time.time(), "phase": self.phase, "sample_error": "psutil missing"}
        now = time.perf_counter()
        memory = psutil.virtual_memory()
        system_cpu = psutil.cpu_percent(interval=None)
        logical = max(psutil.cpu_count(logical=True) or 1, 1)

        ollama_cpu_seconds = 0.0
        ollama_rss = 0
        ollama_pids: list[int] = []
        for process in self._ollama_processes():
            try:
                pid = process.pid
                cpu = process.cpu_times()
                total = float(cpu.user + cpu.system)
                previous = self._previous_cpu.get(pid)
                delta = max(0.0, total - previous) if previous is not None else 0.0
                self._previous_cpu[pid] = total
                ollama_cpu_seconds += delta
                ollama_rss += process.memory_info().rss
                ollama_pids.append(pid)
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue

        elapsed = (now - self._previous_sample) if self._previous_sample is not None else 0.0
        self._previous_sample = now
        machine_pct = (ollama_cpu_seconds / elapsed * 100.0 / logical) if elapsed > 0 else None

        try:
            python_rss = psutil.Process(self._self_pid).memory_info().rss / (1024 ** 2)
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            python_rss = None

        return {
            "timestamp": time.time(),
            "elapsed_sec": (now - self._start_perf) if self._start_perf else None,
            "phase": self.phase,
            "system_cpu_pct": system_cpu,
            "ram_used_gb": (memory.total - memory.available) / (1024 ** 3),
            "ram_available_gb": memory.available / (1024 ** 3),
            "ram_total_gb": memory.total / (1024 ** 3),
            "ram_used_pct": memory.percent,
            "python_rss_mb": python_rss,
            "ollama_cpu_machine_pct": machine_pct,
            "ollama_rss_mb": (ollama_rss / (1024 ** 2)) if ollama_pids else None,
            "ollama_pids": ",".join(str(pid) for pid in sorted(ollama_pids)),
            "ollama_process_count": len(ollama_pids),
            "logical_cpus": logical,
        }

    # ------------------------------------------------------------------ #
    def sample_gpu_once(self) -> GpuSample:
        started = time.perf_counter()
        payload = self._query_windows_gpu()
        elapsed = time.perf_counter() - started
        if payload is None:
            return GpuSample(
                timestamp=time.time(),
                phase=self.phase,
                available=False,
                error="شمارنده GPU پاسخ نداد.",
                query_elapsed_sec=elapsed,
            )

        live_roles = set(self.detected_processes)
        ollama_engines: list[float] = []
        all_engines: list[float] = []
        ollama_dedicated = ollama_shared = all_dedicated = all_shared = 0.0
        adapters: set[str] = set()
        engine_instances = 0
        for item in payload:
            instance = str(item.get("instance") or "")
            kind = str(item.get("kind") or "")
            try:
                value = max(0.0, float(item.get("value") or 0.0))
            except (TypeError, ValueError):
                continue
            adapters.add(str(item.get("adapter") or instance[:40]))
            pid_match = re.search(r"pid_(\d+)_", instance, re.IGNORECASE)
            pid = int(pid_match.group(1)) if pid_match else None
            is_ollama = pid is not None and pid in live_roles
            if kind == "engine":
                engine_instances += 1
                all_engines.append(value)
                if is_ollama:
                    ollama_engines.append(value)
            elif kind == "dedicated":
                all_dedicated += value
                if is_ollama:
                    ollama_dedicated += value
            elif kind == "shared":
                all_shared += value
                if is_ollama:
                    ollama_shared += value

        self._gpu_supported = True
        over_100 = any(value > 100.0 for value in all_engines)
        if over_100:
            self.limitations.append(
                "مقدار GPU Engine در بعضی نمونه‌ها از ۱۰۰٪ بیشتر شد؛ این عدد مجموع چند Engine است "
                "و درصد کل کارت گرافیک نیست."
            )
        return GpuSample(
            timestamp=time.time(),
            phase=self.phase,
            available=bool(payload),
            query_elapsed_sec=elapsed,
            ollama_engine_max_pct=max(ollama_engines) if ollama_engines else None,
            ollama_engine_sum_pct=sum(ollama_engines) if ollama_engines else None,
            all_engine_max_pct=max(all_engines) if all_engines else None,
            all_engine_sum_pct=sum(all_engines) if all_engines else None,
            ollama_dedicated_mb=(ollama_dedicated / (1024 ** 2)) if ollama_dedicated else None,
            ollama_shared_mb=(ollama_shared / (1024 ** 2)) if ollama_shared else None,
            all_process_dedicated_mb=(all_dedicated / (1024 ** 2)) if all_dedicated else None,
            all_process_shared_mb=(all_shared / (1024 ** 2)) if all_shared else None,
            ollama_pids=",".join(str(pid) for pid in sorted(live_roles)),
            adapter_count=len(adapters) if adapters else None,
            engine_instance_count=engine_instances,
            engine_over_100=over_100,
        )

    @staticmethod
    def _query_windows_gpu() -> list[dict[str, Any]] | None:
        if not IS_WINDOWS:
            return None
        script = r"""
$ErrorActionPreference = 'Stop'
$paths = @(
  '\GPU Engine(*)\Utilization Percentage',
  '\GPU Process Memory(*)\Dedicated Usage',
  '\GPU Process Memory(*)\Shared Usage'
)
try {
  $result = Get-Counter -Counter $paths -ErrorAction Stop
  $out = @()
  foreach ($s in $result.CounterSamples) {
    $kind = 'other'
    if ($s.Path -like '*GPU Engine*') { $kind = 'engine' }
    elseif ($s.Path -like '*Dedicated Usage*') { $kind = 'dedicated' }
    elseif ($s.Path -like '*Shared Usage*') { $kind = 'shared' }
    $adapter = ''
    if ($s.InstanceName -match '(luid_0x[0-9a-fA-F]+_0x[0-9a-fA-F]+)') { $adapter = $Matches[1] }
    $out += [pscustomobject]@{Kind=$kind; Instance=$s.InstanceName; Adapter=$adapter; Value=[double]$s.CookedValue}
  }
  ConvertTo-Json -InputObject $out -Compress -Depth 4
} catch {
  [pscustomobject]@{__error=$_.Exception.Message} | ConvertTo-Json -Compress
}
"""
        try:
            completed = subprocess.run(
                ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", script],
                capture_output=True,
                text=True,
                timeout=12,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except (OSError, subprocess.SubprocessError):
            return None
        raw = (completed.stdout or "").strip()
        if not raw:
            return None
        try:
            payload = json.loads(raw)
        except json.JSONDecodeError:
            return None
        if isinstance(payload, dict) and payload.get("__error"):
            return None
        items = payload if isinstance(payload, list) else [payload]
        return [
            {
                "kind": str(item.get("Kind") or "").lower(),
                "instance": str(item.get("Instance") or ""),
                "adapter": str(item.get("Adapter") or ""),
                "value": item.get("Value"),
            }
            for item in items
            if isinstance(item, dict)
        ]
