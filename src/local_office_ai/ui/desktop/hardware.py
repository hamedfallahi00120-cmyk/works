"""Independent hardware readings for the optional monitor panel.

CPU/RAM come from psutil; NVIDIA can use ``nvidia-smi``; Linux AMD/Intel can
use sysfs; Windows sensor temperatures and AMD GPU telemetry come from the
local LibreHardwareMonitor REST endpoint (127.0.0.1:8085/data.json). No WMI
namespace or external Python package is required. Nothing is downloaded or
installed here. Missing readings remain ``None`` and the UI shows
«در دسترس نیست» rather than inventing zero values.

These readings are for the interface only; the report writer never sees them.
"""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import time
import urllib.error
import urllib.request
from math import isfinite
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

try:  # psutil is a project dependency, but the panel must degrade gracefully
    import psutil
except ImportError:  # pragma: no cover - exercised only on a broken install
    psutil = None  # type: ignore[assignment]

UNAVAILABLE = "در دسترس نیست"
IS_WINDOWS = os.name == "nt"
_CREATE_NO_WINDOW = 0x08000000 if IS_WINDOWS else 0


@dataclass
class GpuReading:
    index: int
    name: str
    utilization: float | None = None
    temperature: float | None = None
    vram_used_mb: float | None = None
    vram_total_mb: float | None = None
    source: str = ""


@dataclass
class CpuReading:
    utilization: float | None = None
    temperature: float | None = None
    temperature_source: str = ""
    physical_cores: int | None = None
    logical_cores: int | None = None
    frequency_mhz: float | None = None


@dataclass
class RamReading:
    percent: float | None = None
    used_bytes: int | None = None
    total_bytes: int | None = None
    available_bytes: int | None = None


@dataclass
class HardwareSnapshot:
    cpu: CpuReading = field(default_factory=CpuReading)
    ram: RamReading = field(default_factory=RamReading)
    gpus: list[GpuReading] = field(default_factory=list)
    providers: dict[str, str] = field(default_factory=dict)
    notes: list[str] = field(default_factory=list)
    timestamp: float = 0.0


class HardwareMonitor:
    """Sample the machine on demand; safe to call from one worker thread."""

    def __init__(self, *, enable_gpu: bool = True) -> None:
        self.enable_gpu = enable_gpu
        self._nvidia_path: str | None = None
        self._nvidia_checked_at = 0.0
        self._lhm_retry_after = 0.0
        self._lhm_available: bool | None = None
        self._cpu_primed = False
        self._sample_lhm: list[dict[str, Any]] | None = None

    # ------------------------------------------------------------------ #
    def sample(self) -> HardwareSnapshot:
        snapshot = HardwareSnapshot(timestamp=time.time())
        self._sample_lhm = None
        self._read_cpu(snapshot)
        self._read_ram(snapshot)
        if self.enable_gpu:
            self._read_gpus(snapshot)
            if not snapshot.gpus:
                snapshot.gpus = self._lhm_gpus()
                if snapshot.gpus:
                    snapshot.providers["gpu"] = "LibreHardwareMonitor REST"
            if snapshot.gpus:
                snapshot.providers.setdefault("gpu", snapshot.gpus[0].source or UNAVAILABLE)
            else:
                snapshot.providers["gpu"] = UNAVAILABLE
        else:
            snapshot.providers["gpu"] = "غیرفعال در تنظیمات"
        if snapshot.cpu.temperature is None and IS_WINDOWS:
            self._read_cpu_temperature_from_lhm(snapshot)
        snapshot.providers.setdefault("cpu", "psutil" if psutil is not None else UNAVAILABLE)
        snapshot.providers.setdefault("ram", "psutil" if psutil is not None else UNAVAILABLE)
        snapshot.providers.setdefault("cpu_temp", snapshot.cpu.temperature_source or UNAVAILABLE)
        if snapshot.cpu.temperature is None:
            snapshot.notes.append(
                "دمای CPU در این سیستم از سنسورهای در دسترس خوانده نشد."
            )
        return snapshot

    # ------------------------------------------------------------------ #
    def _read_cpu(self, snapshot: HardwareSnapshot) -> None:
        cpu = snapshot.cpu
        if psutil is None:
            return
        try:
            if not self._cpu_primed:
                psutil.cpu_percent(interval=None)
                self._cpu_primed = True
            cpu.utilization = float(psutil.cpu_percent(interval=None))
        except Exception:  # noqa: BLE001 - a sensor failure is not a crash
            cpu.utilization = None
        try:
            cpu.physical_cores = psutil.cpu_count(logical=False)
            cpu.logical_cores = psutil.cpu_count(logical=True)
        except Exception:  # noqa: BLE001
            pass
        try:
            frequency = psutil.cpu_freq()
            if frequency is not None and frequency.current:
                cpu.frequency_mhz = float(frequency.current)
        except Exception:  # noqa: BLE001
            pass
        temperature, source = self._psutil_cpu_temperature()
        cpu.temperature = temperature
        cpu.temperature_source = source

    @staticmethod
    def _psutil_cpu_temperature() -> tuple[float | None, str]:
        if psutil is None or IS_WINDOWS:
            return None, ""
        try:
            groups = psutil.sensors_temperatures() or {}
        except Exception:  # noqa: BLE001
            return None, ""
        preferred = ("coretemp", "k10temp", "zenpower", "cpu_thermal", "acpitz")
        for name in preferred:
            entries = groups.get(name) or []
            values = [entry.current for entry in entries if entry.current is not None]
            if not values:
                continue
            package = [
                entry.current
                for entry in entries
                if entry.current is not None
                and str(entry.label or "").lower() in {"package id 0", "tctl", "tdie", "cpu"}
            ]
            return float(max(package or values)), f"psutil ({name})"
        for name, entries in groups.items():
            lowered = name.lower()
            if not any(token in lowered for token in ("cpu", "core", "k10", "zen", "soc", "package")):
                continue
            values = [entry.current for entry in entries if entry.current is not None]
            if values:
                return float(max(values)), f"psutil ({name})"
        return None, ""

    def _read_ram(self, snapshot: HardwareSnapshot) -> None:
        if psutil is None:
            return
        try:
            memory = psutil.virtual_memory()
        except Exception:  # noqa: BLE001
            return
        snapshot.ram = RamReading(
            percent=float(memory.percent),
            used_bytes=int(memory.used),
            total_bytes=int(memory.total),
            available_bytes=int(memory.available),
        )

    # ------------------------------------------------------------------ #
    def _read_gpus(self, snapshot: HardwareSnapshot) -> None:
        gpus = self._nvidia_gpus()
        if gpus:
            snapshot.gpus = gpus
            snapshot.providers["gpu"] = "nvidia-smi"
            return
        if not IS_WINDOWS:
            gpus = self._sysfs_gpus()
            if gpus:
                snapshot.gpus = gpus
                snapshot.providers["gpu"] = "sysfs (Linux)"

    # -- NVIDIA --------------------------------------------------------- #
    def _find_nvidia_smi(self) -> str | None:
        now = time.monotonic()
        if self._nvidia_path is not None:
            return self._nvidia_path
        if now - self._nvidia_checked_at < 60:
            return None
        self._nvidia_checked_at = now
        found = shutil.which("nvidia-smi")
        if not found and IS_WINDOWS:
            candidate = Path(r"C:\Program Files\NVIDIA Corporation\NVSMI\nvidia-smi.exe")
            if candidate.is_file():
                found = str(candidate)
        self._nvidia_path = found
        return found

    def _nvidia_gpus(self) -> list[GpuReading]:
        executable = self._find_nvidia_smi()
        if not executable:
            return []
        query = "index,name,utilization.gpu,temperature.gpu,memory.used,memory.total"
        try:
            completed = subprocess.run(
                [
                    executable,
                    f"--query-gpu={query}",
                    "--format=csv,noheader,nounits",
                ],
                capture_output=True,
                text=True,
                timeout=4,
                check=False,
                creationflags=_CREATE_NO_WINDOW,
            )
        except (OSError, subprocess.SubprocessError):
            return []
        if completed.returncode != 0:
            return []
        readings: list[GpuReading] = []
        for line in completed.stdout.strip().splitlines():
            parts = [part.strip() for part in line.split(",", 5)]
            if len(parts) < 6:
                continue
            readings.append(
                GpuReading(
                    index=_int_or_default(parts[0], len(readings)),
                    name=parts[1] or f"GPU {len(readings) + 1}",
                    utilization=_float_or_none(parts[2]),
                    temperature=_float_or_none(parts[3]),
                    vram_used_mb=_float_or_none(parts[4]),
                    vram_total_mb=_float_or_none(parts[5]),
                    source="nvidia-smi",
                )
            )
        return readings

    # -- Linux sysfs ---------------------------------------------------- #
    @staticmethod
    def _sysfs_gpus() -> list[GpuReading]:
        root = Path("/sys/class/drm")
        if not root.is_dir():
            return []
        readings: list[GpuReading] = []
        for entry in sorted(root.iterdir()):
            if not re.fullmatch(r"card\d+", entry.name):
                continue
            device = entry / "device"
            if not device.is_dir():
                continue
            driver = ""
            try:
                for line in (device / "uevent").read_text(encoding="utf-8").splitlines():
                    if line.startswith("DRIVER="):
                        driver = line.split("=", 1)[1].strip()
                        break
            except OSError:
                pass
            utilization = _read_float_file(device / "gpu_busy_percent")
            used = _read_float_file(device / "mem_info_vram_used")
            total = _read_float_file(device / "mem_info_vram_total")
            temperature = None
            for hwmon in sorted((device / "hwmon").glob("hwmon*")):
                raw = _read_float_file(hwmon / "temp1_input")
                if raw is not None:
                    temperature = raw / 1000.0
                    break
            if all(value is None for value in (utilization, used, total, temperature)):
                continue
            vendor = {
                "amdgpu": "AMD Radeon",
                "i915": "Intel Graphics",
                "xe": "Intel Graphics",
                "nouveau": "NVIDIA (nouveau)",
            }.get(driver, driver or "GPU")
            readings.append(
                GpuReading(
                    index=len(readings),
                    name=f"{vendor} ({entry.name})",
                    utilization=utilization,
                    temperature=temperature,
                    vram_used_mb=used / (1024 * 1024) if used is not None else None,
                    vram_total_mb=total / (1024 * 1024) if total is not None else None,
                    source="sysfs (Linux)",
                )
            )
        return readings

    # -- LibreHardwareMonitor local REST API ----------------------------- #
    LHM_REST_URL = "http://127.0.0.1:8085/data.json"

    def _lhm_sensors(self) -> list[dict[str, Any]]:
        """Return flattened sensor rows from LHM's local JSON tree.

        LHM 0.9.6 exposes the JSON endpoint even when its WMI namespace is not
        registered. A short timeout keeps the optional monitor responsive.
        The response is cached for one hardware sample so CPU and GPU parsing
        share a single local HTTP request.
        """
        if self._sample_lhm is not None:
            return self._sample_lhm
        now = time.monotonic()
        if self._lhm_available is False and now < self._lhm_retry_after:
            return []
        if not IS_WINDOWS:
            self._lhm_available = False
            return []

        rows = self._query_lhm_rest(self.LHM_REST_URL)
        self._lhm_available = bool(rows)
        self._sample_lhm = rows
        if not rows:
            self._lhm_retry_after = now + 10.0
        else:
            self._lhm_retry_after = 0.0
        return rows

    @staticmethod
    def _query_lhm_rest(url: str) -> list[dict[str, Any]]:
        """Fetch and flatten the LHM tree from localhost only."""
        if not url.startswith("http://127.0.0.1:") and not url.startswith("http://localhost:"):
            return []
        request = urllib.request.Request(
            url,
            headers={"Accept": "application/json", "Cache-Control": "no-cache"},
            method="GET",
        )
        try:
            with urllib.request.urlopen(request, timeout=2.0) as response:
                if response.status != 200:
                    return []
                payload = json.loads(response.read().decode("utf-8-sig"))
        except (OSError, TimeoutError, urllib.error.URLError, json.JSONDecodeError, UnicodeDecodeError):
            return []
        if not isinstance(payload, dict):
            return []
        return HardwareMonitor._flatten_lhm_payload(payload)

    @staticmethod
    def _flatten_lhm_payload(payload: dict[str, Any]) -> list[dict[str, Any]]:
        """Flatten the documented LHM tree while retaining hardware ancestry."""
        flattened: list[dict[str, Any]] = []

        def visit(node: Any, path: tuple[str, ...] = ()) -> None:
            if not isinstance(node, dict):
                return
            text = str(node.get("Text") or "").strip()
            current_path = path + ((text,) if text else ())
            sensor_type = str(node.get("Type") or "").strip()
            if sensor_type and ("RawValue" in node or "Value" in node):
                flattened.append(
                    {
                        "Name": text,
                        "SensorType": sensor_type,
                        "Value": node.get("Value"),
                        "RawValue": node.get("RawValue"),
                        "Path": current_path,
                        "PathText": " / ".join(current_path),
                        "SensorId": node.get("SensorId"),
                    }
                )
            children = node.get("Children")
            if isinstance(children, list):
                for child in children:
                    visit(child, current_path)

        visit(payload)
        return flattened

    def _lhm_gpus(self) -> list[GpuReading]:
        sensors = self._lhm_sensors()
        if not sensors:
            return []

        groups: dict[str, list[dict[str, Any]]] = {}
        display_names: dict[str, str] = {}
        for sensor in sensors:
            path = tuple(sensor.get("Path") or ())
            hardware = _lhm_hardware_name(path, "gpu")
            if not hardware:
                continue
            key = hardware.casefold()
            groups.setdefault(key, []).append(sensor)
            display_names[key] = hardware

        readings: list[GpuReading] = []
        for key, rows in groups.items():
            load = _pick_lhm_sensor(rows, "Load", ("gpu core", "gpu", "core"))
            temperature = _pick_lhm_sensor(rows, "Temperature", ("gpu core", "core"))
            used = _pick_lhm_memory(rows, used=True)
            total = _pick_lhm_memory(rows, used=False)
            if all(value is None for value in (load, temperature, used, total)):
                continue
            name = display_names[key]
            readings.append(
                GpuReading(
                    index=len(readings),
                    name=name,
                    utilization=load,
                    temperature=temperature,
                    vram_used_mb=used,
                    vram_total_mb=total,
                    source="LibreHardwareMonitor REST",
                )
            )

        # The user has one discrete AMD card. If the integrated Intel adapter
        # also exposes sensors, prefer the discrete Radeon rather than showing
        # it as a second physical card in the compact monitor panel.
        discrete = [item for item in readings if any(token in item.name.casefold() for token in ("radeon", "nvidia", "geforce"))]
        return discrete or readings

    def _read_cpu_temperature_from_lhm(self, snapshot: HardwareSnapshot) -> None:
        sensors = self._lhm_sensors()
        if not sensors:
            return
        cpu_rows = [
            row for row in sensors
            if _lhm_hardware_name(tuple(row.get("Path") or ()), "cpu")
            and str(row.get("SensorType") or "").casefold() == "temperature"
        ]
        if not cpu_rows:
            return
        preferred = [
            row for row in cpu_rows
            if str(row.get("Name") or "").strip().casefold() == "cpu package"
        ]
        if not preferred:
            preferred = [
                row for row in cpu_rows
                if str(row.get("Name") or "").strip().casefold() in {"core max", "cpu"}
            ]
        candidates = preferred or cpu_rows
        values = [value for row in candidates if (value := _lhm_sensor_value(row)) is not None]
        if values:
            snapshot.cpu.temperature = float(max(values))
            snapshot.cpu.temperature_source = "LibreHardwareMonitor REST"


# --------------------------------------------------------------------------- #
def _lhm_hardware_name(path: tuple[str, ...], kind: str) -> str:
    """Find the hardware node in an LHM sensor path."""
    candidates = [part for part in path if part]
    if kind == "gpu":
        tokens = ("radeon", "nvidia", "geforce", "gpu", "graphics")
        reject = ("gpu engine", "gpu adapter memory", "sensor")
    else:
        tokens = ("intel core", "13th gen intel", "cpu")
        reject = ("cpu package", "cpu core", "temperatures", "load", "clocks")
    for part in candidates:
        lowered = part.casefold()
        if any(token in lowered for token in tokens) and not any(token in lowered for token in reject):
            return part
    return ""


def _lhm_sensor_value(row: dict[str, Any]) -> float | None:
    """Prefer LHM's unformatted RawValue; fall back to its display Value."""
    raw = _float_or_none(row.get("RawValue"))
    if raw is not None:
        return raw
    return _float_from_display(row.get("Value"))


def _float_from_display(value: Any) -> float | None:
    if value is None:
        return None
    match = re.search(r"[-+]?\d+(?:[.,]\d+)?", str(value).replace(",", "."))
    return _float_or_none(match.group(0)) if match else None


def _pick_lhm_sensor(
    rows: list[dict[str, Any]], sensor_type: str, name_hints: tuple[str, ...]
) -> float | None:
    candidates = [
        row for row in rows
        if str(row.get("SensorType") or "").casefold() == sensor_type.casefold()
    ]
    for hint in name_hints:
        for row in candidates:
            if hint in str(row.get("Name") or "").casefold():
                value = _lhm_sensor_value(row)
                if value is not None:
                    return value
    return None


def _display_memory_mb(row: dict[str, Any]) -> float | None:
    """Convert LHM's human-readable memory value to MiB."""
    text = str(row.get("Value") or "")
    match = re.search(r"([-+]?\d+(?:[.,]\d+)?)\s*(B|KB|MB|GB|TB)\b", text, re.IGNORECASE)
    if match:
        number = float(match.group(1).replace(",", "."))
        unit = match.group(2).upper()
        factors = {"B": 1 / (1024 * 1024), "KB": 1 / 1024, "MB": 1, "GB": 1024, "TB": 1024 * 1024}
        return number * factors[unit]
    raw = _float_or_none(row.get("RawValue"))
    if raw is None:
        return None
    sensor_type = str(row.get("SensorType") or "").casefold()
    if sensor_type == "smalldata":
        return raw  # LHM SmallData raw values are already represented in MiB.
    if sensor_type == "data":
        return raw * 1024  # LHM Data sensors are commonly displayed in GB.
    return None


def _pick_lhm_memory(rows: list[dict[str, Any]], *, used: bool) -> float | None:
    candidates = [
        row for row in rows
        if str(row.get("SensorType") or "").casefold() in {"data", "smalldata"}
    ]
    if used:
        hints = ("memory used", "dedicated memory used", "memory allocated", "used memory")
    else:
        hints = ("memory total", "dedicated memory total", "total memory", "memory size")
    for hint in hints:
        for row in candidates:
            if hint in str(row.get("Name") or "").casefold():
                value = _display_memory_mb(row)
                if value is not None:
                    return value
    return None


def _float_or_none(value: Any) -> float | None:
    if value is None:
        return None
    text = str(value).strip().replace("[", "").replace("]", "")
    if text.lower() in {"n/a", "na", "none", ""}:
        return None
    try:
        result = float(text)
        return result if isfinite(result) else None
    except (TypeError, ValueError):
        return None


def _int_or_default(value: Any, fallback: int) -> int:
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return fallback


def _read_float_file(path: Path) -> float | None:
    try:
        return float(path.read_text(encoding="utf-8").strip())
    except (OSError, ValueError):
        return None


def _pick_sensor(
    rows: list[dict[str, Any]],
    sensor_type: str,
    name_hints: tuple[str, ...],
) -> float | None:
    candidates = [
        row
        for row in rows
        if str(row.get("SensorType")).lower() == sensor_type.lower()
    ]
    for hint in name_hints:
        for row in candidates:
            if hint in str(row.get("Name") or "").lower():
                value = _float_or_none(row.get("Value"))
                if value is not None:
                    return value
    return None
