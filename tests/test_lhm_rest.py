"""Focused tests for the local LibreHardwareMonitor REST sensor bridge."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from local_office_ai.ui.desktop import hardware  # noqa: E402
from local_office_ai.ui.desktop.hardware import HardwareMonitor, HardwareSnapshot  # noqa: E402


def _lhm_payload() -> dict:
    return {
        "Text": "Sensor",
        "Children": [
            {
                "Text": "SHAVIN",
                "Children": [
                    {
                        "Text": "13th Gen Intel Core i5-13400",
                        "HardwareId": "/intelcpu/0",
                        "Children": [
                            {
                                "Text": "Temperatures",
                                "Children": [
                                    {"Text": "CPU Package", "Type": "Temperature", "Value": "38.0 °C", "RawValue": 38.0, "Children": []},
                                    {"Text": "Core Max", "Type": "Temperature", "Value": "39.0 °C", "RawValue": 39.0, "Children": []},
                                ],
                            }
                        ],
                    },
                    {
                        "Text": "Radeon RX 580 Series",
                        "HardwareId": "/gpu-amd/0",
                        "Children": [
                            {
                                "Text": "Temperatures",
                                "Children": [
                                    {"Text": "GPU Core", "Type": "Temperature", "Value": "51.0 °C", "RawValue": 51.0, "Children": []}
                                ],
                            },
                            {
                                "Text": "Load",
                                "Children": [
                                    {"Text": "GPU Core", "Type": "Load", "Value": "12.0 %", "RawValue": 12.0, "Children": []}
                                ],
                            },
                            {
                                "Text": "Data",
                                "Children": [
                                    {"Text": "GPU Memory Used", "Type": "Data", "Value": "1.5 GB", "RawValue": 1.5, "Children": []},
                                    {"Text": "GPU Memory Total", "Type": "Data", "Value": "8.0 GB", "RawValue": 8.0, "Children": []},
                                ],
                            },
                        ],
                    },
                ],
            }
        ],
    }


def test_lhm_rest_tree_flattens_cpu_and_rx580_sensors():
    rows = HardwareMonitor._flatten_lhm_payload(_lhm_payload())
    assert any(row["Name"] == "CPU Package" and row["SensorType"] == "Temperature" for row in rows)
    assert any(row["Name"] == "GPU Core" and row["SensorType"] == "Load" for row in rows)
    assert any("Radeon RX 580 Series" in row["PathText"] for row in rows)


def test_lhm_rest_extracts_cpu_temperature_and_rx580_telemetry(monkeypatch):
    monkeypatch.setattr(hardware, "IS_WINDOWS", True)
    monitor = HardwareMonitor(enable_gpu=True)
    rows = HardwareMonitor._flatten_lhm_payload(_lhm_payload())
    monkeypatch.setattr(monitor, "_query_lhm_rest", lambda _url: rows)

    snapshot = HardwareSnapshot()
    monitor._read_cpu_temperature_from_lhm(snapshot)
    gpu_readings = monitor._lhm_gpus()

    assert snapshot.cpu.temperature == 38.0  # CPU Package is preferred to Core Max.
    assert snapshot.cpu.temperature_source == "LibreHardwareMonitor REST"
    assert len(gpu_readings) == 1
    gpu = gpu_readings[0]
    assert gpu.name == "Radeon RX 580 Series"
    assert gpu.temperature == 51.0
    assert gpu.utilization == 12.0
    assert gpu.vram_used_mb == 1536.0
    assert gpu.vram_total_mb == 8192.0


def test_lhm_rest_rejects_non_loopback_urls():
    assert HardwareMonitor._query_lhm_rest("http://192.168.1.10:8085/data.json") == []

