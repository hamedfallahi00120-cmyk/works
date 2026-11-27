# -*- coding: utf-8 -*-
"""Local Office Assistant V4.2: Qwen3.5 + Excel auditing + enhanced GPU telemetry + ComfyUI image generation.

Windows 10/11, Python 3.10+, Ollama. ComfyUI is an optional separate local service at 127.0.0.1:8188.
The source workbook is never overwritten. Excel output: *_AI_V4_2_Monitored.xlsx
Image generation is delegated to ComfyUI; Qwen only prepares the prompt.
"""

from __future__ import annotations

import json
import os
import re
import subprocess
import threading
import time
import traceback
import urllib.error
import urllib.request
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from tkinter import Tk, filedialog, messagebox

try:
    import psutil
    import openpyxl
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, LineChart, Reference
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError as exc:
    root = Tk()
    root.withdraw()
    messagebox.showerror(
        "کتابخانه موردنیاز نصب نیست",
        f"کتابخانه {exc.name or 'موردنیاز'} نصب نیست.\n\n"
        "در CMD اجرا کن:\npy -m pip install openpyxl psutil\n\nسپس برنامه را دوباره اجرا کن."
    )
    raise SystemExit(1)

MODEL = None  # selected interactively from locally installed Qwen3.5 models
OLLAMA_CHAT_URL = "http://127.0.0.1:11434/api/chat"
OLLAMA_TAGS_URL = "http://127.0.0.1:11434/api/tags"
SAMPLE_INTERVAL = 1.0
GPU_INTERVAL = 2.5
COMFY_DEFAULT_URL = "http://127.0.0.1:8188"

# Visual palette
NAVY = "17365D"
BLUE = "4472C4"
LIGHT_BLUE = "D9EAF7"
PALE_GREEN = "E2F0D9"
PALE_RED = "FCE4D6"
PALE_YELLOW = "FFF2CC"
PALE_GREY = "F2F2F2"
DARK = "1F1F1F"
WHITE = "FFFFFF"
THIN_GREY = Side(style="thin", color="D9E1F2")


def as_number(value):
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.replace(",", "").replace("٬", "").strip()
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def normalize_lender(name: str) -> str:
    """Normalize common Arabic/Persian character and spelling variants for matching only."""
    text = str(name or "").strip().replace("ي", "ی").replace("ك", "ک")
    text = text.replace("صندق", "صندوق")
    text = re.sub(r"\s+", " ", text)
    return text


def format_money(value):
    try:
        return f"{int(round(float(value))):,}"
    except (TypeError, ValueError):
        return str(value)


def find_declared_totals(ws, cached_ws=None):
    value_ws = cached_ws or ws
    declared_principal = None
    declared_installments = None
    for row in range(1, min(ws.max_row, 80) + 1):
        c_text = str(ws.cell(row, 3).value or "")
        e_text = str(ws.cell(row, 5).value or "")
        if "جمع کل وام" in c_text or "جمع کل وام" in c_text.replace(" ", ""):
            declared_principal = as_number(value_ws.cell(row, 4).value)
        if "جمع اقساط" in e_text:
            declared_installments = as_number(value_ws.cell(row, 6).value)
    return declared_principal, declared_installments


def extract_workbook(path: str):
    wb = openpyxl.load_workbook(path)
    if "Sheet1" not in wb.sheetnames:
        raise ValueError("برگه Sheet1 در فایل پیدا نشد.")
    ws = wb["Sheet1"]

    loans = []
    loan_row_numbers = []
    # Current source layout: C=Institution, D=Loan principal, E=Installment, F=number of installments.
    # Read all populated candidate rows rather than assuming there are exactly six loans.
    for r in range(3, ws.max_row + 1):
        lender = ws.cell(r, 3).value
        principal = as_number(ws.cell(r, 4).value)
        installment = as_number(ws.cell(r, 5).value)
        count = as_number(ws.cell(r, 6).value)
        if lender and principal is not None and installment is not None and count is not None:
            item = {
                "row": r,
                "institution_original": str(lender).strip(),
                "institution": normalize_lender(lender),
                "principal": principal,
                "installment": installment,
                "installment_count": count,
                "last_installment_date": ws.cell(r, 7).value,
            }
            loans.append(item)
            loan_row_numbers.append(r)

    if not loans:
        raise ValueError("هیچ ردیف وام قابل‌خواندنی پیدا نشد. ساختار فایل باید مشابه Book2.xlsx باشد.")

    # Commitments are separately listed below the loan table (institution in C, monthly amount in D).
    commitments = []
    last_loan_row = max(loan_row_numbers)
    for r in range(last_loan_row + 1, ws.max_row + 1):
        label = ws.cell(r, 3).value
        amount = as_number(ws.cell(r, 4).value)
        if label and amount is not None:
            text = str(label).strip()
            if "جمع" not in text:
                commitments.append({
                    "row": r,
                    "institution_original": text,
                    "institution": normalize_lender(text),
                    "monthly_amount": amount,
                })

    # Source totals are formulas; read their cached results from a data_only copy.
    cached_wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    cached_ws = cached_wb["Sheet1"]
    declared_principal, declared_installments = find_declared_totals(ws, cached_ws)
    cached_wb.close()

    # Deterministic calculations: the model does not perform arithmetic for official figures.
    by_lender = defaultdict(float)
    principal_by_lender = defaultdict(float)
    for loan in loans:
        by_lender[loan["institution"]] += loan["installment"]
        principal_by_lender[loan["institution"]] += loan["principal"]

    obligations_by_lender = defaultdict(float)
    for item in commitments:
        obligations_by_lender[item["institution"]] += item["monthly_amount"]

    computed_principal = sum(x["principal"] for x in loans)
    computed_installments = sum(x["installment"] for x in loans)
    total_monthly_obligations = sum(x["monthly_amount"] for x in commitments)
    total_contract_repayment = sum(x["installment"] * x["installment_count"] for x in loans)

    # Audit summary totals declared in the source against independent Python totals.
    audits = [
        {
            "title": "مجموع اصل وام‌ها",
            "declared": declared_principal,
            "computed": computed_principal,
            "difference": None if declared_principal is None else computed_principal - declared_principal,
            "status": "قابل بررسی نیست" if declared_principal is None else ("تأیید شد" if abs(computed_principal - declared_principal) < 0.01 else "مغایرت"),
        },
        {
            "title": "مجموع اقساط ماهانه وام‌ها",
            "declared": declared_installments,
            "computed": computed_installments,
            "difference": None if declared_installments is None else computed_installments - declared_installments,
            "status": "قابل بررسی نیست" if declared_installments is None else ("تأیید شد" if abs(computed_installments - declared_installments) < 0.01 else "مغایرت"),
        },
    ]

    reconciliation = []
    for lender in sorted(set(by_lender) | set(obligations_by_lender)):
        loan_value = by_lender.get(lender, 0.0)
        obligation_value = obligations_by_lender.get(lender, 0.0)
        if loan_value and obligation_value:
            diff = obligation_value - loan_value
            status = "مطابقت" if abs(diff) < 0.01 else "مغایرت نیازمند بررسی"
        elif obligation_value and not loan_value:
            diff = None
            status = "تعهد ثبت‌شده خارج از جدول وام‌ها"
        else:
            diff = None
            status = "برای وام، ردیف تعهد ماهانه پیدا نشد"
        reconciliation.append({
            "institution": lender,
            "loan_installments": loan_value,
            "monthly_obligation": obligation_value,
            "difference": diff,
            "status": status,
        })

    # Detect spelling variants for transparent normalization; source values are not overwritten.
    spelling_variants = []
    originals_by_normalized = defaultdict(set)
    for x in loans:
        originals_by_normalized[x["institution"]].add(x["institution_original"])
    for x in commitments:
        originals_by_normalized[x["institution"]].add(x["institution_original"])
    for canonical, originals in originals_by_normalized.items():
        if len(originals) > 1:
            spelling_variants.append(f"برای تطبیق، شکل‌های نوشتاری { ' / '.join(sorted(originals)) } یکسان‌سازی شدند؛ متن منبع تغییر نکرد.")

    return wb, ws, {
        "loans": loans,
        "commitments": commitments,
        "by_lender": dict(by_lender),
        "principal_by_lender": dict(principal_by_lender),
        "obligations_by_lender": dict(obligations_by_lender),
        "computed_principal": computed_principal,
        "computed_installments": computed_installments,
        "total_monthly_obligations": total_monthly_obligations,
        "total_contract_repayment": total_contract_repayment,
        "declared_principal": declared_principal,
        "declared_installments": declared_installments,
        "audits": audits,
        "reconciliation": reconciliation,
        "spelling_variants": spelling_variants,
    }


class ResourceMonitor:
    """Windows telemetry for Ollama and ComfyUI; GPU counters are queried in one PowerShell call."""
    PROCESS_EXACT = {
        "ollama.exe", "ollama app.exe", "ollama_llama_server.exe",
        "llama-server.exe", "llama_server.exe", "llama-runner.exe",
        "llama_runner.exe", "ollama-runner.exe", "ollama_runner.exe",
    }

    def __init__(self, interval=1.0, gpu_interval=2.5):
        self.interval = interval
        self.gpu_interval = gpu_interval
        self.samples = []
        self.gpu_samples = []
        self.stop_event = threading.Event()
        self.thread = None
        self.gpu_thread = None
        self.start_perf = None
        self.start_wall = None
        self.phase = "شروع"
        self.phase_lock = threading.Lock()
        self.detected_processes = {}
        self.detected_roles = {}
        self._previous_cpu = {}
        self._previous_sample = None
        self._cpu_total_seconds = defaultdict(float)
        self._io_last = {}
        self._last_gpu_start = None
        self._gpu_counter_status = {"engine": False, "dedicated": False, "shared": False}
        self._self_pid = os.getpid()

    def set_phase(self, phase):
        with self.phase_lock:
            self.phase = phase

    @staticmethod
    def _is_ollama_name(name):
        name = (name or "").strip().lower()
        return name in ResourceMonitor.PROCESS_EXACT or name.startswith("ollama_") or name.startswith("llama-") or name.startswith("llama_")

    def _find_tracked_processes(self):
        tracked = []
        for proc in psutil.process_iter(["pid", "name", "cmdline"]):
            try:
                pid = int(proc.info["pid"])
                if pid == self._self_pid:
                    continue
                name = (proc.info.get("name") or "").strip()
                if self._is_ollama_name(name):
                    role = "ollama"
                else:
                    lname = name.lower()
                    if lname not in {"python.exe", "pythonw.exe", "comfyui.exe"}:
                        continue
                    cmd = " ".join(proc.info.get("cmdline") or []).lower()
                    try:
                        cwd = proc.cwd().lower()
                    except (psutil.AccessDenied, psutil.NoSuchProcess, OSError):
                        cwd = ""
                    is_comfy = ("main.py" in cmd and ("comfyui" in cmd or "comfyui" in cwd)) or "comfyui.exe" == lname
                    if not is_comfy:
                        continue
                    role = "comfyui"
                self.detected_processes[pid] = name
                self.detected_roles[pid] = role
                tracked.append((psutil.Process(pid), role))
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess, TypeError):
                continue
        return tracked

    @staticmethod
    def _adapter_parts(instance):
        text = str(instance or "")
        luid_match = re.search(r"(luid_0x[0-9a-f]+_0x[0-9a-f]+)", text, re.IGNORECASE)
        phys_match = re.search(r"phys_(\d+)", text, re.IGNORECASE)
        luid = luid_match.group(1) if luid_match else "LUID_UNKNOWN"
        phys = int(phys_match.group(1)) if phys_match else None
        key = f"{luid}_phys_{phys}" if phys is not None else luid
        return key, luid, phys

    @staticmethod
    def _counter_pid(instance):
        match = re.search(r"pid_(\d+)_", str(instance or ""), re.IGNORECASE)
        return int(match.group(1)) if match else None

    def _collect_system_sample(self):
        now_perf = time.perf_counter()
        now_wall = time.time()
        vm = psutil.virtual_memory()
        system_cpu = psutil.cpu_percent(interval=None)
        logical = max(psutil.cpu_count(logical=True) or 1, 1)
        role_data = {
            "ollama": {"cpu_delta": 0.0, "rss": 0, "vms": 0, "read": 0, "write": 0, "pids": []},
            "comfyui": {"cpu_delta": 0.0, "rss": 0, "vms": 0, "read": 0, "write": 0, "pids": []},
        }
        process_rows = []
        for proc, role in self._find_tracked_processes():
            try:
                pid = proc.pid
                cpu = proc.cpu_times()
                cpu_total = float(cpu.user + cpu.system)
                prev = self._previous_cpu.get(pid)
                delta = max(0.0, cpu_total - prev) if prev is not None else 0.0
                self._previous_cpu[pid] = cpu_total
                mem = proc.memory_info()
                entry = role_data[role]
                entry["cpu_delta"] += delta
                entry["rss"] += mem.rss
                entry["vms"] += mem.vms
                entry["pids"].append(pid)
                try:
                    io = proc.io_counters()
                    entry["read"] += io.read_bytes
                    entry["write"] += io.write_bytes
                except (psutil.AccessDenied, AttributeError, OSError):
                    pass
                process_rows.append({
                    "pid": pid, "name": self.detected_processes.get(pid, ""), "role": role,
                    "rss_mb": mem.rss / (1024 ** 2), "cpu_seconds": cpu_total
                })
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue

        elapsed = (now_perf - self._previous_sample) if self._previous_sample is not None else 0.0
        self._previous_sample = now_perf
        cpu_vals = {}
        for role, entry in role_data.items():
            raw_pct = (entry["cpu_delta"] / elapsed * 100.0) if elapsed > 0 else 0.0
            cpu_vals[role] = {"raw": raw_pct, "machine": raw_pct / logical}
            self._cpu_total_seconds[role] += entry["cpu_delta"]
        try:
            python_rss = psutil.Process(self._self_pid).memory_info().rss
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            python_rss = 0
        with self.phase_lock:
            phase = self.phase
        self.samples.append({
            "timestamp": now_wall, "elapsed_sec": now_perf - self.start_perf, "phase": phase,
            "system_cpu_pct": system_cpu,
            "system_ram_used_gb": (vm.total - vm.available) / (1024 ** 3),
            "system_ram_used_pct": vm.percent,
            "system_ram_available_gb": vm.available / (1024 ** 3),
            "system_ram_total_gb": vm.total / (1024 ** 3),
            "python_rss_mb": python_rss / (1024 ** 2),
            "ollama_cpu_machine_pct": cpu_vals["ollama"]["machine"],
            "ollama_cpu_raw_pct": cpu_vals["ollama"]["raw"],
            "ollama_cpu_total_seconds": self._cpu_total_seconds["ollama"],
            "ollama_rss_mb": role_data["ollama"]["rss"] / (1024 ** 2),
            "ollama_vms_mb": role_data["ollama"]["vms"] / (1024 ** 2),
            "ollama_read_bytes_cumulative": role_data["ollama"]["read"],
            "ollama_write_bytes_cumulative": role_data["ollama"]["write"],
            "ollama_pids": ",".join(str(x) for x in sorted(role_data["ollama"]["pids"])),
            "comfyui_cpu_machine_pct": cpu_vals["comfyui"]["machine"],
            "comfyui_cpu_raw_pct": cpu_vals["comfyui"]["raw"],
            "comfyui_cpu_total_seconds": self._cpu_total_seconds["comfyui"],
            "comfyui_rss_mb": role_data["comfyui"]["rss"] / (1024 ** 2),
            "comfyui_pids": ",".join(str(x) for x in sorted(role_data["comfyui"]["pids"])),
            "process_rows": process_rows,
        })

    def _sample_system(self):
        psutil.cpu_percent(interval=None)
        while not self.stop_event.is_set():
            try:
                self._collect_system_sample()
            except Exception as exc:
                self.samples.append({"timestamp": time.time(), "elapsed_sec": time.perf_counter() - self.start_perf,
                                     "phase": self.phase, "sample_error": str(exc)[:250]})
            self.stop_event.wait(self.interval)

    def _sample_gpu(self):
        # One Get-Counter call for all three counter families. This reduces query overhead
        # and avoids the 7-8 second effective cadence observed when each family was queried separately.
        ps_script = r"""
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
    $out += [pscustomobject]@{Kind=$kind; InstanceName=$s.InstanceName; Path=$s.Path; Value=[double]$s.CookedValue}
  }
  if ($out.Count -eq 0) { '[]' } else { ConvertTo-Json -InputObject $out -Compress -Depth 4 }
} catch {
  [pscustomobject]@{__error=$_.Exception.Message} | ConvertTo-Json -Compress
}
"""
        next_start = time.perf_counter()
        while not self.stop_event.is_set():
            tick_start = time.perf_counter()
            live_roles = {}
            for proc, role in self._find_tracked_processes():
                live_roles[proc.pid] = role
            try:
                cp = subprocess.run(
                    ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", ps_script],
                    capture_output=True, text=True, timeout=12,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)
                )
                query_elapsed = time.perf_counter() - tick_start
                raw = (cp.stdout or "").strip()
                if not raw:
                    raise RuntimeError((cp.stderr or "شمارنده‌های GPU خروجی ندادند")[:250])
                payload = json.loads(raw)
                if isinstance(payload, dict) and payload.get("__error"):
                    raise RuntimeError(payload["__error"])
                items = payload if isinstance(payload, list) else [payload]
                now = time.time()
                actual_interval = (tick_start - self._last_gpu_start) if self._last_gpu_start is not None else None
                self._last_gpu_start = tick_start
                adapters = {}
                all_engine, oll_engine, comfy_engine = [], [], []
                total_oll_ded = total_oll_shared = total_comfy_ded = total_comfy_shared = 0.0
                all_ded = all_shared = 0.0
                for item in items:
                    kind = str(item.get("Kind", ""))
                    instance = str(item.get("InstanceName") or "")
                    val = max(0.0, float(item.get("Value") or 0.0))
                    adapter_key, luid, phys = self._adapter_parts(instance)
                    row = adapters.setdefault(adapter_key, {
                        "adapter_key": adapter_key, "luid": luid, "physical_index": phys,
                        "engine_values_all": [], "engine_values_ollama": [], "engine_values_comfyui": [],
                        "engine_types": set(), "ollama_dedicated_mb": 0.0, "ollama_shared_mb": 0.0,
                        "comfyui_dedicated_mb": 0.0, "comfyui_shared_mb": 0.0,
                        "all_process_dedicated_mb": 0.0, "all_process_shared_mb": 0.0,
                        "ollama_pids": set(), "comfyui_pids": set()
                    })
                    pid = self._counter_pid(instance)
                    role = live_roles.get(pid) if pid is not None else None
                    engine_type_match = re.search(r"engtype_([^_]+)", instance, re.IGNORECASE)
                    if engine_type_match:
                        row["engine_types"].add(engine_type_match.group(1))
                    if kind == "engine":
                        row["engine_values_all"].append(val)
                        all_engine.append(val)
                        if role == "ollama":
                            row["engine_values_ollama"].append(val)
                            row["ollama_pids"].add(pid)
                            oll_engine.append(val)
                        elif role == "comfyui":
                            row["engine_values_comfyui"].append(val)
                            row["comfyui_pids"].add(pid)
                            comfy_engine.append(val)
                    elif kind == "dedicated":
                        row["all_process_dedicated_mb"] += val / (1024 ** 2)
                        all_ded += val
                        if role == "ollama":
                            row["ollama_dedicated_mb"] += val / (1024 ** 2)
                            total_oll_ded += val
                            row["ollama_pids"].add(pid)
                        elif role == "comfyui":
                            row["comfyui_dedicated_mb"] += val / (1024 ** 2)
                            total_comfy_ded += val
                            row["comfyui_pids"].add(pid)
                    elif kind == "shared":
                        row["all_process_shared_mb"] += val / (1024 ** 2)
                        all_shared += val
                        if role == "ollama":
                            row["ollama_shared_mb"] += val / (1024 ** 2)
                            total_oll_shared += val
                            row["ollama_pids"].add(pid)
                        elif role == "comfyui":
                            row["comfyui_shared_mb"] += val / (1024 ** 2)
                            total_comfy_shared += val
                            row["comfyui_pids"].add(pid)

                adapter_rows = []
                for row in adapters.values():
                    adapter_rows.append({
                        "adapter_key": row["adapter_key"], "luid": row["luid"], "physical_index": row["physical_index"],
                        "ollama_engine_max_pct": max(row["engine_values_ollama"]) if row["engine_values_ollama"] else None,
                        "ollama_engine_sum_pct": sum(row["engine_values_ollama"]) if row["engine_values_ollama"] else None,
                        "comfyui_engine_max_pct": max(row["engine_values_comfyui"]) if row["engine_values_comfyui"] else None,
                        "comfyui_engine_sum_pct": sum(row["engine_values_comfyui"]) if row["engine_values_comfyui"] else None,
                        "all_engine_max_pct": max(row["engine_values_all"]) if row["engine_values_all"] else None,
                        "engine_count": len(row["engine_values_all"]),
                        "engine_types": ",".join(sorted(row["engine_types"])),
                        "ollama_dedicated_mb": row["ollama_dedicated_mb"], "ollama_shared_mb": row["ollama_shared_mb"],
                        "comfyui_dedicated_mb": row["comfyui_dedicated_mb"], "comfyui_shared_mb": row["comfyui_shared_mb"],
                        "all_process_dedicated_mb": row["all_process_dedicated_mb"],
                        "all_process_shared_mb": row["all_process_shared_mb"],
                        "ollama_pids": ",".join(str(x) for x in sorted(row["ollama_pids"])),
                        "comfyui_pids": ",".join(str(x) for x in sorted(row["comfyui_pids"]))
                    })
                with self.phase_lock:
                    phase = self.phase
                self._gpu_counter_status["engine"] = self._gpu_counter_status["engine"] or bool(all_engine)
                self._gpu_counter_status["dedicated"] = self._gpu_counter_status["dedicated"] or all_ded > 0
                self._gpu_counter_status["shared"] = self._gpu_counter_status["shared"] or all_shared > 0
                self.gpu_samples.append({
                    "timestamp": now, "phase": phase, "available": bool(items),
                    "query_elapsed_sec": query_elapsed, "sample_interval_sec": actual_interval,
                    "ollama_engine_max_pct": max(oll_engine) if oll_engine else None,
                    "ollama_engine_sum_pct": sum(oll_engine) if oll_engine else None,
                    "comfyui_engine_max_pct": max(comfy_engine) if comfy_engine else None,
                    "comfyui_engine_sum_pct": sum(comfy_engine) if comfy_engine else None,
                    "all_engine_max_pct": max(all_engine) if all_engine else None,
                    "ollama_dedicated_mb": total_oll_ded / (1024 ** 2) if total_oll_ded else None,
                    "ollama_shared_mb": total_oll_shared / (1024 ** 2) if total_oll_shared else None,
                    "comfyui_dedicated_mb": total_comfy_ded / (1024 ** 2) if total_comfy_ded else None,
                    "comfyui_shared_mb": total_comfy_shared / (1024 ** 2) if total_comfy_shared else None,
                    "all_process_dedicated_mb": all_ded / (1024 ** 2) if all_ded else None,
                    "all_process_shared_mb": all_shared / (1024 ** 2) if all_shared else None,
                    "ollama_engine_count": len(oll_engine), "comfyui_engine_count": len(comfy_engine),
                    "ollama_pids": ",".join(str(x) for x, role in live_roles.items() if role == "ollama"),
                    "comfyui_pids": ",".join(str(x) for x, role in live_roles.items() if role == "comfyui"),
                    "adapter_rows": adapter_rows,
                })
            except Exception as exc:
                self.gpu_samples.append({
                    "timestamp": time.time(), "phase": self.phase, "available": False,
                    "query_elapsed_sec": time.perf_counter() - tick_start,
                    "sample_interval_sec": (tick_start - self._last_gpu_start) if self._last_gpu_start else None,
                    "error": str(exc)[:300]
                })
                self._last_gpu_start = tick_start
            # Fixed start-to-start schedule; query duration is included, not added to the sleep interval.
            remaining = self.gpu_interval - (time.perf_counter() - tick_start)
            if self.stop_event.wait(max(0.0, remaining)):
                break

    def start(self):
        self.start_perf = time.perf_counter()
        self.start_wall = datetime.now()
        self.thread = threading.Thread(target=self._sample_system, daemon=True)
        self.gpu_thread = threading.Thread(target=self._sample_gpu, daemon=True)
        self.thread.start()
        self.gpu_thread.start()

    def stop(self):
        self.stop_event.set()
        if self.thread:
            self.thread.join(timeout=4)
        if self.gpu_thread:
            self.gpu_thread.join(timeout=14)
        return time.perf_counter() - self.start_perf


def get_installed_qwen_models():
    req = urllib.request.Request(OLLAMA_TAGS_URL, method="GET")
    with urllib.request.urlopen(req, timeout=8) as response:
        payload = json.loads(response.read().decode("utf-8"))
    names = sorted({m.get("name", "") for m in payload.get("models", []) if m.get("name", "").startswith("qwen3.5:")})
    return names


def choose_model_dialog(root, models):
    """Let the operator select an installed Qwen3.5 model for a fair same-task comparison."""
    import tkinter as tk
    from tkinter import ttk
    selected = {"value": None}
    win = tk.Toplevel(root)
    win.title("انتخاب مدل Qwen3.5 برای آزمایش")
    win.geometry("470x190")
    win.resizable(False, False)
    win.grab_set()
    tk.Label(win, text="مدل محلی موردنظر را برای اجرای مأموریت انتخاب کن:",
             font=("Tahoma", 10, "bold")).pack(padx=16, pady=(18, 10), anchor="e")
    default = next((m for m in models if m.endswith(":9b")), models[0])
    variable = tk.StringVar(value=default)
    combo = ttk.Combobox(win, textvariable=variable, values=models, state="readonly", width=42)
    combo.pack(padx=16, pady=8)
    tk.Label(win, text="مدل‌های نصب‌شده Qwen3.5 نمایش داده می‌شوند؛ برای مقایسه منصفانه فایل و Context ثابت می‌ماند.",
             font=("Tahoma", 8)).pack(padx=12, pady=4)
    buttons = tk.Frame(win)
    buttons.pack(pady=12)
    def accept():
        selected["value"] = variable.get().strip()
        win.destroy()
    def cancel():
        win.destroy()
    ttk.Button(buttons, text="شروع آزمایش", command=accept).pack(side="left", padx=6)
    ttk.Button(buttons, text="لغو", command=cancel).pack(side="left", padx=6)
    win.protocol("WM_DELETE_WINDOW", cancel)
    win.wait_window()
    return selected["value"]


def ollama_ready():
    req = urllib.request.Request(OLLAMA_TAGS_URL, method="GET")
    with urllib.request.urlopen(req, timeout=8) as response:
        payload = json.loads(response.read().decode("utf-8"))
    names = [m.get("name", "") for m in payload.get("models", [])]
    if not any(n == MODEL or n.startswith(MODEL + ":") for n in names):
        raise RuntimeError(f"مدل {MODEL} در Ollama پیدا نشد. خروجی ollama list را بررسی کن.")


def ask_model(facts):
    model_input = {
        "verified_calculations_from_python": {
            "loan_count": len(facts["loans"]),
            "principal_sum": facts["computed_principal"],
            "monthly_installments_sum": facts["computed_installments"],
            "monthly_obligations_sum": facts["total_monthly_obligations"],
            "contract_repayment_sum": facts["total_contract_repayment"],
            "institution_reconciliation": facts["reconciliation"],
            "source_declared_totals_audit": facts["audits"],
        },
        "loan_profile_by_institution": [
            {
                "institution": lender,
                "loan_count": sum(1 for x in facts["loans"] if x["institution"] == lender),
                "last_installment_dates": [
                    str(x["last_installment_date"] or "")
                    for x in facts["loans"] if x["institution"] == lender
                ],
            }
            for lender in sorted({x["institution"] for x in facts["loans"]})
        ],
        "monthly_commitments": facts["commitments"],
        "task": [
            "تحلیل مدیریتی فارسی و کوتاه ارائه کن.",
            "هیچ جمع، ضرب، تفریق یا محاسبه عددی انجام نده و هیچ مبلغ جدیدی تولید نکن.",
            "تنها اعداد مجاز برای نقل‌قول، اعداد موجود در verified_calculations_from_python هستند.",
            "مغایرتی را ادعا نکن مگر اینکه دقیقاً در حسابرسی پایتون به‌عنوان مغایرت ثبت شده باشد.",
            "برای توضیح تطبیق مؤسسات فقط از خروجی institution_reconciliation استفاده کن؛ داده خام ردیف‌به‌ردیف مالی در اختیار تو نیست.",
            "تفاوت بین اقساط وام‌های اصلی و تعهدهای ماهانه خارج از جدول وام‌ها را توضیح بده.",
            "از تاریخ آخرین قسط، تعداد اقساط پرداخت‌شده یا مانده بدهی را حدس نزن.",
            "اگر اطلاعاتی برای تحلیل موجود نیست، آن را به عنوان داده ناقص مشخص کن."
        ]
    }
    body = {
        "model": MODEL,
        "stream": False,
        "think": False,
        "format": "json",
        "options": {"temperature": 0.1, "num_ctx": 4096},
        "messages": [
            {
                "role": "system",
                "content": (
                    "تو دستیار تحلیل‌گر اداری هستی. داده‌ها و محاسبات تأییدشده را تحلیل کن. "
                    "محاسبات مجدد انجام نده و عدد تازه نساز. پاسخ فقط JSON معتبر با این کلیدها باشد: "
                    "summary (string), findings (array of strings), missing_data (array of strings), "
                    "recommendations (array of strings). متن فارسی باشد."
                )
            },
            {"role": "user", "content": json.dumps(model_input, ensure_ascii=False)}
        ]
    }
    req = urllib.request.Request(
        OLLAMA_CHAT_URL,
        data=json.dumps(body, ensure_ascii=False).encode("utf-8"),
        headers={"Content-Type": "application/json"}, method="POST"
    )
    call_start = time.perf_counter()
    with urllib.request.urlopen(req, timeout=900) as response:
        result = json.loads(response.read().decode("utf-8"))
    client_elapsed = time.perf_counter() - call_start
    content = result.get("message", {}).get("content", "")
    try:
        parsed = json.loads(content)
    except json.JSONDecodeError:
        parsed = {"summary": content[:3500], "findings": [], "missing_data": ["پاسخ مدل JSON معتبر نبود."], "recommendations": []}
    for key in ("findings", "missing_data", "recommendations"):
        if isinstance(parsed.get(key), str):
            parsed[key] = [parsed[key]]
        if not isinstance(parsed.get(key), list):
            parsed[key] = []
    if not isinstance(parsed.get("summary"), str):
        parsed["summary"] = str(parsed.get("summary", ""))
    return parsed, result, client_elapsed


def ollama_ps_snapshot():
    try:
        cp = subprocess.run(["ollama", "ps"], capture_output=True, text=True, timeout=10,
                            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        return (cp.stdout or cp.stderr or "").strip()
    except Exception as exc:
        return f"اجرای ollama ps ناموفق بود: {exc}"


def set_title(ws, cell_range, title):
    ws.merge_cells(cell_range)
    cell = ws[cell_range.split(":")[0]]
    cell.value = title
    cell.fill = PatternFill("solid", fgColor=NAVY)
    cell.font = Font(name="Tahoma", size=15, bold=True, color=WHITE)
    cell.alignment = Alignment(horizontal="center", vertical="center")


def style_header(ws, row, start_col, end_col):
    for col in range(start_col, end_col + 1):
        c = ws.cell(row, col)
        c.fill = PatternFill("solid", fgColor=BLUE)
        c.font = Font(name="Tahoma", bold=True, color=WHITE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(bottom=THIN_GREY)


def style_body(ws, min_row, max_row, min_col, max_col):
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for c in row:
            c.font = Font(name="Tahoma", size=10, color=DARK, bold=c.font.bold)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = Border(bottom=THIN_GREY)


def add_excel_output(wb, source_ws, facts, ai):
    # Preserve source entries; append a calculated repayment column only.
    source_ws.sheet_view.rightToLeft = True
    repay_col = max(source_ws.max_column if source_ws.max_column < 12 else 8, 8)
    # This source file has unused columns beyond G; always use H for the new computed field.
    repay_col = 8
    source_ws.cell(2, repay_col, "مجموع بازپرداخت قراردادی")
    for loan in facts["loans"]:
        r = loan["row"]
        source_ws.cell(r, repay_col, f"=E{r}*F{r}")
        source_ws.cell(r, repay_col).number_format = '#,##0'
    source_ws.cell(2, repay_col).fill = PatternFill("solid", fgColor=NAVY)
    source_ws.cell(2, repay_col).font = Font(name="Tahoma", bold=True, color=WHITE)
    source_ws.cell(2, repay_col).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    source_ws.column_dimensions[get_column_letter(repay_col)].width = 25
    source_ws.freeze_panes = "C3"
    source_ws.auto_filter.ref = f"B2:H{max(x['row'] for x in facts['loans'])}"
    for row in source_ws.iter_rows(min_row=2, max_row=max(x['row'] for x in facts['loans']), min_col=2, max_col=8):
        for c in row:
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    if "داشبورد" in wb.sheetnames:
        del wb["داشبورد"]
    dash = wb.create_sheet("داشبورد", 0)
    dash.sheet_view.rightToLeft = True
    set_title(dash, "A1:F1", "داشبورد مدیریت وام‌ها | نسخه ۴.۲ با کنترل صحت محاسبات و مانیتورینگ تفکیکی GPU")
    dash.row_dimensions[1].height = 34

    dash.append(["شاخص کلیدی", "مقدار", "مبنای محاسبه", "وضعیت", "", ""])
    style_header(dash, 2, 1, 6)
    kpis = [
        ["تعداد وام‌های ثبت‌شده", len(facts["loans"]), "استخراج مستقیم از ردیف‌های وام", "تأیید شده"],
        ["مجموع اصل وام‌ها", facts["computed_principal"], "محاسبه مستقل با Python", audits_status(facts, 0)],
        ["مجموع اقساط ماهانه وام‌ها", facts["computed_installments"], "محاسبه مستقل با Python", audits_status(facts, 1)],
        ["تعهدهای ماهانه ثبت‌شده در پایین جدول", facts["total_monthly_obligations"], "جمع ردیف‌های تعهدات", "اطلاعات منبع"],
        ["کل بازپرداخت قراردادی وام‌ها", facts["total_contract_repayment"], "مبلغ هر قسط × تعداد کل اقساط", "محاسبه Python"],
    ]
    for row in kpis:
        dash.append(row)
    for r in range(3, 8):
        dash.cell(r, 2).number_format = '#,##0'
    style_body(dash, 3, 7, 1, 4)

    row = 10
    set_title(dash, f"A{row}:F{row}", "حسابرسی اعداد ثبت‌شده در فایل اصلی")
    row += 1
    dash.append(["شاخص", "عدد اعلام‌شده در فایل", "محاسبه مستقل Python", "اختلاف", "نتیجه", ""])
    style_header(dash, row, 1, 6)
    for audit in facts["audits"]:
        row += 1
        dash.cell(row, 1, audit["title"])
        dash.cell(row, 2, audit["declared"])
        dash.cell(row, 3, audit["computed"])
        dash.cell(row, 4, audit["difference"])
        dash.cell(row, 5, audit["status"])
        for col in (2, 3, 4):
            dash.cell(row, col).number_format = '#,##0'
        dash.cell(row, 5).fill = PatternFill("solid", fgColor=PALE_GREEN if audit["status"] == "تأیید شد" else PALE_RED)
    style_body(dash, 12, row, 1, 5)

    row += 2
    set_title(dash, f"A{row}:F{row}", "تطبیق اقساط وام‌ها با تعهدهای ماهانه ثبت‌شده")
    row += 1
    dash.append(["مؤسسه", "اقساط از جدول وام‌ها", "تعهد ماهانه ثبت‌شده", "اختلاف", "نتیجه تطبیق", ""])
    style_header(dash, row, 1, 6)
    rec_start = row + 1
    for item in facts["reconciliation"]:
        row += 1
        values = [item["institution"], item["loan_installments"], item["monthly_obligation"], item["difference"], item["status"], ""]
        for col, val in enumerate(values, start=1):
            dash.cell(row, col, val)
        for col in (2, 3, 4):
            dash.cell(row, col).number_format = '#,##0'
        dash.cell(row, 5).fill = PatternFill("solid", fgColor=PALE_GREEN if item["status"] == "مطابقت" else PALE_YELLOW)
    rec_end = row
    style_body(dash, rec_start, rec_end, 1, 5)

    if facts["spelling_variants"]:
        row += 2
        set_title(dash, f"A{row}:F{row}", "یکسان‌سازی نوشتار برای مقایسه")
        for note in facts["spelling_variants"]:
            row += 1
            dash.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
            dash.cell(row, 1, note)
            dash.cell(row, 1).alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
            dash.row_dimensions[row].height = 30

    row += 2
    set_title(dash, f"A{row}:F{row}", f"تحلیل کیفی {MODEL} | یافته‌های مدل مستقل از حسابرسی عددی هستند")
    ai_sections = [
        ("خلاصه مدیریتی", ai.get("summary", "")),
        ("یافته‌های مدل", "\n".join("• " + str(x) for x in ai.get("findings", []))),
        ("اطلاعات ناقص", "\n".join("• " + str(x) for x in ai.get("missing_data", []))),
        ("پیشنهادهای مدل", "\n".join("• " + str(x) for x in ai.get("recommendations", []))),
    ]
    for label, text in ai_sections:
        row += 1
        dash.cell(row, 1, label)
        dash.cell(row, 1).font = Font(name="Tahoma", bold=True)
        dash.merge_cells(start_row=row, start_column=2, end_row=row, end_column=6)
        dash.cell(row, 2, text or "موردی گزارش نشد.")
        dash.cell(row, 2).alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)
        dash.row_dimensions[row].height = max(30, min(100, 18 * (str(text).count("\n") + 2)))
    row += 2
    # Chart from commitments: show separate monthly obligations as entered in source.
    chart_start = row
    dash.cell(chart_start, 1, "مؤسسه")
    dash.cell(chart_start, 2, "تعهد ماهانه")
    for idx, item in enumerate(facts["commitments"], start=chart_start + 1):
        dash.cell(idx, 1, item["institution"])
        dash.cell(idx, 2, item["monthly_amount"])
        dash.cell(idx, 2).number_format = '#,##0'
    if facts["commitments"]:
        chart = BarChart()
        chart.type = "bar"
        chart.style = 10
        chart.title = "تعهد ماهانه به تفکیک مؤسسه"
        chart.x_axis.title = "تومان"
        chart.y_axis.title = "مؤسسه"
        data = Reference(dash, min_col=2, min_row=chart_start, max_row=chart_start + len(facts["commitments"]))
        cats = Reference(dash, min_col=1, min_row=chart_start + 1, max_row=chart_start + len(facts["commitments"]))
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.height = 7
        chart.width = 13
        dash.add_chart(chart, "H2")
    dash.column_dimensions["A"].width = 38
    dash.column_dimensions["B"].width = 24
    dash.column_dimensions["C"].width = 28
    dash.column_dimensions["D"].width = 20
    dash.column_dimensions["E"].width = 34
    dash.column_dimensions["F"].width = 18
    dash.freeze_panes = "A3"


def audits_status(facts, index):
    return facts["audits"][index]["status"]


def metric_stats(samples, key):
    vals = [float(x[key]) for x in samples if x.get(key) is not None and isinstance(x.get(key), (int, float))]
    if not vals:
        return None, None, None, None
    vals_sorted = sorted(vals)
    p95 = vals_sorted[min(len(vals_sorted) - 1, max(0, int(round(0.95 * (len(vals_sorted) - 1)))))]
    return sum(vals) / len(vals), max(vals), min(vals), p95


def median_gpu_interval(gpu_samples):
    vals = [float(x["sample_interval_sec"]) for x in gpu_samples
            if isinstance(x.get("sample_interval_sec"), (int, float)) and x["sample_interval_sec"] > 0]
    if not vals:
        return None
    vals.sort()
    n = len(vals)
    return vals[n // 2] if n % 2 else (vals[n // 2 - 1] + vals[n // 2]) / 2


def hardware_snapshot():
    info = {
        "CPU logical cores": psutil.cpu_count(logical=True),
        "CPU physical cores": psutil.cpu_count(logical=False),
        "RAM total GB": round(psutil.virtual_memory().total / (1024 ** 3), 2),
        "CPU model": "",
        "GPU names": "",
        "Windows version": os.environ.get("OS", "Windows"),
        "Ollama version": "",
    }
    try:
        ps = r"Get-CimInstance Win32_Processor | Select-Object -ExpandProperty Name | ConvertTo-Json -Compress"
        cp = subprocess.run(["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", ps],
                            capture_output=True, text=True, timeout=10,
                            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        if cp.stdout.strip():
            names = json.loads(cp.stdout.strip())
            info["CPU model"] = ", ".join(names if isinstance(names, list) else [str(names)])
    except Exception:
        info["CPU model"] = "از WMI قابل خواندن نبود"
    try:
        ps = r"Get-CimInstance Win32_VideoController | Select-Object -ExpandProperty Name | ConvertTo-Json -Compress"
        cp = subprocess.run(["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", ps],
                            capture_output=True, text=True, timeout=10,
                            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        if cp.stdout.strip():
            names = json.loads(cp.stdout.strip())
            info["GPU names"] = ", ".join(names if isinstance(names, list) else [str(names)])
    except Exception:
        info["GPU names"] = "از WMI قابل خواندن نبود"
    try:
        ps = r"Get-CimInstance Win32_VideoController | Select-Object Name,PNPDeviceID,DriverVersion,VideoProcessor | ConvertTo-Json -Compress -Depth 3"
        cp = subprocess.run(["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", ps],
                            capture_output=True, text=True, timeout=10,
                            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        if cp.stdout.strip():
            rows = json.loads(cp.stdout.strip())
            if isinstance(rows, dict): rows = [rows]
            info["GPU adapter inventory"] = " | ".join(
                f"{i}: {x.get('Name','')} | Driver {x.get('DriverVersion','')} | {x.get('VideoProcessor','')} | PNP {x.get('PNPDeviceID','')}"
                for i, x in enumerate(rows)
            )
    except Exception:
        info["GPU adapter inventory"] = "جزئیات WMI قابل خواندن نبود"
    try:
        cp = subprocess.run(["ollama", "--version"], capture_output=True, text=True, timeout=8,
                            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        info["Ollama version"] = (cp.stdout or cp.stderr).strip()
    except Exception:
        info["Ollama version"] = "در دسترس نبود"
    return info


def add_monitor_sheets(wb, monitor, timing, api_result, client_elapsed, ollama_ps_text,
                       started_at, core_elapsed, model_name, ollama_ps_before=""):

    for sheet_name in ("گزارش عملکرد", "نمونه‌های پایش", "نمونه‌های GPU", "تفکیک آداپتورها", "جزئیات پردازش‌ها"):
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]

    samples = monitor.samples
    gpu = monitor.gpu_samples
    inference_samples = [x for x in samples if x.get("phase") == "استنتاج مدل" and not x.get("sample_error")]
    hw = hardware_snapshot()

    def stat(key, data=samples):
        return metric_stats(data, key)

    sys_cpu_avg, sys_cpu_max, _, sys_cpu_p95 = stat("system_cpu_pct")
    ram_avg, ram_max, _, _ = stat("system_ram_used_gb")
    ram_available_avg, ram_available_max, ram_available_min, _ = stat("system_ram_available_gb")
    oll_cpu_avg, oll_cpu_max, _, oll_cpu_p95 = stat("ollama_cpu_machine_pct")
    oll_ram_avg, oll_ram_max, _, _ = stat("ollama_rss_mb")
    py_ram_avg, py_ram_max, _, _ = stat("python_rss_mb")
    inf_cpu_avg, inf_cpu_max, _, _ = stat("system_cpu_pct", inference_samples)
    inf_ram_avg, inf_ram_max, _, _ = stat("system_ram_used_gb", inference_samples)
    inf_oll_cpu_avg, inf_oll_cpu_max, _, _ = stat("ollama_cpu_machine_pct", inference_samples)
    inf_oll_ram_avg, inf_oll_ram_max, _, _ = stat("ollama_rss_mb", inference_samples)

    g_oll_util = [x["ollama_engine_max_pct"] for x in gpu if x.get("ollama_engine_max_pct") is not None]
    g_all_util = [x["all_engine_max_pct"] for x in gpu if x.get("all_engine_max_pct") is not None]
    g_ded = [x["ollama_dedicated_mb"] for x in gpu if x.get("ollama_dedicated_mb") is not None]
    g_shared = [x["ollama_shared_mb"] for x in gpu if x.get("ollama_shared_mb") is not None]
    g_all_ded = [x["all_process_dedicated_mb"] for x in gpu if x.get("all_process_dedicated_mb") is not None]

    disk_read_mb = disk_write_mb = None
    if len(samples) >= 2:
        disk_read_mb = max(0.0, (samples[-1].get("ollama_read_bytes_cumulative", 0) - samples[0].get("ollama_read_bytes_cumulative", 0)) / (1024 ** 2))
        disk_write_mb = max(0.0, (samples[-1].get("ollama_write_bytes_cumulative", 0) - samples[0].get("ollama_write_bytes_cumulative", 0)) / (1024 ** 2))

    prompt_tokens = api_result.get("prompt_eval_count")
    output_tokens = api_result.get("eval_count")
    eval_ns = api_result.get("eval_duration")
    tps = output_tokens / (eval_ns / 1e9) if isinstance(output_tokens, (int, float)) and isinstance(eval_ns, (int, float)) and eval_ns > 0 else None

    def api_sec(key):
        val = api_result.get(key)
        return val / 1e9 if isinstance(val, (int, float)) else None

    ws = wb.create_sheet("گزارش عملکرد")
    ws.sheet_view.rightToLeft = True
    set_title(ws, "A1:E1", f"گزارش عملکرد و مانیتورینگ V4.2 | {model_name}")
    ws.append(["شاخص", "مقدار", "واحد", "دامنه", "توضیحات اندازه‌گیری"])
    style_header(ws, 2, 1, 5)

    rows = [
        ("مدل انتخاب‌شده", model_name, "", "کل مأموریت", "از فهرست مدل‌های نصب‌شده Ollama انتخاب شده است"),
        ("زمان شروع پایش", started_at.strftime("%Y-%m-%d %H:%M:%S"), "", "کل مأموریت", "پس از انتخاب فایل و مدل"),
        ("زمان کل پردازش", core_elapsed, "ثانیه", "کل مأموریت", "از شروع خواندن فایل تا ذخیره اولیه خروجی؛ بدون زمان بازکردن پنجره‌ها"),
        ("زمان خواندن و استخراج", timing.get("read_extract"), "ثانیه", "برنامه", "بازکردن workbook و استخراج جدول‌ها"),
        ("زمان ساخت داشبورد", timing.get("workbook_build"), "ثانیه", "برنامه", "ساخت جدول، تطبیق و نمودار اصلی"),
        ("زمان ذخیره اولیه فایل", timing.get("first_save"), "ثانیه", "برنامه", "ذخیره workbook پیش از برگه‌های پایش"),
        ("زمان ساخت و ذخیره گزارش پایش", timing.get("monitor_report_export"), "ثانیه", "برنامه", "ساخت گزارش عملکرد و نمودارهای مانیتورینگ؛ خارج از بازه پایش اصلی"),
        ("زمان درخواست مدل در برنامه", client_elapsed, "ثانیه", "مدل", "زمان رفت‌وبرگشت HTTP"),
        ("زمان کل Ollama", api_sec("total_duration"), "ثانیه", "مدل", "آمار دقیق بازگردانده‌شده توسط API"),
        ("زمان بارگذاری مدل", api_sec("load_duration"), "ثانیه", "مدل", "آمار Ollama"),
        ("زمان پردازش ورودی", api_sec("prompt_eval_duration"), "ثانیه", "مدل", "آمار Ollama"),
        ("زمان تولید پاسخ", api_sec("eval_duration"), "ثانیه", "مدل", "آمار Ollama"),
        ("تعداد توکن ورودی", prompt_tokens, "توکن", "مدل", "prompt_eval_count"),
        ("تعداد توکن خروجی", output_tokens, "توکن", "مدل", "eval_count"),
        ("سرعت تولید خروجی", tps, "توکن/ثانیه", "مدل", "eval_count تقسیم بر eval_duration"),
        ("میانگین CPU سیستم", sys_cpu_avg, "%", "کل مأموریت", "نمونه‌برداری حدود هر یک ثانیه"),
        ("حداکثر CPU سیستم", sys_cpu_max, "%", "کل مأموریت", "اوج نمونه‌برداری‌شده"),
        ("صدک ۹۵ CPU سیستم", sys_cpu_p95, "%", "کل مأموریت", "برای کاهش اثر جهش‌های تک‌نمونه‌ای"),
        ("میانگین RAM مصرفی سیستم", ram_avg, "GB", "کل مأموریت", "مصرف کل RAM ویندوز"),
        ("حداکثر RAM مصرفی سیستم", ram_max, "GB", "کل مأموریت", "اوج نمونه‌برداری‌شده"),
        ("میانگین RAM آزاد", ram_available_avg, "GB", "کل مأموریت", "از حافظه Available ویندوز"),
        ("حداکثر RAM آزاد", ram_available_max, "GB", "کل مأموریت", "بیشترین Available ثبت‌شده"),
        ("حداقل RAM آزاد", ram_available_min, "GB", "کل مأموریت", "کمترین Available ثبت‌شده؛ اصلاح محاسبه V4"),
        ("میانگین CPU Ollama/Runner", oll_cpu_avg, "% کل CPU", "کل مأموریت", "CPU تجمعی پردازش‌های تشخیص‌داده‌شده تقسیم بر تعداد رشته‌های منطقی"),
        ("حداکثر CPU Ollama/Runner", oll_cpu_max, "% کل CPU", "کل مأموریت", "حداکثر نمونه نرمال‌شده"),
        ("صدک ۹۵ CPU Ollama/Runner", oll_cpu_p95, "% کل CPU", "کل مأموریت", "صدک ۹۵ نمونه‌ها"),
        ("میانگین RAM Ollama/Runner", oll_ram_avg, "MB", "کل مأموریت", "مجموع RSS پردازش‌های شناسایی‌شده"),
        ("حداکثر RAM Ollama/Runner", oll_ram_max, "MB", "کل مأموریت", "اوج RSS پردازش‌های شناسایی‌شده"),
        ("میانگین RAM برنامه Python", py_ram_avg, "MB", "کل مأموریت", "RSS خود ابزار؛ از پردازش Ollama جداست"),
        ("اوج RAM برنامه Python", py_ram_max, "MB", "کل مأموریت", "اوج RSS خود ابزار"),
        ("میانگین CPU سیستم حین استنتاج", inf_cpu_avg, "%", "فقط زمان استنتاج", "نمونه‌هایی که فاز آن‌ها استنتاج مدل بوده است"),
        ("اوج CPU سیستم حین استنتاج", inf_cpu_max, "%", "فقط زمان استنتاج", "نمونه‌های همان فاز"),
        ("میانگین CPU Ollama حین استنتاج", inf_oll_cpu_avg, "% کل CPU", "فقط زمان استنتاج", "پردازش‌های Ollama/Runner شناسایی‌شده"),
        ("اوج CPU Ollama حین استنتاج", inf_oll_cpu_max, "% کل CPU", "فقط زمان استنتاج", "اوج نمونه همان فاز"),
        ("میانگین RAM سیستم حین استنتاج", inf_ram_avg, "GB", "فقط زمان استنتاج", "مصرف کل سیستم در فاز مدل"),
        ("اوج RAM سیستم حین استنتاج", inf_ram_max, "GB", "فقط زمان استنتاج", "اوج نمونه همان فاز"),
        ("میانگین RAM Ollama حین استنتاج", inf_oll_ram_avg, "MB", "فقط زمان استنتاج", "مجموع RSS backend"),
        ("اوج RAM Ollama حین استنتاج", inf_oll_ram_max, "MB", "فقط زمان استنتاج", "اوج RSS backend"),
        ("میانگین GPU Engine منتسب به Ollama", sum(g_oll_util)/len(g_oll_util) if g_oll_util else None, "% یک Engine", "GPU", "بیشترین Engine در هر نمونه؛ درصد کل کارت نیست"),
        ("حداکثر GPU Engine منتسب به Ollama", max(g_oll_util) if g_oll_util else None, "% یک Engine", "GPU", "شمارنده ویندوز؛ درایور ممکن است پشتیبانی نکند"),
        ("حداکثر Dedicated VRAM منتسب به Ollama", max(g_ded) if g_ded else None, "MB", "GPU", "از شمارنده GPU Process Memory بر اساس PID"),
        ("حداکثر Shared GPU Memory منتسب به Ollama", max(g_shared) if g_shared else None, "MB", "GPU", "حافظه اشتراکی ثبت‌شده بر اساس PID"),
        ("حداکثر Dedicated Memory همه برنامه‌ها", max(g_all_ded) if g_all_ded else None, "MB", "GPU", "مجموع همه پردازش‌ها؛ مختص Ollama نیست"),
        ("حداکثر GPU Engine کل سیستم", max(g_all_util) if g_all_util else None, "% یک Engine", "GPU", "حداکثر Engine مشاهده‌شده، نه درصد کلی GPU"),
        ("میانگین VRAM Dedicated مصرفی ComfyUI", metric_stats(gpu, "comfyui_dedicated_mb")[0], "MB", "GPU", "فقط پردازش‌های ComfyUI شناسایی‌شده بر اساس PID"),
        ("حداکثر VRAM Dedicated مصرفی ComfyUI", metric_stats(gpu, "comfyui_dedicated_mb")[1], "MB", "GPU", "فقط پردازش‌های ComfyUI شناسایی‌شده بر اساس PID"),
        ("میانگین CPU نرمال‌شده ComfyUI", stat("comfyui_cpu_machine_pct")[0], "% کل CPU", "ComfyUI", "در صورت شناسایی پردازش Python مربوط به ComfyUI"),
        ("حداکثر RAM پردازش ComfyUI", stat("comfyui_rss_mb")[1], "MB", "ComfyUI", "در صورت شناسایی پردازش Python مربوط به ComfyUI"),
        ("میانه فاصله واقعی نمونه‌های GPU", median_gpu_interval(gpu), "ثانیه", "مانیتور GPU", "فاصله بین شروع دو پرس‌وجوی متوالی؛ هدف ۲٫۵ ثانیه، مقدار واقعی گزارش می‌شود"),
        ("میانه زمان پرس‌وجوی شمارنده GPU", metric_stats(gpu, "query_elapsed_sec")[0], "ثانیه", "مانیتور GPU", "زمان اجرای یک درخواست ترکیبی Get-Counter"),
        ("تعداد نمونه‌های GPU موفق", sum(1 for x in gpu if x.get("available")), "نمونه", "مانیتور GPU", "نمونه‌های موفق از یک درخواست ترکیبی برای Engine و حافظه پردازش"),
        ("تعداد شناسه‌های آداپتور مشاهده‌شده", len({a.get("adapter_key") for x in gpu for a in x.get("adapter_rows", [])}), "آداپتور", "GPU", "شناسه LUID و phys استخراج‌شده از شمارنده ویندوز؛ نام WMI جداگانه فهرست شده است"),
        ("خواندن دیسک پردازش‌های Ollama", disk_read_mb, "MB", "Disk I/O", "اختلاف شمارنده تجمعی اول و آخر"),
        ("نوشتن دیسک پردازش‌های Ollama", disk_write_mb, "MB", "Disk I/O", "اختلاف شمارنده تجمعی اول و آخر"),
        ("تعداد نمونه‌های CPU/RAM", len(samples), "نمونه", "مانیتور", "تقریباً هر یک ثانیه"),
    ]
    for label, value, unit, domain, note in rows:
        shown = value if value is not None else "در دسترس نبود"
        ws.append([label, shown, unit, domain, note])
        if isinstance(value, (int, float)):
            ws.cell(ws.max_row, 2).number_format = "#,##0.00"
    style_body(ws, 3, ws.max_row, 1, 5)

    # Hardware inventory section
    marker = ws.max_row + 2
    set_title(ws, f"A{marker}:E{marker}", "مشخصات سخت‌افزار میزبان")
    for key, value in hw.items():
        ws.append([key, value, "", "سیستم", "مقادیر جمع‌آوری‌شده از psutil/WMI"])
    style_body(ws, marker + 1, ws.max_row, 1, 5)

    # Process identification and ollama ps snapshots
    marker = ws.max_row + 2
    set_title(ws, f"A{marker}:E{marker}", "خروجی ollama ps پیش از مأموریت")
    marker += 1
    ws.merge_cells(start_row=marker, start_column=1, end_row=marker + 2, end_column=5)
    ws.cell(marker, 1, ollama_ps_before or "خروجی در دسترس نبود.")
    ws.cell(marker, 1).font = Font(name="Consolas", size=9)
    ws.cell(marker, 1).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    marker += 4
    set_title(ws, f"A{marker}:E{marker}", "خروجی ollama ps پس از پاسخ مدل")
    marker += 1
    ws.merge_cells(start_row=marker, start_column=1, end_row=marker + 2, end_column=5)
    ws.cell(marker, 1, ollama_ps_text or "خروجی در دسترس نبود.")
    ws.cell(marker, 1).font = Font(name="Consolas", size=9)
    ws.cell(marker, 1).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    marker += 4
    set_title(ws, f"A{marker}:E{marker}", "پردازش‌های Ollama/Runner شناسایی‌شده")
    marker += 1
    ws.append(["PID", "نام پردازش", "", "", ""])
    style_header(ws, marker, 1, 5)
    if monitor.detected_processes:
        for pid, name in sorted(monitor.detected_processes.items()):
            ws.append([pid, name, "", "", ""])
    else:
        ws.append(["پردازش backend شناسایی نشد", "", "", "", ""])

    notes_row = ws.max_row + 2
    set_title(ws, f"A{notes_row}:E{notes_row}", "محدودیت‌ها و تفسیر صحیح شاخص‌ها")
    notes = [
        "مصرف RAM آزاد به‌درستی با حداقل/حداکثر جداگانه محاسبه شده است؛ نسخه V4 در برچسب حداقل/حداکثر یک جابه‌جایی داشت.",
        "پردازش Python این ابزار عمداً از گروه Ollama حذف شده تا نام فایل حاوی Ollama باعث شمارش اشتباه نشود.",
        "CPU Ollama به دو شکل قابل تفسیر است: درصد نرمال‌شده از کل توان CPU و درصد تجمعی هسته‌ها؛ گزارش از مقدار نرمال‌شده استفاده می‌کند.",
        "درصد GPU Engine، درصد کلی کارت گرافیک نیست. چند Engine مستقل وجود دارند؛ بنابراین فقط به‌عنوان شاخص فعالیت Engine گزارش می‌شود.",
        "VRAM و Shared GPU Memory فقط در صورت وجود شمارنده ویندوز و امکان اتصال PID ثبت می‌شوند. نبود داده به معنی صفر بودن مصرف نیست.",
        "ollama ps محل بارگذاری مدل (GPU/CPU) را نشان می‌دهد، اما درصد اشغال محاسباتی GPU را نشان نمی‌دهد.",
        "دمای GPU/CPU و توان مصرفی وات در این نسخه اندازه‌گیری نمی‌شود؛ برای آن‌ها ابزار سنسورخوان جداگانه و سازگار با RX 580 لازم است.",
        "نمونه‌برداری GPU با یک فراخوانی ترکیبی Get-Counter انجام می‌شود؛ فاصله واقعی بین نمونه‌ها جداگانه محاسبه شده و ممکن است به دلیل تأخیر PowerShell از هدف ۲٫۵ ثانیه بیشتر شود.",
        "LUID و phys شناسه آداپتور از شمارنده Windows هستند. نگاشت خودکار LUID به نام RX 580 یا GT 440 قطعی فرض نشده؛ نام/درایور کارت‌ها در موجودی WMI جداگانه ثبت می‌شود.",
        "در آزمون تولید تصویر، ComfyUI فقط زمانی در آمار پردازش/GPU جدا می‌شود که فرایند Python مربوط به main.py شناسایی شود.",
    ]
    for i, note in enumerate(notes, start=notes_row + 1):
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=5)
        ws.cell(i, 1, "• " + note)
        ws.cell(i, 1).alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
        ws.row_dimensions[i].height = 30
    for col, width in {"A": 46, "B": 24, "C": 18, "D": 22, "E": 76}.items():
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A3"

    # Detailed time-series rows; last valid GPU sample is aligned to each CPU/RAM timestamp.
    raw = wb.create_sheet("نمونه‌های پایش")
    raw.sheet_view.rightToLeft = True
    set_title(raw, "A1:O1", "تاریخچه زمانی مصرف منابع و فاز اجرای برنامه")
    headers = ["زمان سپری‌شده ثانیه", "فاز", "CPU سیستم %", "RAM سیستم مصرفی GB", "RAM سیستم آزاد GB",
               "CPU Ollama از کل CPU %", "CPU تجمعی پردازش‌ها %", "RAM Ollama RSS GB", "RAM Python MB",
               "Dedicated VRAM Ollama MB", "Shared GPU Memory Ollama MB", "GPU Engine Ollama Max %",
               "GPU Engine کل سیستم Max %", "خواندن دیسک تجمعی MB", "نوشتن دیسک تجمعی MB"]
    raw.append(headers)
    style_header(raw, 2, 1, len(headers))
    gpu_sorted = sorted(gpu, key=lambda x: x["timestamp"])
    gi = 0
    last_gpu = None
    for sample in samples:
        while gi < len(gpu_sorted) and gpu_sorted[gi]["timestamp"] <= sample["timestamp"]:
            if gpu_sorted[gi].get("available"):
                last_gpu = gpu_sorted[gi]
            gi += 1
        raw.append([
            sample.get("elapsed_sec"), sample.get("phase"), sample.get("system_cpu_pct"),
            sample.get("system_ram_used_gb"), sample.get("system_ram_available_gb"),
            sample.get("ollama_cpu_machine_pct"), sample.get("ollama_cpu_raw_pct"),
            (sample.get("ollama_rss_mb") / 1024.0) if sample.get("ollama_rss_mb") is not None else None, sample.get("python_rss_mb"),
            last_gpu.get("ollama_dedicated_mb") if last_gpu else None,
            last_gpu.get("ollama_shared_mb") if last_gpu else None,
            last_gpu.get("ollama_engine_max_pct") if last_gpu else None,
            last_gpu.get("all_engine_max_pct") if last_gpu else None,
            sample.get("ollama_read_bytes_cumulative", 0) / (1024 ** 2),
            sample.get("ollama_write_bytes_cumulative", 0) / (1024 ** 2),
        ])
    if raw.max_row >= 3:
        style_body(raw, 3, raw.max_row, 1, len(headers))
        raw.auto_filter.ref = f"A2:O{raw.max_row}"

        cpu_chart = LineChart()
        cpu_chart.title = "روند CPU سیستم و Ollama"
        cpu_chart.y_axis.title = "درصد"
        cpu_chart.x_axis.title = "زمان (ثانیه)"
        cpu_chart.height = 7
        cpu_chart.width = 14
        cpu_chart.add_data(Reference(raw, min_col=3, max_col=3, min_row=2, max_row=raw.max_row), titles_from_data=True)
        cpu_chart.add_data(Reference(raw, min_col=6, max_col=6, min_row=2, max_row=raw.max_row), titles_from_data=True)
        cpu_chart.set_categories(Reference(raw, min_col=1, min_row=3, max_row=raw.max_row))
        raw.add_chart(cpu_chart, "Q2")

        ram_chart = LineChart()
        ram_chart.title = "روند RAM سیستم و پردازش Ollama"
        ram_chart.y_axis.title = "GB"
        ram_chart.x_axis.title = "زمان (ثانیه)"
        ram_chart.height = 7
        ram_chart.width = 14
        ram_chart.add_data(Reference(raw, min_col=4, max_col=4, min_row=2, max_row=raw.max_row), titles_from_data=True)
        ram_chart.add_data(Reference(raw, min_col=8, max_col=8, min_row=2, max_row=raw.max_row), titles_from_data=True)
        ram_chart.set_categories(Reference(raw, min_col=1, min_row=3, max_row=raw.max_row))
        raw.add_chart(ram_chart, "Q18")

        gpu_chart = LineChart()
        gpu_chart.title = "فعالیت GPU Engine منتسب به Ollama"
        gpu_chart.y_axis.title = "درصد Engine"
        gpu_chart.x_axis.title = "زمان (ثانیه)"
        gpu_chart.height = 7
        gpu_chart.width = 14
        gpu_chart.add_data(Reference(raw, min_col=12, max_col=12, min_row=2, max_row=raw.max_row), titles_from_data=True)
        gpu_chart.set_categories(Reference(raw, min_col=1, min_row=3, max_row=raw.max_row))
        raw.add_chart(gpu_chart, "Q34")

    for col in range(1, len(headers) + 1):
        raw.column_dimensions[get_column_letter(col)].width = 23
    raw.column_dimensions["B"].width = 22
    raw.freeze_panes = "A3"

    # Raw GPU counter samples: keep them separate because GPU sampling cadence differs from CPU/RAM.
    gpu_ws = wb.create_sheet("نمونه‌های GPU")
    gpu_ws.sheet_view.rightToLeft = True
    set_title(gpu_ws, "A1:P1", "نمونه‌های خام شمارنده‌های GPU ویندوز")
    gpu_headers = ["زمان سپری‌شده ثانیه", "فاز", "داده موجود", "فاصله واقعی نمونه ثانیه", "زمان پرس‌وجو ثانیه",
                   "Ollama Engine Max %", "Ollama Dedicated MB", "Ollama Shared MB", "ComfyUI Engine Max %",
                   "ComfyUI Dedicated MB", "ComfyUI Shared MB", "کل Engine Max %",
                   "Dedicated همه پردازش‌ها MB", "PIDهای Ollama", "PIDهای ComfyUI", "خطای شمارنده"]
    gpu_ws.append(gpu_headers)
    style_header(gpu_ws, 2, 1, len(gpu_headers))
    for item in gpu:
        gpu_ws.append([
            max(0.0, item.get("timestamp", time.time()) - monitor.start_wall.timestamp()) if monitor.start_wall else None,
            item.get("phase", ""), "بله" if item.get("available") else "خیر",
            item.get("sample_interval_sec"), item.get("query_elapsed_sec"),
            item.get("ollama_engine_max_pct"), item.get("ollama_dedicated_mb"),
            item.get("ollama_shared_mb"), item.get("comfyui_engine_max_pct"),
            item.get("comfyui_dedicated_mb"), item.get("comfyui_shared_mb"),
            item.get("all_engine_max_pct"), item.get("all_process_dedicated_mb"),
            item.get("ollama_pids", ""), item.get("comfyui_pids", ""), item.get("error", "")
        ])
    if gpu_ws.max_row >= 3:
        style_body(gpu_ws, 3, gpu_ws.max_row, 1, len(gpu_headers))
        gpu_ws.auto_filter.ref = f"A2:J{gpu_ws.max_row}"
    for col, width in {"A": 23, "B": 22, "C": 14, "D": 22, "E": 21, "F": 23, "G": 23, "H": 23, "I": 23, "J": 23, "K": 23, "L": 20, "M": 28, "N": 24, "O": 24, "P": 50}.items():
        gpu_ws.column_dimensions[col].width = width
    gpu_ws.freeze_panes = "A3"

    # Per-adapter view based on Windows LUID + phys identifiers. Do not guess physical model names.
    adapter_ws = wb.create_sheet("تفکیک آداپتورها")
    adapter_ws.sheet_view.rightToLeft = True
    set_title(adapter_ws, "A1:Q1", "تفکیک مصرف GPU بر اساس LUID و phys")
    adapter_headers = ["زمان سپری‌شده ثانیه", "فاز", "Adapter Key", "LUID", "phys index", "Engine count",
                       "Ollama Engine Max %", "Ollama Engine Sum %", "ComfyUI Engine Max %", "ComfyUI Engine Sum %",
                       "Ollama Dedicated MB", "Ollama Shared MB", "ComfyUI Dedicated MB", "ComfyUI Shared MB",
                       "All Process Dedicated MB", "Engine Types", "Ollama PID / ComfyUI PID"]
    adapter_ws.append(adapter_headers)
    style_header(adapter_ws, 2, 1, len(adapter_headers))
    for sample in gpu:
        elapsed = max(0.0, sample.get("timestamp", time.time()) - monitor.start_wall.timestamp()) if monitor.start_wall else None
        for a in sample.get("adapter_rows", []):
            adapter_ws.append([
                elapsed, sample.get("phase", ""), a.get("adapter_key"), a.get("luid"), a.get("physical_index"),
                a.get("engine_count"), a.get("ollama_engine_max_pct"), a.get("ollama_engine_sum_pct"),
                a.get("comfyui_engine_max_pct"), a.get("comfyui_engine_sum_pct"),
                a.get("ollama_dedicated_mb"), a.get("ollama_shared_mb"),
                a.get("comfyui_dedicated_mb"), a.get("comfyui_shared_mb"),
                a.get("all_process_dedicated_mb"), a.get("engine_types"),
                f"{a.get('ollama_pids','')} / {a.get('comfyui_pids','')}"
            ])
    if adapter_ws.max_row >= 3:
        style_body(adapter_ws, 3, adapter_ws.max_row, 1, len(adapter_headers))
        adapter_ws.auto_filter.ref = f"A2:Q{adapter_ws.max_row}"
    for col, width in {"A": 22, "B": 20, "C": 46, "D": 28, "E": 13, "F": 15, "G": 22, "H": 22,
                       "I": 22, "J": 22, "K": 22, "L": 22, "M": 22, "N": 22, "O": 25, "P": 24, "Q": 36}.items():
        adapter_ws.column_dimensions[col].width = width
    adapter_ws.freeze_panes = "A3"

    # Per-process RSS/CPU history, with the Python assistant excluded from Ollama totals.
    proc_ws = wb.create_sheet("جزئیات پردازش‌ها")
    proc_ws.sheet_view.rightToLeft = True
    set_title(proc_ws, "A1:G1", "جزئیات پردازش‌های Ollama و ComfyUI")
    proc_headers = ["زمان سپری‌شده ثانیه", "فاز", "گروه", "PID", "نام پردازش", "RAM RSS MB", "زمان CPU تجمعی ثانیه"]
    proc_ws.append(proc_headers)
    style_header(proc_ws, 2, 1, len(proc_headers))
    for sample in samples:
        for proc in sample.get("process_rows", []):
            proc_ws.append([sample.get("elapsed_sec"), sample.get("phase"), proc.get("role"), proc.get("pid"),
                            proc.get("name"), proc.get("rss_mb"), proc.get("cpu_seconds")])
    if proc_ws.max_row >= 3:
        style_body(proc_ws, 3, proc_ws.max_row, 1, len(proc_headers))
        proc_ws.auto_filter.ref = f"A2:F{proc_ws.max_row}"
    for col, width in {"A": 23, "B": 22, "C": 14, "D": 14, "E": 30, "F": 20, "G": 28}.items():
        proc_ws.column_dimensions[col].width = width
    proc_ws.freeze_panes = "A3"


def http_json(base_url, path, method="GET", payload=None, timeout=20):
    url = base_url.rstrip("/") + path
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8") if payload is not None else None
    headers = {"Content-Type": "application/json"} if data is not None else {}
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    with urllib.request.urlopen(req, timeout=timeout) as response:
        raw = response.read()
        content_type = response.headers.get("Content-Type", "")
        if "application/json" in content_type or raw[:1] in (b"{", b"["):
            return json.loads(raw.decode("utf-8"))
        return raw


def comfy_checkpoints(base_url):
    """Read available checkpoint names from the standard ComfyUI model API."""
    try:
        value = http_json(base_url, "/models/checkpoints", timeout=12)
        names = []
        if isinstance(value, list):
            for item in value:
                if isinstance(item, str): names.append(item)
                elif isinstance(item, dict): names.append(item.get("name") or item.get("filename") or "")
        elif isinstance(value, dict):
            for item in value.get("models", []):
                if isinstance(item, str): names.append(item)
                elif isinstance(item, dict): names.append(item.get("name") or item.get("filename") or "")
        names = sorted({x for x in names if x})
        if names:
            return names
    except Exception:
        pass
    # Compatibility fallback: read the checkpoint dropdown values from node metadata.
    info = http_json(base_url, "/object_info/CheckpointLoaderSimple", timeout=12)
    node = info.get("CheckpointLoaderSimple", info) if isinstance(info, dict) else {}
    required = (node.get("input") or {}).get("required") or {}
    values = required.get("ckpt_name") or []
    if values and isinstance(values[0], list):
        return [str(x) for x in values[0]]
    return []


def build_comfy_workflow(checkpoint, positive, negative, width=512, height=512, steps=20, cfg=7.0, seed=None):
    import secrets
    if seed is None:
        seed = secrets.randbelow(2**32)
    # Standard ComfyUI API graph for SD 1.x / SD 2.x checkpoint-based workflows.
    # For RX 580, start with a Stable Diffusion 1.5 checkpoint at 512x512.
    return {
        "1": {"class_type": "CheckpointLoaderSimple", "inputs": {"ckpt_name": checkpoint}},
        "2": {"class_type": "CLIPTextEncode", "inputs": {"text": positive, "clip": ["1", 1]}},
        "3": {"class_type": "CLIPTextEncode", "inputs": {"text": negative, "clip": ["1", 1]}},
        "4": {"class_type": "EmptyLatentImage", "inputs": {"width": int(width), "height": int(height), "batch_size": 1}},
        "5": {"class_type": "KSampler", "inputs": {
            "seed": int(seed), "steps": int(steps), "cfg": float(cfg), "sampler_name": "euler",
            "scheduler": "normal", "denoise": 1.0, "model": ["1", 0],
            "positive": ["2", 0], "negative": ["3", 0], "latent_image": ["4", 0]
        }},
        "6": {"class_type": "VAEDecode", "inputs": {"samples": ["5", 0], "vae": ["1", 2]}},
        "7": {"class_type": "SaveImage", "inputs": {"filename_prefix": "LocalOfficeAssistant", "images": ["6", 0]}}
    }


def improve_image_prompt(model_name, user_prompt):
    body = {
        "model": model_name, "stream": False, "think": False, "keep_alive": 0,
        "options": {"temperature": 0.35, "num_ctx": 4096},
        "messages": [
            {"role": "system", "content": (
                "You are an expert Stable Diffusion 1.5 prompt designer. Convert the user's request into one detailed, "
                "coherent English image-generation prompt. Preserve the requested subject and intent. Specify composition, "
                "lighting, materials, environment, camera/viewpoint and visual style when appropriate. Do not add text, logos, "
                "watermarks, or extra subjects unless requested. Return only the final prompt, without explanation or quotation marks."
            )},
            {"role": "user", "content": user_prompt}
        ]
    }
    req = urllib.request.Request(OLLAMA_CHAT_URL, data=json.dumps(body, ensure_ascii=False).encode("utf-8"),
                                 headers={"Content-Type": "application/json"}, method="POST")
    t0 = time.perf_counter()
    with urllib.request.urlopen(req, timeout=900) as response:
        result = json.loads(response.read().decode("utf-8"))
    elapsed = time.perf_counter() - t0
    # keep_alive=0 asks Ollama to unload the model; an explicit stop is a best-effort fallback
    # so the image-generation backend has a better chance of getting the RX 580 VRAM back.
    try:
        subprocess.run(["ollama", "stop", model_name], capture_output=True, text=True, timeout=45,
                       creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except Exception:
        pass
    text = (result.get("message") or {}).get("content", "").strip()
    if not text:
        raise RuntimeError("Qwen پاسخ قابل استفاده برای پرامپت تصویر برنگرداند.")
    return text, result, elapsed


def submit_comfy_image(base_url, workflow, progress_callback=None, timeout_seconds=1200):
    import uuid
    client_id = str(uuid.uuid4())
    response = http_json(base_url, "/prompt", method="POST", payload={"prompt": workflow, "client_id": client_id}, timeout=30)
    if not isinstance(response, dict) or not response.get("prompt_id"):
        raise RuntimeError("ComfyUI درخواست را نپذیرفت: " + json.dumps(response, ensure_ascii=False)[:1200])
    prompt_id = response["prompt_id"]
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        history = http_json(base_url, f"/history/{prompt_id}", timeout=20)
        entry = history.get(prompt_id, {}) if isinstance(history, dict) else {}
        if isinstance(entry, dict):
            status = entry.get("status") or {}
            status_text = str(status.get("status_str", "")).lower() if isinstance(status, dict) else ""
            if status_text in {"error", "failed"}:
                messages = status.get("messages", []) if isinstance(status, dict) else []
                raise RuntimeError("ComfyUI workflow failed: " + json.dumps(messages, ensure_ascii=False)[:1600])
        outputs = entry.get("outputs", {}) if isinstance(entry, dict) else {}
        image_items = []
        for node_output in outputs.values():
            if isinstance(node_output, dict):
                image_items.extend(node_output.get("images", []))
        if image_items:
            return prompt_id, image_items
        if progress_callback:
            progress_callback("ComfyUI در حال تولید تصویر است... شناسه کار: " + str(prompt_id)[:12])
        time.sleep(2.0)
    raise TimeoutError("زمان انتظار برای ComfyUI تمام شد؛ شناسه کار: " + str(prompt_id))


def download_comfy_images(base_url, image_items, output_dir):
    from urllib.parse import urlencode
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    saved = []
    for i, item in enumerate(image_items, start=1):
        query = urlencode({
            "filename": item.get("filename", ""),
            "subfolder": item.get("subfolder", ""),
            "type": item.get("type", "output")
        })
        req = urllib.request.Request(base_url.rstrip("/") + "/view?" + query)
        with urllib.request.urlopen(req, timeout=120) as response:
            data = response.read()
        suffix = Path(item.get("filename", "image.png")).suffix or ".png"
        out = output_dir / f"generated_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{i}{suffix}"
        out.write_bytes(data)
        saved.append(out)
    return saved


def save_image_monitor_report(monitor, output_dir, model_name, image_prompt, positive_prompt,
                              comfy_prompt_id, qwen_elapsed, generation_started, generation_finished,
                              saved_paths, api_result=None):
    from openpyxl import Workbook
    output_dir = Path(output_dir)
    report_path = output_dir / f"Image_Generation_Monitor_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "گزارش عملکرد"
    ws.sheet_view.rightToLeft = True
    ws.append(["شاخص", "مقدار", "توضیحات"])
    style_header(ws, 1, 1, 3)
    samples = [x for x in monitor.samples if not x.get("sample_error")]
    gpu = monitor.gpu_samples
    def avg_peak(key):
        vals = [float(x[key]) for x in samples if isinstance(x.get(key), (int, float))]
        return (sum(vals)/len(vals), max(vals)) if vals else (None, None)
    cpu_avg, cpu_peak = avg_peak("system_cpu_pct")
    ram_avg, ram_peak = avg_peak("system_ram_used_gb")
    oll_ded = [x["ollama_dedicated_mb"] for x in gpu if isinstance(x.get("ollama_dedicated_mb"), (int, float))]
    comfy_ded = [x["comfyui_dedicated_mb"] for x in gpu if isinstance(x.get("comfyui_dedicated_mb"), (int, float))]
    engine = [x["comfyui_engine_max_pct"] for x in gpu if isinstance(x.get("comfyui_engine_max_pct"), (int, float))]
    intervals = [x["sample_interval_sec"] for x in gpu if isinstance(x.get("sample_interval_sec"), (int, float))]
    rows = [
        ("مدل Qwen برای بهینه‌سازی پرامپت", model_name, "در صورت فعال بودن گزینه بهینه‌سازی"),
        ("شناسه ComfyUI prompt", comfy_prompt_id, "شناسه درخواست API"),
        ("زمان بهینه‌سازی پرامپت با Qwen", qwen_elapsed, "ثانیه"),
        ("زمان تولید و دریافت تصویر", (generation_finished-generation_started), "ثانیه؛ شامل صف و اجرای ComfyUI"),
        ("تعداد تصاویر ذخیره‌شده", len(saved_paths), "فایل"),
        ("میانگین CPU سیستم", cpu_avg, "%"), ("حداکثر CPU سیستم", cpu_peak, "%"),
        ("میانگین RAM مصرفی سیستم", ram_avg, "GB"), ("حداکثر RAM مصرفی سیستم", ram_peak, "GB"),
        ("حداکثر Dedicated VRAM ComfyUI", max(comfy_ded) if comfy_ded else None, "MB؛ بر اساس PID و شمارنده Windows"),
        ("حداکثر Dedicated VRAM Ollama", max(oll_ded) if oll_ded else None, "MB؛ اگر PIDهای Ollama در زمان نمونه موجود بوده باشند"),
        ("حداکثر GPU Engine منتسب به ComfyUI", max(engine) if engine else None, "% یک Engine، نه درصد کل GPU"),
        ("میانه فاصله نمونه‌برداری GPU", sorted(intervals)[len(intervals)//2] if intervals else None, "ثانیه"),
        ("تعداد نمونه‌های CPU/RAM", len(samples), "نمونه"),
        ("تعداد نمونه‌های موفق GPU", sum(1 for x in gpu if x.get("available")), "نمونه"),
        ("مسیر تصاویر", "\n".join(str(p) for p in saved_paths), "فایل‌های ذخیره‌شده محلی"),
        ("پرامپت اولیه", image_prompt, "درخواست اصلی کاربر"),
        ("پرامپت ارسال‌شده به ComfyUI", positive_prompt, "متن نهایی تولید تصویر"),
    ]
    for row in rows:
        ws.append(list(row))
    style_body(ws, 2, ws.max_row, 1, 3)
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 70
    ws.column_dimensions["C"].width = 48

    raw = wb.create_sheet("نمونه‌های سیستم")
    raw.sheet_view.rightToLeft = True
    headers = ["Elapsed sec", "Phase", "CPU system %", "RAM used GB", "RAM available GB", "Ollama CPU %", "Ollama RSS MB", "ComfyUI CPU %", "ComfyUI RSS MB", "Ollama PIDs", "ComfyUI PIDs"]
    raw.append(headers); style_header(raw, 1, 1, len(headers))
    for x in samples:
        raw.append([x.get("elapsed_sec"), x.get("phase"), x.get("system_cpu_pct"), x.get("system_ram_used_gb"),
                    x.get("system_ram_available_gb"), x.get("ollama_cpu_machine_pct"), x.get("ollama_rss_mb"),
                    x.get("comfyui_cpu_machine_pct"), x.get("comfyui_rss_mb"), x.get("ollama_pids"), x.get("comfyui_pids")])
    raw.freeze_panes = "A2"
    raw.auto_filter.ref = f"A1:K{raw.max_row}"
    for col in range(1, len(headers)+1): raw.column_dimensions[get_column_letter(col)].width = 22

    gpu_ws = wb.create_sheet("نمونه‌های GPU")
    gpu_ws.sheet_view.rightToLeft = True
    gpu_headers = ["Elapsed sec", "Phase", "Available", "Interval sec", "Query sec", "Ollama Engine Max %", "Ollama Dedicated MB", "ComfyUI Engine Max %", "ComfyUI Dedicated MB", "Adapter rows", "Error"]
    gpu_ws.append(gpu_headers); style_header(gpu_ws, 1, 1, len(gpu_headers))
    for x in gpu:
        adapter_summary = " | ".join(f"{a.get('adapter_key')}: OllamaVRAM={a.get('ollama_dedicated_mb')}MB; ComfyUIVRAM={a.get('comfyui_dedicated_mb')}MB" for a in x.get("adapter_rows", []))
        gpu_ws.append([max(0,x.get("timestamp",time.time())-monitor.start_wall.timestamp()), x.get("phase"), x.get("available"),
                       x.get("sample_interval_sec"), x.get("query_elapsed_sec"), x.get("ollama_engine_max_pct"),
                       x.get("ollama_dedicated_mb"), x.get("comfyui_engine_max_pct"), x.get("comfyui_dedicated_mb"),
                       adapter_summary, x.get("error", "")])
    gpu_ws.freeze_panes = "A2"
    gpu_ws.auto_filter.ref = f"A1:K{gpu_ws.max_row}"
    for col in range(1, len(gpu_headers)+1): gpu_ws.column_dimensions[get_column_letter(col)].width = 26

    adapters_ws = wb.create_sheet("تفکیک آداپتورها")
    adapters_ws.sheet_view.rightToLeft = True
    ah = ["Elapsed sec", "Adapter Key", "LUID", "Physical index", "Ollama Dedicated MB", "ComfyUI Dedicated MB", "Ollama Engine Max %", "ComfyUI Engine Max %", "Engine Types"]
    adapters_ws.append(ah); style_header(adapters_ws, 1, 1, len(ah))
    for x in gpu:
        elapsed = max(0,x.get("timestamp",time.time())-monitor.start_wall.timestamp())
        for a in x.get("adapter_rows", []):
            adapters_ws.append([elapsed, a.get("adapter_key"), a.get("luid"), a.get("physical_index"), a.get("ollama_dedicated_mb"),
                                a.get("comfyui_dedicated_mb"), a.get("ollama_engine_max_pct"), a.get("comfyui_engine_max_pct"), a.get("engine_types")])
    adapters_ws.freeze_panes = "A2"
    adapters_ws.auto_filter.ref = f"A1:I{adapters_ws.max_row}"
    for col in range(1, len(ah)+1): adapters_ws.column_dimensions[get_column_letter(col)].width = 26

    proc = wb.create_sheet("جزئیات پردازش‌ها")
    proc.sheet_view.rightToLeft = True
    ph = ["Elapsed sec", "Phase", "Role", "PID", "Process", "RSS MB", "CPU time sec"]
    proc.append(ph); style_header(proc, 1, 1, len(ph))
    for x in samples:
        for p in x.get("process_rows", []):
            proc.append([x.get("elapsed_sec"), x.get("phase"), p.get("role"), p.get("pid"), p.get("name"), p.get("rss_mb"), p.get("cpu_seconds")])
    proc.freeze_panes = "A2"
    proc.auto_filter.ref = f"A1:G{proc.max_row}"
    for col in range(1, len(ph)+1): proc.column_dimensions[get_column_letter(col)].width = 24
    wb.save(report_path)
    return report_path


def choose_task_dialog(root):
    import tkinter as tk
    from tkinter import ttk
    selected = {"value": None}
    win = tk.Toplevel(root)
    win.title("Local Office Assistant V4.2")
    win.geometry("520x250")
    win.resizable(False, False)
    win.grab_set()
    tk.Label(win, text="دستیار اداری آفلاین | نسخه ۴.۲", font=("Tahoma", 14, "bold")).pack(pady=(22, 8))
    tk.Label(win, text="یکی از دو بخش را انتخاب کن:", font=("Tahoma", 10)).pack(pady=5)
    buttons = tk.Frame(win)
    buttons.pack(pady=18)
    def select(value):
        selected["value"] = value
        win.destroy()
    ttk.Button(buttons, text="تحلیل و حسابرسی Excel", command=lambda: select("excel"), width=25).pack(pady=6)
    ttk.Button(buttons, text="استودیو تولید تصویر (ComfyUI)", command=lambda: select("image"), width=25).pack(pady=6)
    tk.Label(win, text="تولید تصویر از طریق ComfyUI محلی انجام می‌شود؛ Ollama فقط پرامپت را بهبود می‌دهد.",
             font=("Tahoma", 8), wraplength=470).pack(pady=6)
    win.protocol("WM_DELETE_WINDOW", lambda: select(None))
    win.wait_window()
    return selected["value"]


def launch_image_studio(root):
    import tkinter as tk
    from tkinter import ttk, messagebox, scrolledtext
    win = tk.Toplevel(root)
    win.title("استودیو تولید تصویر | Ollama + ComfyUI")
    win.geometry("840x760")
    win.minsize(760, 650)
    win.grab_set()

    models = []
    try:
        models = get_installed_qwen_models()
    except Exception:
        pass
    default_model = next((m for m in models if m.endswith(":27b")), next((m for m in models if m.endswith(":9b")), models[0] if models else ""))

    frame = tk.Frame(win, padx=14, pady=12)
    frame.pack(fill="both", expand=True)
    tk.Label(frame, text="استودیو تولید تصویر", font=("Tahoma", 15, "bold")).pack(anchor="e", pady=(0, 8))

    conn = tk.Frame(frame); conn.pack(fill="x", pady=4)
    tk.Label(conn, text="نشانی ComfyUI:", font=("Tahoma", 9, "bold")).pack(side="right", padx=6)
    base_var = tk.StringVar(value=COMFY_DEFAULT_URL)
    tk.Entry(conn, textvariable=base_var, width=38, justify="left").pack(side="right", padx=4)
    status_var = tk.StringVar(value="ComfyUI هنوز بررسی نشده است.")
    tk.Label(frame, textvariable=status_var, fg="#1F4E79", anchor="e").pack(fill="x", pady=4)

    cp_frame = tk.Frame(frame); cp_frame.pack(fill="x", pady=5)
    tk.Label(cp_frame, text="Checkpoint مدل تصویر:", font=("Tahoma", 9, "bold")).pack(side="right", padx=6)
    checkpoint_var = tk.StringVar()
    checkpoint_combo = ttk.Combobox(cp_frame, textvariable=checkpoint_var, width=48, state="normal")
    checkpoint_combo.pack(side="right", padx=4)
    def refresh_checkpoints():
        try:
            names = comfy_checkpoints(base_var.get().strip())
            checkpoint_combo["values"] = names
            if names and not checkpoint_var.get().strip(): checkpoint_var.set(names[0])
            status_var.set(f"اتصال برقرار است؛ {len(names)} مدل checkpoint شناسایی شد.")
            if not names: status_var.set("ComfyUI متصل است، اما checkpoint پیدا نشد. ابتدا یک SD 1.5 checkpoint نصب کن.")
        except Exception as exc:
            status_var.set("اتصال ناموفق: " + str(exc)[:220])
            messagebox.showerror("اتصال ComfyUI", f"ComfyUI پاسخ نداد. باید سرور روی 127.0.0.1:8188 اجرا باشد.\n\n{exc}", parent=win)
    ttk.Button(cp_frame, text="بررسی اتصال / تازه‌سازی مدل‌ها", command=refresh_checkpoints).pack(side="right", padx=6)

    model_frame = tk.Frame(frame); model_frame.pack(fill="x", pady=5)
    tk.Label(model_frame, text="مدل Qwen برای بهبود پرامپت:", font=("Tahoma", 9, "bold")).pack(side="right", padx=6)
    model_var = tk.StringVar(value=default_model)
    model_combo = ttk.Combobox(model_frame, textvariable=model_var, values=models, width=30, state="readonly" if models else "disabled")
    model_combo.pack(side="right", padx=4)
    improve_var = tk.BooleanVar(value=bool(models))
    ttk.Checkbutton(model_frame, text="بهبود پرامپت با Qwen (پس از آن مدل تخلیه می‌شود تا GPU آزاد شود)", variable=improve_var).pack(side="right", padx=8)

    tk.Label(frame, text="توصیف تصویر (فارسی یا انگلیسی):", font=("Tahoma", 9, "bold")).pack(anchor="e", pady=(8, 2))
    prompt_box = scrolledtext.ScrolledText(frame, height=7, wrap="word", font=("Tahoma", 10))
    prompt_box.pack(fill="both", expand=True)
    prompt_box.insert("1.0", "یک تصویر تبلیغاتی حرفه‌ای و واقع‌گرایانه برای یک محصول فناوری، نورپردازی استودیویی، ترکیب‌بندی مینیمال")
    tk.Label(frame, text="Negative prompt:", font=("Tahoma", 9, "bold")).pack(anchor="e", pady=(8, 2))
    negative_box = scrolledtext.ScrolledText(frame, height=3, wrap="word", font=("Tahoma", 9))
    negative_box.pack(fill="x")
    negative_box.insert("1.0", "low quality, blurry, distorted, bad anatomy, extra fingers, watermark, logo, text")

    opts = tk.Frame(frame); opts.pack(fill="x", pady=8)
    tk.Label(opts, text="ابعاد:").pack(side="right", padx=5)
    size_var = tk.StringVar(value="512x512")
    ttk.Combobox(opts, textvariable=size_var, values=["512x512", "512x768", "768x512", "640x640", "768x768"], width=12, state="readonly").pack(side="right")
    tk.Label(opts, text="Steps:").pack(side="right", padx=(14, 4))
    steps_var = tk.IntVar(value=20)
    tk.Spinbox(opts, from_=10, to=35, textvariable=steps_var, width=5).pack(side="right")
    tk.Label(opts, text="CFG:").pack(side="right", padx=(14, 4))
    cfg_var = tk.DoubleVar(value=7.0)
    tk.Spinbox(opts, from_=4.0, to=12.0, increment=0.5, textvariable=cfg_var, width=5).pack(side="right")

    out_frame = tk.Frame(frame); out_frame.pack(fill="x", pady=4)
    tk.Label(out_frame, text="پوشه خروجی:").pack(side="right", padx=6)
    default_out = Path.home() / "Desktop" / "AI_Generated_Images"
    out_var = tk.StringVar(value=str(default_out))
    tk.Entry(out_frame, textvariable=out_var, width=58, justify="left").pack(side="right", padx=4)
    def browse_out():
        from tkinter import filedialog
        chosen = filedialog.askdirectory(initialdir=str(Path(out_var.get()).expanduser()))
        if chosen: out_var.set(chosen)
    ttk.Button(out_frame, text="انتخاب", command=browse_out).pack(side="right", padx=4)

    progress_var = tk.StringVar(value="برای شروع، ComfyUI را اجرا و اتصال را بررسی کن.")
    tk.Label(frame, textvariable=progress_var, anchor="e", wraplength=790, fg="#7030A0").pack(fill="x", pady=6)
    generate_btn = ttk.Button(frame, text="تولید تصویر")
    generate_btn.pack(pady=8)

    def set_status(text):
        try: win.after(0, lambda: progress_var.set(text))
        except Exception: pass

    def work(config):
        monitor = ResourceMonitor()
        monitor.start()
        generation_started = time.perf_counter()
        qwen_elapsed = 0.0
        prompt_id = ""
        final_prompt = config["prompt"]
        initial_prompt = final_prompt
        saved = []
        api_result = None
        try:
            base = config["base"]
            checkpoint = config["checkpoint"]
            if not checkpoint:
                raise ValueError("ابتدا به ComfyUI وصل شو و یک SD 1.5 checkpoint انتخاب کن.")
            width, height = config["width"], config["height"]
            steps = config["steps"]
            cfg = config["cfg"]
            negative = config["negative"]
            if config["improve"]:
                model_name = config["model"]
                if not model_name:
                    raise ValueError("مدل Qwen نصب‌شده انتخاب نشده است.")
                monitor.set_phase("بهبود پرامپت با Qwen")
                set_status(f"Qwen در حال آماده‌سازی پرامپت با {model_name} است...")
                final_prompt, api_result, qwen_elapsed = improve_image_prompt(model_name, initial_prompt)
            monitor.set_phase("تولید تصویر ComfyUI")
            set_status("درخواست به ComfyUI ارسال شد؛ تولید تصویر آغاز می‌شود...")
            workflow = build_comfy_workflow(checkpoint, final_prompt, negative, width, height, steps, cfg)
            prompt_id, images = submit_comfy_image(base, workflow, set_status)
            set_status("تصویر آماده شد؛ در حال دریافت فایل...")
            saved = download_comfy_images(base, images, config["output_dir"])
            generation_finished = time.perf_counter()
            monitor.stop()
            report = save_image_monitor_report(monitor, config["output_dir"], config["model"], initial_prompt,
                                               final_prompt, prompt_id, qwen_elapsed, generation_started,
                                               generation_finished, saved, api_result)
            set_status("پایان موفق؛ تصاویر و گزارش مانیتورینگ ذخیره شدند.")
            def done():
                generate_btn.configure(state="normal")
                messagebox.showinfo("تولید تصویر کامل شد",
                    "فایل‌های تصویر ذخیره شدند:\n" + "\n".join(str(p) for p in saved) +
                    f"\n\nگزارش مصرف منابع:\n{report}", parent=win)
                try: os.startfile(str(config["output_dir"]))
                except Exception: pass
            win.after(0, done)
        except Exception as exc:
            monitor.stop()
            detail = traceback.format_exc()
            try:
                log_path = Path(config["output_dir"] or Path.home() / "Desktop") / f"ComfyUI_Image_Error_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
                log_path.parent.mkdir(parents=True, exist_ok=True)
                log_path.write_text(detail, encoding="utf-8")
            except Exception:
                log_path = Path.cwd() / "ComfyUI_Image_Error.log"
                try: log_path.write_text(detail, encoding="utf-8")
                except Exception: pass
            error_type = type(exc).__name__
            error_text = str(exc)
            def failed(error_type=error_type, error_text=error_text, log_path=log_path):
                generate_btn.configure(state="normal")
                progress_var.set("خطا: " + error_text[:240])
                messagebox.showerror("تولید تصویر ناموفق بود", f"{error_type}: {error_text}\n\nگزارش خطا:\n{log_path}", parent=win)
            win.after(0, failed)

    def start_generation():
        if not prompt_box.get("1.0", "end").strip():
            messagebox.showwarning("پرامپت خالی است", "توصیف تصویر را وارد کن.", parent=win)
            return
        try:
            config = {
                "base": base_var.get().strip().rstrip("/"),
                "checkpoint": checkpoint_var.get().strip(),
                "model": model_var.get().strip(),
                "improve": bool(improve_var.get()),
                "prompt": prompt_box.get("1.0", "end").strip(),
                "negative": negative_box.get("1.0", "end").strip(),
                "width": int(size_var.get().split("x")[0]),
                "height": int(size_var.get().split("x")[1]),
                "steps": int(steps_var.get()),
                "cfg": float(cfg_var.get()),
                "output_dir": out_var.get().strip()
            }
        except Exception as exc:
            messagebox.showerror("تنظیمات نامعتبر", str(exc), parent=win)
            return
        generate_btn.configure(state="disabled")
        threading.Thread(target=work, args=(config,), daemon=True).start()

    generate_btn.configure(command=start_generation)
    ttk.Button(frame, text="بستن", command=win.destroy).pack(pady=(0, 5))
    win.protocol("WM_DELETE_WINDOW", win.destroy)
    root.wait_window(win)


def main():
    global MODEL
    root = Tk()
    root.withdraw()
    try:
        task = choose_task_dialog(root)
        if task == "image":
            launch_image_studio(root)
            return
        if task != "excel":
            return

        source = filedialog.askopenfilename(
            title="فایل اکسل را برای آزمایش انتخاب کن",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not source:
            return

        try:
            models = get_installed_qwen_models()
        except Exception as exc:
            messagebox.showerror("اتصال به Ollama برقرار نشد",
                                 f"Ollama در دسترس نیست. برنامه Ollama را باز کن.\n\n{exc}")
            return
        if not models:
            messagebox.showerror("مدل Qwen3.5 پیدا نشد",
                                 "هیچ مدل Qwen3.5 نصب‌شده‌ای در Ollama پیدا نشد.\nبرای نمونه qwen3.5:9b یا qwen3.5:27b را نصب کن.")
            return
        MODEL = choose_model_dialog(root, models)
        if not MODEL:
            return
        ollama_ready()

        monitor = ResourceMonitor()
        timing = {}
        started_at = datetime.now()
        ollama_ps_before = ollama_ps_snapshot()
        try:
            monitor.start()
            monitor.set_phase("خواندن فایل")
            print("[1/6] خواندن و استخراج اطلاعات Excel...", flush=True)
            phase = time.perf_counter()
            wb, source_ws, facts = extract_workbook(source)
            timing["read_extract"] = time.perf_counter() - phase
            print(f"      {len(facts['loans'])} وام و {len(facts['commitments'])} تعهد استخراج شد.", flush=True)

            monitor.set_phase("محاسبات قطعی Python")
            print("[2/6] محاسبات مستقل و حسابرسی اعداد...", flush=True)
            print(f"      اصل وام‌ها: {format_money(facts['computed_principal'])}", flush=True)
            print(f"      اقساط وام‌ها: {format_money(facts['computed_installments'])}", flush=True)
            print(f"      تعهدات ماهانه: {format_money(facts['total_monthly_obligations'])}", flush=True)

            monitor.set_phase("استنتاج مدل")
            print(f"[3/6] ارسال اطلاعات تأییدشده به {MODEL} ...", flush=True)
            phase = time.perf_counter()
            ai, api_result, client_elapsed = ask_model(facts)
            timing["model_client"] = time.perf_counter() - phase
            ollama_ps_after_model = ollama_ps_snapshot()
            print(f"      پاسخ دریافت شد؛ زمان API: {client_elapsed:.2f} ثانیه", flush=True)

            monitor.set_phase("ساخت داشبورد")
            print("[4/6] ساخت داشبورد، تطبیق و تحلیل کیفی...", flush=True)
            phase = time.perf_counter()
            add_excel_output(wb, source_ws, facts, ai)
            timing["workbook_build"] = time.perf_counter() - phase

            monitor.set_phase("ذخیره فایل اصلی خروجی")
            output = Path(source).with_name(Path(source).stem + "_AI_V4_2_Monitored.xlsx")
            print("[5/6] ذخیره خروجی اولیه...", flush=True)
            phase = time.perf_counter()
            wb.calculation.fullCalcOnLoad = True
            wb.calculation.forceFullCalc = True
            wb.save(output)
            timing["first_save"] = time.perf_counter() - phase
            core_elapsed = monitor.stop()
            finished_at = datetime.now()

            print("[6/6] ساخت گزارش کامل مانیتورینگ و نمودارهای مصرف منابع...", flush=True)
            report_start = time.perf_counter()
            add_monitor_sheets(wb, monitor, timing, api_result, client_elapsed,
                               ollama_ps_after_model, started_at, core_elapsed, MODEL,
                               ollama_ps_before=ollama_ps_before)
            wb.save(output)
            timing["monitor_report_export"] = time.perf_counter() - report_start
            report_ws = wb["گزارش عملکرد"]
            for r in range(1, report_ws.max_row + 1):
                if report_ws.cell(r, 1).value == "زمان ساخت و ذخیره گزارش پایش":
                    report_ws.cell(r, 2, timing["monitor_report_export"])
                    break
            wb.save(output)

            print("\n========== نتیجه حسابرسی ==========", flush=True)
            for audit in facts["audits"]:
                print(f"{audit['title']}: {audit['status']}", flush=True)
            for item in facts["reconciliation"]:
                print(f"{item['institution']}: {item['status']}", flush=True)
            print(f"\nمدل: {MODEL}", flush=True)
            print(f"زمان پردازش اصلی: {core_elapsed:.2f} ثانیه", flush=True)
            print(f"زمان درخواست مدل: {client_elapsed:.2f} ثانیه", flush=True)
            print(f"تعداد نمونه‌های CPU/RAM: {len(monitor.samples)}", flush=True)
            print(f"تعداد نمونه‌های GPU موفق: {sum(1 for x in monitor.gpu_samples if x.get('available'))}", flush=True)
            print(f"فایل خروجی: {output}", flush=True)

            messagebox.showinfo(
                "آزمایش V4.2 کامل شد",
                f"فایل خروجی ساخته شد:\n{output}\n\n"
                f"مدل: {MODEL}\n"
                f"زمان کل پردازش: {core_elapsed:.1f} ثانیه\n"
                f"زمان پاسخ مدل: {client_elapsed:.1f} ثانیه\n"
                f"نمونه‌های CPU/RAM: {len(monitor.samples)}\n"
                f"نمونه‌های GPU موفق: {sum(1 for x in monitor.gpu_samples if x.get('available'))}\n\n"
                "برگه‌ها: داشبورد، Sheet1، گزارش عملکرد، نمونه‌های پایش، نمونه‌های GPU، تفکیک آداپتورها، جزئیات پردازش‌ها"
            )
            try:
                os.startfile(str(output))
            except Exception:
                pass
        except Exception:
            if monitor.start_perf:
                monitor.stop()
            raise
    except urllib.error.URLError as exc:
        messagebox.showerror("خطا در ارتباط با Ollama", f"ارتباط با Ollama برقرار نشد:\n{exc}")
    except Exception as exc:
        detail = traceback.format_exc()
        log_path = Path.home() / "Desktop" / "Ollama_Office_Assistant_V4_2_error.log"
        try:
            log_path.write_text(detail, encoding="utf-8")
        except Exception:
            log_path = Path.cwd() / "Ollama_Office_Assistant_V4_2_error.log"
            try:
                log_path.write_text(detail, encoding="utf-8")
            except Exception:
                pass
        print(detail, flush=True)
        messagebox.showerror(
            "خطا در اجرای V4.2",
            f"{type(exc).__name__}: {exc}\n\nگزارش کامل خطا:\n{log_path}\n\n"
            "این گزارش را نگه دار تا خطا دقیق بررسی شود."
        )
    finally:
        try:
            root.destroy()
        except Exception:
            pass


if __name__ == "__main__":
    main()
