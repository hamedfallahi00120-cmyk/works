"""Optional Tkinter front-end.

All imports of tkinter are local so the engine and its tests run headless.  If
tkinter is missing the launcher explains it instead of crashing.

Threading contract: worker threads only post :class:`UiEvent` objects through
:class:`UiBridge`.  Widgets are read and written exclusively by ``pump_events``
on the Tk main loop.
"""

from __future__ import annotations

import threading
from pathlib import Path
from typing import Any

from local_office_ai.core.settings import Settings
from local_office_ai.ui.event_bridge import UiBridge

TASK_LABELS = {
    "تحلیل عمومی داده": "general_analysis",
    "تحلیل مدیریتی": "management_analysis",
    "کیفیت داده": "data_quality",
    "تحلیل روند": "trend_analysis",
    "مقایسه گروه‌ها": "group_comparison",
    "همبستگی": "correlation_analysis",
    "توضیح مغایرت‌ها": "explain_discrepancies",
    "معنای ستون‌ها": "column_meaning",
    "خلاصه اجرایی": "executive_summary",
}

NO_MODEL = "بدون مدل (قواعد قطعی)"


def launch(settings: Settings) -> int:
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, ttk
    except ImportError as exc:  # pragma: no cover - depends on the host OS packages
        print(
            "رابط گرافیکی در دسترس نیست (tkinter نصب نیست).\n"
            f"جزئیات: {exc}\n"
            "از حالت خط فرمان استفاده کنید:  python run_engine.py --file <path>",
        )
        return 1

    from local_office_ai.ai.ollama_client import OllamaClient
    from local_office_ai.app.coordinator import Coordinator

    root = tk.Tk()
    root.title("Local Office AI | تحلیل هوشمند Excel")
    root.geometry("900x720")
    root.minsize(820, 640)

    bridge = UiBridge()
    state: dict[str, Any] = {
        "worker": None,
        "output": None,
        "coordinator": None,
        "analysis": None,
        "image_status": None,
        "folder_abort": None,
    }

    header = tk.Label(
        root,
        text=(
            "موتور تحلیل داده و آمار Excel\n"
            "درخواست را فارسی بنویسید؛ ساختار، محاسبه و گزارش با Python انجام می‌شود"
        ),
        font=("Tahoma", 13, "bold"),
        fg="#17365D",
        justify="center",
    )
    header.pack(pady=(14, 8))

    form = tk.Frame(root)
    form.pack(fill="x", padx=18)

    tk.Label(form, text="فایل Excel:", font=("Tahoma", 10)).grid(row=0, column=0, sticky="e", pady=4)
    file_var = tk.StringVar()
    tk.Entry(form, textvariable=file_var, width=64, justify="left").grid(row=0, column=1, padx=6, pady=4)

    def browse() -> None:
        path = filedialog.askopenfilename(
            title="فایل Excel را انتخاب کنید",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("همه فایل‌ها", "*.*")],
        )
        if path:
            file_var.set(path)

    tk.Button(form, text="انتخاب…", command=browse, width=10).grid(row=0, column=2, padx=4)

    tk.Label(form, text="مدل Ollama:", font=("Tahoma", 10)).grid(row=1, column=0, sticky="e", pady=4)
    model_var = tk.StringVar(value=NO_MODEL)
    model_box = ttk.Combobox(form, textvariable=model_var, width=40, state="readonly")
    model_box["values"] = [NO_MODEL]
    model_box.grid(row=1, column=1, padx=6, pady=4, sticky="w")

    def refresh_models() -> None:
        bridge.log("در حال خواندن فهرست مدل‌های Ollama…")

        def work() -> None:
            try:
                names = [model.name for model in OllamaClient(settings).preferred_models()]
            except Exception as exc:  # noqa: BLE001 - surfaced in the log pane
                bridge.log(f"فهرست مدل‌ها خوانده نشد: {exc}")
                bridge.models([])
                return
            bridge.models(names)

        threading.Thread(target=work, daemon=True).start()

    tk.Button(form, text="بازخوانی مدل‌ها", command=refresh_models, width=14).grid(row=1, column=2, padx=4)

    tk.Label(form, text="نوع تحلیل:", font=("Tahoma", 10)).grid(row=2, column=0, sticky="e", pady=4)
    task_var = tk.StringVar(value="تحلیل عمومی داده")
    ttk.Combobox(form, textvariable=task_var, values=list(TASK_LABELS), state="readonly", width=40).grid(
        row=2, column=1, padx=6, pady=4, sticky="w"
    )

    tk.Label(form, text="درخواست (فارسی):", font=("Tahoma", 10)).grid(row=3, column=0, sticky="e", pady=4)
    request_var = tk.StringVar()
    tk.Entry(form, textvariable=request_var, width=64, justify="right").grid(row=3, column=1, padx=6, pady=4)
    tk.Label(
        form,
        text="مثال: تحلیل ماهانه شاخص‌ها، مقایسه 1404 با 1403 و بیشترین و کمترین ماه با گزارش مدیریتی",
        font=("Tahoma", 8),
        fg="#6B7280",
        anchor="e",
        justify="right",
    ).grid(row=4, column=1, sticky="e", padx=6)

    tk.Label(form, text="هدف تحلیل (اختیاری):", font=("Tahoma", 10)).grid(row=5, column=0, sticky="e", pady=4)
    goal_var = tk.StringVar()
    tk.Entry(form, textvariable=goal_var, width=64, justify="right").grid(row=5, column=1, padx=6, pady=4)

    tk.Label(form, text="پرسش (اختیاری):", font=("Tahoma", 10)).grid(row=6, column=0, sticky="e", pady=4)
    question_var = tk.StringVar()
    tk.Entry(form, textvariable=question_var, width=64, justify="right").grid(row=6, column=1, padx=6, pady=4)

    actions = tk.Frame(root)
    actions.pack(pady=10)
    run_button = tk.Button(actions, text="شروع پردازش", width=18, bg="#17365D", fg="white", font=("Tahoma", 10, "bold"))
    run_button.pack(side="left", padx=6)
    stop_button = tk.Button(actions, text="توقف", width=10, state="disabled")
    stop_button.pack(side="left", padx=6)
    open_button = tk.Button(actions, text="بازکردن خروجی", width=14, state="disabled")
    open_button.pack(side="left", padx=6)
    folder_button = tk.Button(actions, text="تحلیل پوشه", width=14)
    folder_button.pack(side="left", padx=6)
    image_button = tk.Button(actions, text="استودیو تصویر", width=14)
    image_button.pack(side="left", padx=6)

    progress = ttk.Progressbar(root, mode="indeterminate", length=820)
    progress.pack(pady=4, padx=18, fill="x")

    log_box = tk.Text(root, height=22, font=("Consolas", 9), bg="#0F172A", fg="#E2E8F0")
    log_box.pack(fill="both", expand=True, padx=18, pady=(6, 14))

    def append_log(message: str) -> None:
        log_box.insert("end", message + "\n")
        log_box.see("end")

    def finish_run() -> None:
        progress.stop()
        run_button.configure(state="normal")
        stop_button.configure(state="disabled")

    def confirm_structure(analysis) -> bool:
        if analysis.primary_table is None or analysis.structure_confidence >= 0.6:
            return True
        lines = [
            f"{table.sheet}!{table.address} — {table.kind.value} — اطمینان {table.confidence:.0%}"
            for table in analysis.tables[:8]
        ]
        text = "ساختار فایل با اطمینان پایین شناسایی شد:\n\n" + "\n".join(lines)
        text += f"\n\nجدول انتخاب‌شده: {analysis.primary_table.sheet}!{analysis.primary_table.address}\nادامه می‌دهید؟"
        return bridge.confirm(text)

    def worker(
        file_path: str,
        model: str,
        task_label: str,
        question: str,
        goal: str,
        request_text: str,
    ) -> None:
        try:
            coordinator = Coordinator(settings, progress=bridge.log, console=False)
            state["coordinator"] = coordinator
            state["folder_abort"] = None
            use_ai = model != NO_MODEL
            result = coordinator.run(
                file_path,
                model=model if use_ai else None,
                use_ai=use_ai,
                task=TASK_LABELS.get(task_label, "general_analysis"),
                question=question or None,
                goal=goal or None,
                request=request_text or None,
                confirm_structure=confirm_structure,
            )
            bridge.done(result)
        except Exception as exc:  # noqa: BLE001 - show any failure to the operator
            bridge.error(f"{type(exc).__name__}: {exc}")
        finally:
            bridge.log("پردازش پایان یافت.")

    def start() -> None:
        file_path = file_var.get().strip()
        if not file_path:
            messagebox.showwarning("فایل انتخاب نشده", "ابتدا فایل Excel را انتخاب کنید.")
            return
        if not Path(file_path).exists():
            messagebox.showerror("فایل پیدا نشد", file_path)
            return
        model = model_var.get()
        task_label = task_var.get()
        question = question_var.get().strip()
        goal = goal_var.get().strip()
        request_text = request_var.get().strip()
        log_box.delete("1.0", "end")
        run_button.configure(state="disabled")
        stop_button.configure(state="normal")
        progress.start(12)
        thread = threading.Thread(
            target=worker,
            args=(file_path, model, task_label, question, goal, request_text),
            daemon=True,
        )
        state["worker"] = thread
        thread.start()

    def folder_worker(folder_path: str, model: str, request_text: str, abort_event) -> None:
        try:
            from local_office_ai.app.folder_workflow import FolderWorkflow

            use_ai = model != NO_MODEL
            workflow = FolderWorkflow(settings, progress=bridge.log, abort_event=abort_event)
            result = workflow.run(
                folder_path,
                request=request_text or None,
                model=model if use_ai else None,
                use_ai=use_ai,
            )
            for item in result.files:
                detail = f"  - {item.name}: {item.status} | شاخص {item.metrics}"
                if item.error:
                    detail += f" | خطا: {item.error}"
                bridge.log(detail)
            for warning in result.warnings[:10]:
                bridge.log(f"هشدار: {warning}")
            bridge.done(result)
        except Exception as exc:  # noqa: BLE001 - surfaced in the log pane
            bridge.error(f"{type(exc).__name__}: {exc}")
        finally:
            bridge.log("پردازش پوشه پایان یافت.")

    def analyse_folder() -> None:
        folder = filedialog.askdirectory(title="پوشه حاوی فایل‌های Excel را انتخاب کنید")
        if not folder:
            return
        model = model_var.get()
        request_text = request_var.get().strip()
        abort_event = threading.Event()
        state["folder_abort"] = abort_event
        state["coordinator"] = None
        log_box.delete("1.0", "end")
        run_button.configure(state="disabled")
        stop_button.configure(state="normal")
        progress.start(12)
        thread = threading.Thread(
            target=folder_worker,
            args=(folder, model, request_text, abort_event),
            daemon=True,
        )
        state["worker"] = thread
        thread.start()

    def stop() -> None:
        stopped = False
        coordinator = state.get("coordinator")
        if coordinator is not None:
            coordinator.abort()
            stopped = True
        folder_abort = state.get("folder_abort")
        if folder_abort is not None:
            folder_abort.set()
            stopped = True
        if stopped:
            bridge.log("درخواست توقف ثبت شد؛ پردازش در پایان مرحله جاری متوقف می‌شود.")

    def open_output() -> None:
        output = state.get("output")
        if not output:
            return
        try:
            import os

            os.startfile(output)  # type: ignore[attr-defined]
        except AttributeError:
            import subprocess
            import sys

            subprocess.Popen(["xdg-open", output] if sys.platform.startswith("linux") else ["open", output])
        except OSError as exc:
            messagebox.showerror("بازکردن فایل", str(exc))

    def open_image_studio() -> None:
        """Minimal ComfyUI studio window (optional feature, preserved from V4.2)."""
        from local_office_ai.ai.image_studio import ComfyUIClient, ImageRequest, ImageStudio

        window = tk.Toplevel(root)
        window.title("استودیو تصویر ComfyUI")
        window.geometry("620x460")

        tk.Label(window, text="پرامپت تصویر (خالی = پرامپت خودکار از داده‌ها)", font=("Tahoma", 10)).pack(pady=(10, 4))
        prompt_box = tk.Text(window, height=6, font=("Tahoma", 10), wrap="word")
        prompt_box.pack(fill="x", padx=12)

        controls = tk.Frame(window)
        controls.pack(fill="x", padx=12, pady=8)
        tk.Label(controls, text="Checkpoint:", font=("Tahoma", 10)).pack(side="right")
        checkpoint_var = tk.StringVar()
        checkpoint_box = ttk.Combobox(controls, textvariable=checkpoint_var, width=28, state="readonly")
        checkpoint_box.pack(side="right", padx=6)

        status = tk.Label(window, text="", font=("Tahoma", 9), fg="#17365D", anchor="e", justify="right")
        status.pack(fill="x", padx=12)
        state["image_status"] = status
        state["image_checkpoint_box"] = checkpoint_box
        state["image_checkpoint_var"] = checkpoint_var

        def load_checkpoints() -> None:
            status.configure(text="در حال خواندن checkpoint ها…")

            def work() -> None:
                try:
                    names = ComfyUIClient(settings).checkpoints()
                except Exception as exc:  # noqa: BLE001
                    bridge.image_status(f"ComfyUI در دسترس نیست: {exc}")
                    return
                bridge.checkpoints(names)

            threading.Thread(target=work, daemon=True).start()

        tk.Button(controls, text="بازخوانی", command=load_checkpoints).pack(side="right", padx=4)

        def generate() -> None:
            prompt_text = prompt_box.get("1.0", "end").strip()
            selected_model = model_var.get()
            status.configure(text="در حال تولید تصویر…")
            request = ImageRequest(prompt="", checkpoint=checkpoint_var.get() or None)

            def work() -> None:
                try:
                    result = ImageStudio(settings).generate(
                        prompt_text,
                        analysis=state.get("analysis"),
                        model=None if selected_model == NO_MODEL else selected_model,
                        use_ai=selected_model != NO_MODEL,
                        request=request,
                    )
                    bridge.image_status("تصویر ذخیره شد:\n" + "\n".join(str(path) for path in result.images))
                except Exception as exc:  # noqa: BLE001
                    bridge.image_status(f"خطا: {exc}")

            threading.Thread(target=work, daemon=True).start()

        def on_close() -> None:
            state["image_status"] = None
            state["image_checkpoint_box"] = None
            state["image_checkpoint_var"] = None
            window.destroy()

        window.protocol("WM_DELETE_WINDOW", on_close)
        tk.Button(window, text="تولید تصویر", command=generate, bg="#17365D", fg="white", font=("Tahoma", 10, "bold")).pack(pady=8)
        load_checkpoints()

    def pump_events() -> None:
        for event in bridge.drain():
            if event.kind == "log":
                append_log(event.payload)
            elif event.kind == "models":
                names = event.payload
                model_box["values"] = [NO_MODEL] + names
                if names and model_var.get() not in names:
                    model_var.set(names[0])
                append_log(
                    f"{len(names)} مدل Ollama پیدا شد."
                    if names
                    else "مدلی در Ollama پیدا نشد؛ تحلیل با قواعد قطعی انجام می‌شود."
                )
            elif event.kind == "checkpoints":
                names = event.payload
                box = state.get("image_checkpoint_box")
                var = state.get("image_checkpoint_var")
                label = state.get("image_status")
                if box is not None:
                    box["values"] = names
                if var is not None and names:
                    var.set(names[0])
                if label is not None:
                    label.configure(text=f"{len(names)} checkpoint پیدا شد.")
            elif event.kind == "image_status":
                label = state.get("image_status")
                if label is not None:
                    label.configure(text=event.payload)
                append_log(event.payload)
            elif event.kind == "confirm":
                answer = messagebox.askyesno("تأیید ساختار", event.payload)
                bridge.resolve(event, answer)
            elif event.kind == "done":
                result = event.payload
                output_path = getattr(result, "output_path", "")
                state["output"] = output_path
                analysis = getattr(result, "analysis", None)
                if analysis is not None:
                    state["analysis"] = analysis
                append_log(f"پایان: {output_path}")
                finish_run()
                open_button.configure(state="normal")
                messagebox.showinfo("پایان", f"خروجی ساخته شد:\n{output_path}")
            elif event.kind == "error":
                append_log(f"خطا: {event.payload}")
                finish_run()
                messagebox.showerror("خطا", event.payload)
        root.after(100, pump_events)

    image_button.configure(command=open_image_studio)
    folder_button.configure(command=analyse_folder)
    run_button.configure(command=start)
    stop_button.configure(command=stop)
    open_button.configure(command=open_output)

    refresh_models()
    pump_events()
    bridge.log("آماده. فایل را انتخاب کنید و پردازش را شروع کنید.")
    root.mainloop()
    return 0
