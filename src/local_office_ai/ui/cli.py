"""Command line interface.

The CLI is the primary entry point; the GUI is an optional layer on top of it.
It works headless, which is what makes the engine testable.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from local_office_ai.app.coordinator import Coordinator
from local_office_ai.ai.ollama_client import OllamaClient
from local_office_ai.core.exceptions import AIServiceError, EngineError, ProcessingAborted
from local_office_ai.core.models import AnalysisResult, TableKind
from local_office_ai.core.settings import load_settings
from local_office_ai.app.workflow import ExcelIntelligenceWorkflow

KIND_FA = {
    TableKind.LOAN: "جدول وام",
    TableKind.OBLIGATION: "جدول تعهدات",
    TableKind.TIME_SERIES: "سری زمانی",
    TableKind.KPI: "شاخص‌ها",
    TableKind.GENERIC: "عمومی",
}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="local-office-ai",
        description="موتور تحلیل داده، آمار توصیفی و گزارش‌گیری Excel با مدل محلی Ollama",
    )
    parser.add_argument("--file", "-f", help="مسیر فایل Excel ورودی")
    parser.add_argument(
        "--folder",
        help="تحلیل همه فایل‌های Excel یک پوشه و مقایسه بین آن‌ها",
    )
    parser.add_argument(
        "--request",
        "-r",
        help="درخواست کاربر به زبان فارسی (نوع تحلیل و شاخص‌های هدف از همین متن خوانده می‌شود)",
    )
    parser.add_argument(
        "--audit",
        action="store_true",
        help="افزودن برگه‌های فنی و ردیابی منبع به خروجی (حالت توسعه/حسابرسی)",
    )
    parser.add_argument("--model", "-m", help="نام مدل Ollama (پیش‌فرض: خودکار)")
    parser.add_argument("--no-ai", action="store_true", help="تحلیل مدل را غیرفعال کن")
    parser.add_argument(
        "--task",
        default="general_analysis",
        choices=[
            "general_analysis", "management_analysis", "data_quality", "trend_analysis",
            "group_comparison", "correlation_analysis", "explain_discrepancies",
            "column_meaning", "executive_summary",
        ],
        help="نوع تحلیلی که از مدل خواسته می‌شود",
    )
    parser.add_argument("--question", "-q", help="پرسش کاربر درباره فایل")
    parser.add_argument("--goal", help="هدف تحلیل به زبان طبیعی؛ مکمل یا جایگزین پرسش")
    parser.add_argument("--output", "-o", help="مسیر فایل خروجی")
    parser.add_argument("--json", dest="json_path", help="ذخیره خلاصه ماشین‌خوان در این مسیر")
    parser.add_argument("--list-models", action="store_true", help="فهرست مدل‌های نصب‌شده Ollama")
    parser.add_argument("--list-checkpoints", action="store_true", help="فهرست checkpoint های ComfyUI")
    parser.add_argument(
        "--image",
        nargs="?",
        const="",
        help="تولید تصویر با ComfyUI پس از تحلیل (بدون مقدار: پرامپت خودکار از داده‌ها)",
    )
    parser.add_argument("--image-size", help="اندازه تصویر، مثال 512x512")
    parser.add_argument("--image-steps", type=int, help="تعداد گام نمونه‌گیری")
    parser.add_argument("--checkpoint", help="نام checkpoint در ComfyUI")
    parser.add_argument("--no-image-ai", action="store_true", help="بهبود پرامپت تصویر با مدل انجام نشود")
    parser.add_argument("--check-structure", action="store_true", help="فقط ساختار فایل را گزارش کن")
    parser.add_argument("--yes", "-y", action="store_true", help="تأیید خودکار ساختار مشکوک")
    parser.add_argument("--gui", action="store_true", help="اجرای رابط گرافیکی")
    parser.add_argument("--root", help="ریشه پروژه (پیش‌فرض: محل همین بسته)")
    parser.add_argument("--quiet", action="store_true", help="خروجی کنسول کم‌تر")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    settings = load_settings(args.root)

    if args.gui:
        from local_office_ai.ui.wizard import launch

        return launch(settings)

    if args.list_models:
        return _list_models(settings)
    if args.list_checkpoints:
        return _list_checkpoints(settings)

    profile = "audit" if args.audit else (settings.report_profile or "clean")

    if args.folder:
        return _run_folder(settings, args, profile)

    if not args.file:
        print("خطا: مسیر فایل ورودی لازم است. مثال: --file \"C:\\data\\book.xlsx\"", file=sys.stderr)
        return 2
    source = Path(args.file)
    if not source.exists():
        print(f"خطا: فایل پیدا نشد: {source}", file=sys.stderr)
        return 2

    if args.check_structure:
        return _check_structure(settings, source, request=args.request)

    coordinator = Coordinator(
        settings,
        progress=(lambda message: None) if args.quiet else None,
        console=not args.quiet,
    )
    confirm = _confirmation(settings, args.yes)
    request = args.request or args.goal or args.question
    try:
        result = coordinator.run(
            source,
            model=args.model,
            use_ai=not args.no_ai,
            task=args.task,
            question=args.question,
            goal=args.goal,
            request=request,
            output=args.output,
            profile=profile,
            confirm_structure=confirm,
        )
    except ProcessingAborted as exc:
        print(f"پردازش متوقف شد: {exc}", file=sys.stderr)
        return 2
    except EngineError as exc:
        print(f"خطا: {exc}", file=sys.stderr)
        print(f"لاگ کامل: {coordinator.log_path}", file=sys.stderr)
        return 1
    except Exception as exc:  # noqa: BLE001 - last line of defence for the CLI
        print(f"خطای غیرمنتظره: {type(exc).__name__}: {exc}", file=sys.stderr)
        print(f"لاگ کامل: {coordinator.log_path}", file=sys.stderr)
        return 1

    if args.json_path:
        path = coordinator.export_json(result, args.json_path)
        print(f"خلاصه JSON: {path}")
    if args.image is not None:
        return _generate_image(settings, args, result)
    return 0


def _run_folder(settings, args, profile: str) -> int:
    from local_office_ai.app.folder_workflow import FolderWorkflow

    folder = Path(args.folder)
    if not folder.is_dir():
        print(f"خطا: پوشه پیدا نشد: {folder}", file=sys.stderr)
        return 2
    workflow = FolderWorkflow(
        settings,
        progress=(lambda message: None) if args.quiet else print,
    )
    try:
        result = workflow.run(
            folder,
            request=args.request or args.goal or args.question,
            model=args.model,
            use_ai=not args.no_ai,
            output=args.output,
            profile=profile,
        )
    except EngineError as exc:
        print(f"خطا: {exc}", file=sys.stderr)
        return 1
    print(f"تعداد فایل: {len(result.files)}")
    for item in result.files:
        print(f"  - {item.name}: {item.status} | شاخص {item.metrics} | ماه‌ها {item.months or '—'}")
        if item.error:
            print(f"      خطا: {item.error}")
    if result.output_path:
        print(f"خروجی: {result.output_path}")
    for warning in result.warnings[:10]:
        print(f"هشدار: {warning}", file=sys.stderr)
    return 0


def _list_models(settings) -> int:
    client = OllamaClient(settings)
    try:
        models = client.list_models()
    except AIServiceError as exc:
        print(f"اتصال به Ollama برقرار نشد: {exc}", file=sys.stderr)
        return 1
    if not models:
        print("هیچ مدلی در Ollama نصب نیست.")
        return 0
    print("مدل‌های نصب‌شده Ollama:")
    for model in models:
        marker = "*" if any(model.name.lower().startswith(prefix) for prefix in settings.model_name_prefixes) else " "
        print(f" {marker} {model.name:32} {model.size_gb:6.2f} GB  {model.parameter_size:>6}  {model.family}")
    return 0


def _list_checkpoints(settings) -> int:
    from local_office_ai.ai.image_studio import ComfyUIClient

    try:
        names = ComfyUIClient(settings).checkpoints()
    except EngineError as exc:
        print(f"ComfyUI در دسترس نیست: {exc}", file=sys.stderr)
        return 1
    if not names:
        print("هیچ checkpoint ای در ComfyUI پیدا نشد.")
        return 0
    for name in names:
        print(name)
    return 0


def _generate_image(settings, args, result) -> int:
    from local_office_ai.ai.image_studio import ImageRequest, ImageStudio

    width, height = settings.comfyui_width, settings.comfyui_height
    if args.image_size:
        try:
            width_text, height_text = args.image_size.lower().split("x", 1)
            width, height = int(width_text), int(height_text)
        except ValueError:
            print("خطا: قالب --image-size باید مانند 512x512 باشد.", file=sys.stderr)
            return 2
    request = ImageRequest(
        prompt="",
        checkpoint=args.checkpoint or None,
        width=width,
        height=height,
        steps=args.image_steps or settings.comfyui_steps,
        cfg=settings.comfyui_cfg,
    )
    studio = ImageStudio(settings)
    try:
        outcome = studio.generate(
            args.image or "",
            analysis=result.analysis,
            model=args.model,
            use_ai=not args.no_image_ai,
            request=request,
            progress=lambda message: None if args.quiet else print(message),
        )
    except EngineError as exc:
        print(f"تولید تصویر انجام نشد: {exc}", file=sys.stderr)
        return 1
    for path in outcome.images:
        print(f"تصویر: {path}")
    for warning in outcome.warnings:
        print(f"هشدار: {warning}", file=sys.stderr)
    return 0


def _check_structure(settings, source: Path, request: str | None = None) -> int:
    workflow = ExcelIntelligenceWorkflow(settings)
    try:
        analysis = workflow.analyze_only(source, request=request)
    except EngineError as exc:
        print(f"خطا: {exc}", file=sys.stderr)
        return 1
    print(f"فایل: {source}")
    print(f"برگه‌ها: {len(analysis.meta.sheets)}")
    for sheet in analysis.meta.sheets:
        print(f"  - {sheet.title}: داده {sheet.used_dimensions} | فرمت {sheet.declared_dimensions} | فرمول {sheet.formula_count}")
    print("جدول‌های شناسایی‌شده:")
    for table in analysis.tables:
        print(
            f"  - [{KIND_FA.get(table.kind, table.kind.value)}] {table.sheet}!{table.address} "
            f"ردیف={table.row_count} اطمینان={table.confidence:.0%}"
        )
        for column in table.columns:
            print(f"      {column.letter}: {column.header or '(بدون عنوان)'} → {column.role.value} ({column.data_type.value})")
    if analysis.primary_table is None:
        print("جدول اصلی: پیدا نشد")
        return 0
    print(f"جدول اصلی: {analysis.primary_table.sheet}!{analysis.primary_table.address}")
    integration = getattr(analysis, "integration", None)
    if integration is not None:
        print(
            f"ماتریس ماهانه: {len(integration.months)} ماه ({'، '.join(integration.months)}) | "
            f"دوره‌ها: {'، '.join(integration.periods)} | شاخص‌ها: {len(integration.series)}"
        )
    intent = getattr(analysis, "request_intent", None)
    if intent is not None and intent.text:
        print(f"درخواست: {intent.text}")
        print(f"تفسیر: {'، '.join(intent.scope_labels)}")
    return 0


def _confirmation(settings, auto_yes: bool):
    def confirm(analysis: AnalysisResult) -> bool:
        threshold = 0.6
        if auto_yes or analysis.primary_table is None:
            return True
        if analysis.structure_confidence >= threshold:
            return True
        print("\nساختار فایل با اطمینان پایین شناسایی شد:")
        for table in analysis.tables[:10]:
            print(f"  - {table.sheet}!{table.address} ({table.kind.value}) اطمینان {table.confidence:.0%}")
        print(f"جدول انتخاب‌شده: {analysis.primary_table.sheet}!{analysis.primary_table.address}")
        if not sys.stdin.isatty():
            print("ورودی تعاملی نیست؛ پردازش با جدول انتخاب‌شده ادامه می‌یابد.")
            return True
        answer = input("ادامه پردازش با این ساختار؟ [y/N] ").strip().lower()
        return answer in {"y", "yes", "بله", "ب"}
    return confirm


if __name__ == "__main__":
    raise SystemExit(main())
