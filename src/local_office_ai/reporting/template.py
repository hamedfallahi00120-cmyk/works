"""Inspect an existing DOCX template without using it as a data source.

The inspector records page layout, margins, fonts, paragraph styles, heading
hierarchy, tables, headers/footers, numbering, RTL formatting and placeholder
locations.  It deliberately has no ``values`` collection: a template is a
presentation/structure reference only.
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from local_office_ai.core.exceptions import ReportTemplateError

PLACEHOLDER_PATTERN = re.compile(r"\{\{\s*([^{}\r\n]+?)\s*\}\}")
_NUMERIC_PATTERN = re.compile(r"(?<![\w])[0-9۰-۹٠-٩][0-9۰-۹٠-٩,٬٫./\-]*")
_HEADING_PATTERN = re.compile(r"^(?:heading|عنوان)\s*(\d+)$", re.IGNORECASE)


def _docx_module() -> Any:
    try:
        import docx  # noqa: F401
        from docx import Document
    except ImportError as exc:  # pragma: no cover - optional dependency branch
        raise ReportTemplateError(
            "برای بازرسی قالب Word، کتابخانه python-docx نصب نیست."
        ) from exc
    return Document


def _qn() -> Any:
    from docx.oxml.ns import qn

    return qn


def _twips(value: Any) -> int | None:
    if value is None:
        return None
    try:
        return int(value.twips)
    except (AttributeError, TypeError, ValueError):
        return None


def _points(value: Any) -> float | None:
    if value is None:
        return None
    try:
        return round(float(value.pt), 3)
    except (AttributeError, TypeError, ValueError):
        return None


def _enum_name(value: Any) -> str:
    return getattr(value, "name", None) or str(value) if value is not None else ""


def _font_snapshot(font: Any) -> dict[str, Any]:
    if font is None:
        return {}
    color = ""
    try:
        if font.color is not None and font.color.rgb is not None:
            color = str(font.color.rgb)
    except (AttributeError, ValueError):
        color = ""
    return {
        "name": font.name or "",
        "size_pt": _points(font.size),
        "bold": font.bold,
        "italic": font.italic,
        "underline": font.underline,
        "color": color,
        "rtl": getattr(font, "rtl", None),
    }


@dataclass
class PageLayout:
    index: int
    width_twips: int | None = None
    height_twips: int | None = None
    orientation: str = ""
    top_margin_twips: int | None = None
    bottom_margin_twips: int | None = None
    left_margin_twips: int | None = None
    right_margin_twips: int | None = None
    header_distance_twips: int | None = None
    footer_distance_twips: int | None = None
    gutter_twips: int | None = None
    different_first_page: bool = False
    start_type: str = ""
    bidi: bool = False
    rtl_gutter: bool = False

    def to_dict(self) -> dict[str, Any]:
        return self.__dict__.copy()


@dataclass
class FontUsage:
    name: str = ""
    size_pt: float | None = None
    bold: bool | None = None
    italic: bool | None = None
    rtl: bool | None = None
    count: int = 1
    source: str = ""

    def to_dict(self) -> dict[str, Any]:
        return self.__dict__.copy()


@dataclass
class ParagraphStyleInfo:
    name: str
    style_type: str = "paragraph"
    builtin: bool = False
    base_style: str = ""
    font: dict[str, Any] = field(default_factory=dict)
    alignment: str = ""
    space_before_pt: float | None = None
    space_after_pt: float | None = None
    line_spacing: float | None = None
    keep_with_next: bool | None = None
    keep_together: bool | None = None
    outline_level: int | None = None
    bidi: bool = False

    def to_dict(self) -> dict[str, Any]:
        return self.__dict__.copy()


@dataclass
class ParagraphInfo:
    index: int
    location: str
    text: str
    style: str = ""
    heading_level: int = 0
    alignment: str = ""
    bidi: bool = False
    run_count: int = 0
    rtl_run_count: int = 0
    font_names: tuple[str, ...] = ()
    placeholder_keys: tuple[str, ...] = ()

    def to_dict(self) -> dict[str, Any]:
        data = self.__dict__.copy()
        data["font_names"] = list(self.font_names)
        data["placeholder_keys"] = list(self.placeholder_keys)
        return data


@dataclass
class TableInfo:
    index: int
    location: str
    rows: int = 0
    columns: int = 0
    style: str = ""
    alignment: str = ""
    autofit: bool | None = None
    bidi_visual: bool = False
    header_row: tuple[str, ...] = ()
    cell_count: int = 0
    placeholder_keys: tuple[str, ...] = ()

    def to_dict(self) -> dict[str, Any]:
        data = self.__dict__.copy()
        data["header_row"] = list(self.header_row)
        data["placeholder_keys"] = list(self.placeholder_keys)
        return data


@dataclass
class HeaderFooterInfo:
    kind: str
    variant: str
    section_index: int
    linked_to_previous: bool = False
    text: str = ""
    table_count: int = 0
    placeholder_keys: tuple[str, ...] = ()

    def to_dict(self) -> dict[str, Any]:
        data = self.__dict__.copy()
        data["placeholder_keys"] = list(self.placeholder_keys)
        return data


@dataclass
class NumberingInfo:
    num_id: str
    level: int
    format: str = ""
    text: str = ""
    start: int | None = None
    rtl: bool = False
    paragraph_location: str = ""

    def to_dict(self) -> dict[str, Any]:
        return self.__dict__.copy()


@dataclass
class PlaceholderInfo:
    key: str
    raw: str
    location: str
    paragraph_style: str = ""

    def to_dict(self) -> dict[str, Any]:
        return self.__dict__.copy()


@dataclass
class StaticNumericLiteral:
    text: str
    location: str

    def to_dict(self) -> dict[str, Any]:
        return self.__dict__.copy()


@dataclass
class ReportTemplate:
    """Structural view of a DOCX template; never a source of report figures."""

    path: str
    valid: bool = True
    errors: list[str] = field(default_factory=list)
    section_count: int = 0
    page_layouts: list[PageLayout] = field(default_factory=list)
    styles: list[ParagraphStyleInfo] = field(default_factory=list)
    used_styles: list[str] = field(default_factory=list)
    fonts: list[FontUsage] = field(default_factory=list)
    headings: list[ParagraphInfo] = field(default_factory=list)
    paragraphs: list[ParagraphInfo] = field(default_factory=list)
    tables: list[TableInfo] = field(default_factory=list)
    headers: list[HeaderFooterInfo] = field(default_factory=list)
    footers: list[HeaderFooterInfo] = field(default_factory=list)
    placeholders: list[PlaceholderInfo] = field(default_factory=list)
    numbering: list[NumberingInfo] = field(default_factory=list)
    static_numeric_literals: list[StaticNumericLiteral] = field(default_factory=list)
    rtl: dict[str, Any] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    paragraphs_truncated: bool = False
    tables_truncated: bool = False

    @property
    def placeholder_keys(self) -> tuple[str, ...]:
        return tuple(sorted({item.key for item in self.placeholders}))

    @classmethod
    def inspect(cls, path: Path | str, **kwargs: Any) -> "ReportTemplate":
        """Inspect a DOCX template without treating its figures as data."""

        return inspect_docx_template(path, **kwargs)

    @property
    def is_data_source(self) -> bool:
        """A template is presentation only; historical figures never become data."""

        return False

    @property
    def has_static_figures(self) -> bool:
        return bool(self.static_numeric_literals)

    def placeholder_locations(self, key: str) -> list[str]:
        return [item.location for item in self.placeholders if item.key == key]

    def to_dict(self) -> dict[str, Any]:
        return {
            "path": self.path,
            "valid": self.valid,
            "errors": list(self.errors),
            "section_count": self.section_count,
            "page_layouts": [item.to_dict() for item in self.page_layouts],
            "styles": [item.to_dict() for item in self.styles],
            "used_styles": list(self.used_styles),
            "fonts": [item.to_dict() for item in self.fonts],
            "headings": [item.to_dict() for item in self.headings],
            "paragraphs": [item.to_dict() for item in self.paragraphs],
            "tables": [item.to_dict() for item in self.tables],
            "headers": [item.to_dict() for item in self.headers],
            "footers": [item.to_dict() for item in self.footers],
            "placeholders": [item.to_dict() for item in self.placeholders],
            "placeholder_keys": list(self.placeholder_keys),
            "numbering": [item.to_dict() for item in self.numbering],
            "static_numeric_literals": [item.to_dict() for item in self.static_numeric_literals],
            "rtl": dict(self.rtl),
            "warnings": list(self.warnings),
            "paragraphs_truncated": self.paragraphs_truncated,
            "tables_truncated": self.tables_truncated,
            "is_data_source": self.is_data_source,
        }


def _validate_template_file(path: Path) -> None:
    if not path.exists():
        raise ReportTemplateError(f"قالب Word پیدا نشد: {path}")
    if not path.is_file():
        raise ReportTemplateError(f"مسیر قالب Word یک فایل نیست: {path}")
    if path.suffix.lower() != ".docx":
        raise ReportTemplateError(
            f"قالب پشتیبانی نمی‌شود: {path.suffix or 'بدون پسوند'}؛ قالب باید .docx باشد."
        )
    if path.stat().st_size <= 0:
        raise ReportTemplateError(f"قالب Word خالی است: {path}")
    if not zipfile.is_zipfile(path):
        raise ReportTemplateError(f"قالب Word معتبر نیست (فایل ZIP نیست): {path}")
    try:
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
    except (OSError, zipfile.BadZipFile) as exc:
        raise ReportTemplateError(f"قالب Word قابل خواندن نیست: {path}") from exc
    if "word/document.xml" not in names:
        raise ReportTemplateError(
            f"قالب Word معتبر نیست (word/document.xml ندارد): {path}"
        )


def inspect_docx_template(
    path: Path | str,
    *,
    max_paragraphs: int = 500,
    max_tables: int = 100,
) -> ReportTemplate:
    """Read a DOCX template and return its structure, styles and placeholders."""

    source = Path(path).expanduser().resolve()
    _validate_template_file(source)
    Document = _docx_module()
    try:
        document = Document(str(source))
    except Exception as exc:  # noqa: BLE001 - python-docx raises many parser errors
        raise ReportTemplateError(f"قالب Word قابل خواندن نیست: {source} ({exc})") from exc

    qn = _qn()
    template = ReportTemplate(path=str(source))
    template.section_count = len(document.sections)

    fonts: dict[tuple[Any, ...], FontUsage] = {}
    used_styles: set[str] = set()
    numbering_seen: set[tuple[Any, ...]] = set()
    paragraph_counter = 0

    def record_font(font: Any, source_kind: str) -> None:
        snapshot = _font_snapshot(font)
        if not snapshot.get("name") and snapshot.get("size_pt") is None:
            return
        key = (
            snapshot.get("name"),
            snapshot.get("size_pt"),
            snapshot.get("bold"),
            snapshot.get("italic"),
            snapshot.get("rtl"),
            source_kind,
        )
        if key in fonts:
            fonts[key].count += 1
        else:
            fonts[key] = FontUsage(
                name=snapshot.get("name") or "",
                size_pt=snapshot.get("size_pt"),
                bold=snapshot.get("bold"),
                italic=snapshot.get("italic"),
                rtl=snapshot.get("rtl"),
                count=1,
                source=source_kind,
            )

    def collect_numbering(paragraph: Any, location: str) -> None:
        pPr = paragraph._p.pPr
        numPr = pPr.find(qn("w:numPr")) if pPr is not None else None
        if numPr is None:
            style = getattr(paragraph, "style", None)
            style_pPr = style.element.find(qn("w:pPr")) if style is not None and hasattr(style, "element") else None
            numPr = style_pPr.find(qn("w:numPr")) if style_pPr is not None else None
        if numPr is None:
            return
        num_id = ""
        level = 0
        num_id_node = numPr.find(qn("w:numId"))
        if num_id_node is not None:
            num_id = num_id_node.get(qn("w:val"), "")
        level_node = numPr.find(qn("w:ilvl"))
        if level_node is not None:
            try:
                level = int(level_node.get(qn("w:val"), "0"))
            except ValueError:
                level = 0
        fmt, text, start, rtl = _numbering_definition(document, qn, num_id, level)
        key = (num_id, level, fmt, text, start, rtl)
        if key in numbering_seen:
            return
        numbering_seen.add(key)
        template.numbering.append(
            NumberingInfo(
                num_id=num_id,
                level=level,
                format=fmt,
                text=text,
                start=start,
                rtl=rtl,
                paragraph_location=location,
            )
        )

    def collect_paragraph(paragraph: Any, location: str) -> ParagraphInfo:
        nonlocal paragraph_counter
        paragraph_counter += 1
        text = paragraph.text or ""
        style_name = getattr(getattr(paragraph, "style", None), "name", "") or ""
        if style_name:
            used_styles.add(style_name)
        run_fonts: list[str] = []
        rtl_runs = 0
        for run in paragraph.runs:
            font = run.font
            record_font(font, "run")
            if font is not None and font.name:
                run_fonts.append(font.name)
            rPr = run._r.rPr
            if rPr is not None and rPr.find(qn("w:rtl")) is not None:
                rtl_runs += 1
        pPr = paragraph._p.pPr
        bidi = bool(pPr is not None and pPr.find(qn("w:bidi")) is not None)
        level = _heading_level(paragraph, qn)
        placeholders = _find_placeholders(text, location, style_name)
        template.placeholders.extend(placeholders)
        template.static_numeric_literals.extend(_find_static_numeric(text, location))
        collect_numbering(paragraph, location)
        info = ParagraphInfo(
            index=paragraph_counter,
            location=location,
            text=text,
            style=style_name,
            heading_level=level,
            alignment=_enum_name(getattr(paragraph, "alignment", None)),
            bidi=bidi,
            run_count=len(paragraph.runs),
            rtl_run_count=rtl_runs,
            font_names=tuple(dict.fromkeys(run_fonts)),
            placeholder_keys=tuple(dict.fromkeys(item.key for item in placeholders)),
        )
        return info

    # Sections and page layout.
    for index, section in enumerate(document.sections, start=1):
        template.page_layouts.append(_section_layout(section, index, qn))

    # Defined paragraph styles.
    try:
        from docx.enum.style import WD_STYLE_TYPE

        for style in document.styles:
            if style.type != WD_STYLE_TYPE.PARAGRAPH:
                continue
            snapshot = _font_snapshot(getattr(style, "font", None))
            record_font(getattr(style, "font", None), "style")
            paragraph_format = getattr(style, "paragraph_format", None)
            style_pPr = style.element.find(qn("w:pPr")) if hasattr(style, "element") else None
            template.styles.append(
                ParagraphStyleInfo(
                    name=style.name,
                    builtin=bool(getattr(style, "builtin", False)),
                    base_style=getattr(getattr(style, "base_style", None), "name", "") or "",
                    font=snapshot,
                    alignment=_enum_name(getattr(paragraph_format, "alignment", None)),
                    space_before_pt=_points(getattr(paragraph_format, "space_before", None)),
                    space_after_pt=_points(getattr(paragraph_format, "space_after", None)),
                    line_spacing=_line_spacing(getattr(paragraph_format, "line_spacing", None)),
                    keep_with_next=getattr(paragraph_format, "keep_with_next", None),
                    keep_together=getattr(paragraph_format, "keep_together", None),
                    outline_level=getattr(paragraph_format, "outline_level", None),
                    bidi=bool(style_pPr is not None and style_pPr.find(qn("w:bidi")) is not None),
                )
            )
    except Exception:  # noqa: BLE001 - a broken style table must not hide the rest
        template.warnings.append("خواندن جدول سبک‌های Word کامل نشد؛ بقیه ساختار گزارش شد.")

    # Body paragraphs.
    for index, paragraph in enumerate(document.paragraphs, start=1):
        if len(template.paragraphs) >= max_paragraphs:
            template.paragraphs_truncated = True
            break
        info = collect_paragraph(paragraph, f"body/paragraph[{index}]")
        template.paragraphs.append(info)
        if info.heading_level:
            template.headings.append(info)

    # Body tables (including nested tables).
    for index, table in enumerate(document.tables, start=1):
        if len(template.tables) >= max_tables:
            template.tables_truncated = True
            break
        _collect_table(
            table,
            f"table[{index}]",
            template,
            collect_paragraph,
            max_tables=max_tables,
        )

    # Headers and footers for every section/variant.
    for index, section in enumerate(document.sections, start=1):
        for kind, variants in (
            ("header", (("primary", section.header), ("first", section.first_page_header), ("even", section.even_page_header))),
            ("footer", (("primary", section.footer), ("first", section.first_page_footer), ("even", section.even_page_footer))),
        ):
            for variant, container in variants:
                info = _collect_header_footer(
                    container,
                    kind,
                    variant,
                    index,
                    template,
                    collect_paragraph,
                )
                if kind == "header":
                    template.headers.append(info)
                else:
                    template.footers.append(info)

    template.used_styles = sorted(used_styles)
    template.fonts = sorted(
        fonts.values(),
        key=lambda item: (item.source, item.name, item.size_pt or 0),
    )
    template.rtl = _rtl_summary(document, template, qn)
    template.warnings.extend(_template_warnings(template))
    return template


def _section_layout(section: Any, index: int, qn: Any) -> PageLayout:
    sectPr = section._sectPr
    return PageLayout(
        index=index,
        width_twips=_twips(getattr(section, "page_width", None)),
        height_twips=_twips(getattr(section, "page_height", None)),
        orientation=_enum_name(getattr(section, "orientation", None)),
        top_margin_twips=_twips(getattr(section, "top_margin", None)),
        bottom_margin_twips=_twips(getattr(section, "bottom_margin", None)),
        left_margin_twips=_twips(getattr(section, "left_margin", None)),
        right_margin_twips=_twips(getattr(section, "right_margin", None)),
        header_distance_twips=_twips(getattr(section, "header_distance", None)),
        footer_distance_twips=_twips(getattr(section, "footer_distance", None)),
        gutter_twips=_twips(getattr(section, "gutter", None)),
        different_first_page=bool(getattr(section, "different_first_page_header_footer", False)),
        start_type=_enum_name(getattr(section, "start_type", None)),
        bidi=sectPr is not None and sectPr.find(qn("w:bidi")) is not None,
        rtl_gutter=sectPr is not None and sectPr.find(qn("w:rtlGutter")) is not None,
    )


def _line_spacing(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return round(float(value), 3)
    return _points(value)


def _heading_level(paragraph: Any, qn: Any) -> int:
    style_name = getattr(getattr(paragraph, "style", None), "name", "") or ""
    match = _HEADING_PATTERN.match(style_name.strip())
    if match:
        try:
            return int(match.group(1))
        except ValueError:
            return 1
    pPr = paragraph._p.pPr
    if pPr is not None:
        outline = pPr.find(qn("w:outlineLvl"))
        if outline is not None:
            try:
                return int(outline.get(qn("w:val"), "0")) + 1
            except ValueError:
                return 0
    return 0


def _numbering_definition(document: Any, qn: Any, num_id: str, level: int) -> tuple[str, str, int | None, bool]:
    if not num_id:
        return "", "", None, False
    try:
        numbering = document.part.numbering_part.element
    except (AttributeError, NotImplementedError, KeyError):
        return "", "", None, False
    num = numbering.find(f".//{qn('w:num')}[@{qn('w:numId')}='{num_id}']")
    if num is None:
        return "", "", None, False
    abstract_node = num.find(qn("w:abstractNumId"))
    if abstract_node is None:
        return "", "", None, False
    abstract_id = abstract_node.get(qn("w:val"), "")
    abstract = numbering.find(f".//{qn('w:abstractNum')}[@{qn('w:abstractNumId')}='{abstract_id}']")
    if abstract is None:
        return "", "", None, False
    lvl = abstract.find(f"{qn('w:lvl')}[@{qn('w:ilvl')}='{level}']")
    if lvl is None:
        return "", "", None, False
    fmt_node = lvl.find(qn("w:numFmt"))
    text_node = lvl.find(qn("w:lvlText"))
    start_node = lvl.find(qn("w:start"))
    fmt = fmt_node.get(qn("w:val"), "") if fmt_node is not None else ""
    text = text_node.get(qn("w:val"), "") if text_node is not None else ""
    start = None
    if start_node is not None:
        try:
            start = int(start_node.get(qn("w:val"), "0"))
        except ValueError:
            start = None
    return fmt, text, start, lvl.find(qn("w:rtl")) is not None


def _find_placeholders(text: str, location: str, style: str) -> list[PlaceholderInfo]:
    return [
        PlaceholderInfo(key=match.group(1).strip(), raw=match.group(0), location=location, paragraph_style=style)
        for match in PLACEHOLDER_PATTERN.finditer(text or "")
    ]


def _find_static_numeric(text: str, location: str) -> list[StaticNumericLiteral]:
    stripped = PLACEHOLDER_PATTERN.sub(" ", text or "")
    return [StaticNumericLiteral(text=match.group(0), location=location) for match in _NUMERIC_PATTERN.finditer(stripped)]


def _collect_table(
    table: Any,
    location: str,
    template: ReportTemplate,
    collect_paragraph: Any,
    *,
    max_tables: int,
) -> None:
    qn = _qn()
    tblPr = table._tbl.tblPr
    header_row: tuple[str, ...] = ()
    cell_count = 0
    placeholder_keys: list[str] = []
    nested_tables: list[tuple[Any, str]] = []
    if table.rows:
        header_row = tuple(cell.text.strip() for cell in table.rows[0].cells)
    for row_index, row in enumerate(table.rows, start=1):
        seen_cells: set[int] = set()
        for cell_index, cell in enumerate(row.cells, start=1):
            identity = id(cell._tc)
            if identity in seen_cells:
                continue
            seen_cells.add(identity)
            cell_count += 1
            for paragraph_index, paragraph in enumerate(cell.paragraphs, start=1):
                info = collect_paragraph(
                    paragraph,
                    f"{location}/row[{row_index}]/cell[{cell_index}]/paragraph[{paragraph_index}]",
                )
                placeholder_keys.extend(info.placeholder_keys)
            for nested_index, nested in enumerate(cell.tables, start=1):
                nested_tables.append(
                    (
                        nested,
                        f"{location}/row[{row_index}]/cell[{cell_index}]/table[{nested_index}]",
                    )
                )
    if len(template.tables) >= max_tables:
        template.tables_truncated = True
        return
    template.tables.append(
        TableInfo(
            index=len(template.tables) + 1,
            location=location,
            rows=len(table.rows),
            columns=len(table.columns),
            style=getattr(getattr(table, "style", None), "name", "") or "",
            alignment=_enum_name(getattr(table, "alignment", None)),
            autofit=getattr(table, "autofit", None),
            bidi_visual=bool(tblPr is not None and tblPr.find(qn("w:bidiVisual")) is not None),
            header_row=header_row,
            cell_count=cell_count,
            placeholder_keys=tuple(dict.fromkeys(placeholder_keys)),
        )
    )
    for nested, nested_location in nested_tables:
        if len(template.tables) >= max_tables:
            template.tables_truncated = True
            return
        _collect_table(
            nested,
            nested_location,
            template,
            collect_paragraph,
            max_tables=max_tables,
        )


def _collect_header_footer(
    container: Any,
    kind: str,
    variant: str,
    section_index: int,
    template: ReportTemplate,
    collect_paragraph: Any,
) -> HeaderFooterInfo:
    linked = False
    try:
        linked = bool(container.is_linked_to_previous)
    except (AttributeError, ValueError):
        linked = False
    placeholders: list[str] = []
    texts: list[str] = []
    if not linked:
        for paragraph_index, paragraph in enumerate(container.paragraphs, start=1):
            info = collect_paragraph(
                paragraph,
                f"{kind}[{variant}]/section[{section_index}]/paragraph[{paragraph_index}]",
            )
            if info.text:
                texts.append(info.text)
            placeholders.extend(info.placeholder_keys)
        for table_index, table in enumerate(container.tables, start=1):
            for row_index, row in enumerate(table.rows, start=1):
                for cell_index, cell in enumerate(row.cells, start=1):
                    for paragraph_index, paragraph in enumerate(cell.paragraphs, start=1):
                        info = collect_paragraph(
                            paragraph,
                            f"{kind}[{variant}]/section[{section_index}]/table[{table_index}]"
                            f"/row[{row_index}]/cell[{cell_index}]/paragraph[{paragraph_index}]",
                        )
                        if info.text:
                            texts.append(info.text)
                        placeholders.extend(info.placeholder_keys)
    return HeaderFooterInfo(
        kind=kind,
        variant=variant,
        section_index=section_index,
        linked_to_previous=linked,
        text="\n".join(texts),
        table_count=len(container.tables) if not linked else 0,
        placeholder_keys=tuple(dict.fromkeys(placeholders)),
    )


def _rtl_summary(document: Any, template: ReportTemplate, qn: Any) -> dict[str, Any]:
    document_bidi = False
    document_rtl = False
    try:
        document_bidi = bool(document.styles.element.xpath("w:docDefaults/w:pPrDefault/w:pPr/w:bidi"))
        document_rtl = bool(document.styles.element.xpath("w:docDefaults/w:rPrDefault/w:rPr/w:rtl"))
    except Exception:  # noqa: BLE001 - optional metadata only
        pass
    return {
        "document_bidi": document_bidi,
        "document_rtl": document_rtl,
        "section_bidi": any(item.bidi for item in template.page_layouts),
        "section_rtl_gutter": any(item.rtl_gutter for item in template.page_layouts),
        "bidi_paragraphs": sum(1 for item in template.paragraphs if item.bidi),
        "rtl_runs": sum(item.rtl_run_count for item in template.paragraphs),
        "bidi_tables": sum(1 for item in template.tables if item.bidi_visual),
        "numbering_rtl": sum(1 for item in template.numbering if item.rtl),
    }


def _template_warnings(template: ReportTemplate) -> list[str]:
    warnings: list[str] = []
    if template.static_numeric_literals:
        warnings.append(
            f"قالب شامل {len(template.static_numeric_literals)} عدد ثابت است؛ این اعداد از قالب می‌آیند "
            "و به‌عنوان داده جدید گزارش استفاده نمی‌شوند."
        )
    if not template.placeholders:
        warnings.append("قالب هیچ جایگاه {{...}} ندارد؛ فقط ساختار و سبک‌ها قابل استفاده است.")
    if not template.page_layouts:
        warnings.append("هیچ بخش صفحه‌ای در قالب شناسایی نشد.")
    return warnings
