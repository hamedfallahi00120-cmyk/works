"""Low-level OOXML helpers for Persian right-to-left Word output.

Applying right alignment alone does not make a Word paragraph RTL.  A real RTL
paragraph needs ``<w:bidi/>`` in its properties, every Persian run needs
``<w:rtl/>`` plus a complex-script font (``w:rFonts/@w:cs`` and ``w:szCs``),
and an RTL table needs ``<w:bidiVisual/>`` so its columns flow right to left.
The helpers here set exactly those properties and nothing else.
"""

from __future__ import annotations

import re
from typing import Any, Iterable

PERSIAN_RANGE = re.compile(r"[\u0600-\u06FF\u0750-\u077F\uFB50-\uFDFF\uFE70-\uFEFF]")

_PPR_BIDI_SUCCESSORS = (
    "w:adjustRightInd",
    "w:snapToGrid",
    "w:spacing",
    "w:ind",
    "w:contextualSpacing",
    "w:mirrorIndents",
    "w:suppressOverlap",
    "w:jc",
    "w:textDirection",
    "w:textAlignment",
    "w:textboxTightWrap",
    "w:outlineLvl",
    "w:divId",
    "w:cnfStyle",
    "w:rPr",
    "w:sectPr",
    "w:pPrChange",
)

_SECTPR_BIDI_SUCCESSORS = (
    "w:textDirection",
    "w:rtlGutter",
    "w:docGrid",
    "w:printerSettings",
    "w:sectPrChange",
)

_RPR_SZCS_SUCCESSORS = (
    "w:highlight",
    "w:u",
    "w:effect",
    "w:bdr",
    "w:shd",
    "w:fitText",
    "w:vertAlign",
    "w:rtl",
    "w:cs",
    "w:em",
    "w:lang",
    "w:eastAsianLayout",
    "w:specVanish",
    "w:oMath",
)


def qn(tag: str) -> Any:
    from docx.oxml.ns import qn as _qn

    return _qn(tag)


def oxml(tag: str) -> Any:
    from docx.oxml import OxmlElement

    return OxmlElement(tag)


def contains_persian(text: str) -> bool:
    return bool(PERSIAN_RANGE.search(text or ""))


def ensure_paragraph_bidi(paragraph: Any) -> None:
    """Add ``<w:bidi/>`` without disturbing the paragraph-property order."""

    p_pr = paragraph._p.get_or_add_pPr()
    if p_pr.find(qn("w:bidi")) is not None:
        return
    p_pr.insert_element_before(oxml("w:bidi"), *_PPR_BIDI_SUCCESSORS)


def set_run_rtl(
    run: Any,
    *,
    font_name: str | None = None,
    size_pt: float | None = None,
    bold: bool | None = None,
    color: Any = None,
) -> None:
    """Mark a run as RTL, preserving an existing font when none is supplied."""

    run.font.rtl = True
    r_pr = run._r.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if font_name is not None:
        if r_fonts is None:
            r_fonts = r_pr.get_or_add_rFonts()
        r_fonts.set(qn("w:cs"), font_name)
        r_fonts.set(qn("w:ascii"), font_name)
        r_fonts.set(qn("w:hAnsi"), font_name)
    elif r_fonts is not None and r_fonts.get(qn("w:cs")) is None:
        inherited = r_fonts.get(qn("w:ascii")) or r_fonts.get(qn("w:hAnsi"))
        if inherited:
            r_fonts.set(qn("w:cs"), inherited)
    _ensure_bidi_language(r_pr)
    if size_pt is not None:
        from docx.shared import Pt

        run.font.size = Pt(size_pt)
        sz_cs = r_pr.find(qn("w:szCs"))
        if sz_cs is None:
            sz_cs = oxml("w:szCs")
            r_pr.insert_element_before(sz_cs, *_RPR_SZCS_SUCCESSORS)
        sz_cs.set(qn("w:val"), str(int(round(size_pt * 2))))
    if bold is not None:
        run.font.bold = bold
        if bold:
            r_pr.get_or_add_bCs()
    if color is not None:
        run.font.color.rgb = color


def format_paragraph(
    paragraph: Any,
    *,
    alignment: Any = None,
    font_name: str | None = "B Nazanin",
    size_pt: float | None = None,
    bold: bool | None = None,
    color: Any = None,
    space_before_pt: float | None = None,
    space_after_pt: float | None = None,
    line_spacing: float | None = None,
    keep_with_next: bool | None = None,
) -> Any:
    """Apply RTL and typography to a paragraph and every run inside it."""

    ensure_paragraph_bidi(paragraph)
    if alignment is not None:
        paragraph.alignment = alignment
    fmt = paragraph.paragraph_format
    if space_before_pt is not None:
        fmt.space_before = _pt(space_before_pt)
    if space_after_pt is not None:
        fmt.space_after = _pt(space_after_pt)
    if line_spacing is not None:
        fmt.line_spacing = line_spacing
    if keep_with_next is not None:
        fmt.keep_with_next = keep_with_next
    for run in paragraph.runs:
        set_run_rtl(
            run,
            font_name=font_name,
            size_pt=size_pt,
            bold=bold,
            color=color,
        )
    return paragraph


def set_section_rtl(section: Any) -> None:
    sect_pr = section._sectPr
    if sect_pr.find(qn("w:bidi")) is None:
        sect_pr.insert_element_before(oxml("w:bidi"), *_SECTPR_BIDI_SUCCESSORS)
    if sect_pr.find(qn("w:rtlGutter")) is None:
        sect_pr.insert_element_before(oxml("w:rtlGutter"), "w:docGrid", "w:printerSettings", "w:sectPrChange")


def set_table_rtl(
    table: Any,
    *,
    header_rows: int = 1,
    shade_header: bool = True,
    bold_header: bool = True,
) -> None:
    """Make a table read right to left and style its header row."""

    tbl_pr = table._tbl.tblPr
    if tbl_pr.find(qn("w:bidiVisual")) is None:
        tbl_pr.get_or_add_bidiVisual()
    for index, row in enumerate(table.rows):
        tr_pr = row._tr.get_or_add_trPr()
        if index < header_rows and tr_pr.find(qn("w:tblHeader")) is None:
            tr_pr.insert_element_before(
                oxml("w:tblHeader"),
                "w:tblCellSpacing",
                "w:jc",
                "w:hidden",
                "w:ins",
                "w:del",
                "w:trPrChange",
            )
        for cell in row.cells:
            if index < header_rows and shade_header:
                _shade_cell(cell, "DCE6F1")
            for paragraph in cell.paragraphs:
                ensure_paragraph_bidi(paragraph)
                for run in paragraph.runs:
                    set_run_rtl(
                        run,
                        bold=(
                            True
                            if index < header_rows and bold_header
                            else None
                        ),
                    )


def _ensure_bidi_language(r_pr: Any) -> None:
    lang = r_pr.find(qn("w:lang"))
    if lang is None:
        lang = oxml("w:lang")
        r_pr.insert_element_before(
            lang,
            "w:eastAsianLayout",
            "w:specVanish",
            "w:oMath",
            "w:rPrChange",
        )
    lang.set(qn("w:bidi"), "fa-IR")


def _shade_cell(cell: Any, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    if tc_pr.find(qn("w:shd")) is not None:
        return
    shd = oxml("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tc_pr.insert_element_before(
        shd,
        "w:noWrap",
        "w:tcMar",
        "w:textDirection",
        "w:tcFitText",
        "w:vAlign",
        "w:hideMark",
        "w:headers",
        "w:cellIns",
        "w:cellDel",
        "w:cellMerge",
        "w:tcPrChange",
    )


def add_page_number_field(paragraph: Any, *, font_name: str = "B Nazanin", size_pt: float = 9) -> None:
    """Append ``صفحه X از Y`` using real Word PAGE/NUMPAGES fields."""

    label = paragraph.add_run("صفحه ")
    set_run_rtl(label, font_name=font_name, size_pt=size_pt)
    _field(paragraph, "PAGE", "1", font_name=font_name, size_pt=size_pt)
    middle = paragraph.add_run(" از ")
    set_run_rtl(middle, font_name=font_name, size_pt=size_pt)
    _field(paragraph, "NUMPAGES", "1", font_name=font_name, size_pt=size_pt)


def _field(paragraph: Any, instruction: str, placeholder: str, *, font_name: str, size_pt: float) -> None:
    run = paragraph.add_run()
    set_run_rtl(run, font_name=font_name, size_pt=size_pt)
    begin = oxml("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = oxml("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {instruction} "
    separate = oxml("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = oxml("w:t")
    text.text = placeholder
    end = oxml("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for element in (begin, instr, separate, text, end):
        run._r.append(element)


def paragraph_has_field(paragraph: Any) -> bool:
    return any(element.tag == qn("w:fldChar") for element in paragraph._p.iter())


def iter_header_footer_containers(document: Any) -> Iterable[Any]:
    for section in document.sections:
        for container in (
            section.header,
            section.first_page_header,
            section.even_page_header,
            section.footer,
            section.first_page_footer,
            section.even_page_footer,
        ):
            if _is_linked(container):
                continue
            yield container


def _is_linked(container: Any) -> bool:
    try:
        return bool(container.is_linked_to_previous)
    except (AttributeError, ValueError):
        return False


def _pt(value: float) -> Any:
    from docx.shared import Pt

    return Pt(value)
