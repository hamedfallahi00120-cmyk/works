"""Inline SVG icons rendered by Qt, so no icon pack or download is required."""

from __future__ import annotations

import tempfile
from functools import lru_cache
from pathlib import Path

from PySide6.QtCore import QByteArray, QRectF, Qt
from PySide6.QtGui import QIcon, QPainter, QPixmap
from PySide6.QtSvg import QSvgRenderer

_ICONS = {
    "home": '<path d="M4 10.5 12 4l8 6.5V20a1 1 0 0 1-1 1h-5v-6H10v6H5a1 1 0 0 1-1-1z"/>',
    "file": (
        '<path d="M7 3h7l4 4v14H7z"/><path d="M14 3v4h4"/>'
        '<path d="M9.5 12h5M9.5 15h5M9.5 18h3"/>'
    ),
    "folder": '<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>',
    "folder-open": (
        '<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v2H3z"/>'
        '<path d="M3 11h18l-2.5 8.5a1 1 0 0 1-1 .5H5.5a1 1 0 0 1-1-.7z"/>'
    ),
    "chat": (
        '<path d="M21 12a8 8 0 0 1-8 8H7l-4 3v-6.5A8 8 0 1 1 21 12z"/>'
        '<path d="M8.5 11h7M8.5 14h4"/>'
    ),
    "history": (
        '<path d="M3.5 12a8.5 8.5 0 1 0 2.8-6.3"/><path d="M3 4v4h4"/>'
        '<path d="M12 8v4.5l3 1.8"/>'
    ),
    "settings": (
        '<path d="M4 7h9M17 7h3M4 12h3M11 12h9M4 17h10M18 17h2"/>'
        '<circle cx="15" cy="7" r="2"/><circle cx="9" cy="12" r="2"/>'
        '<circle cx="16" cy="17" r="2"/>'
    ),
    "cpu": (
        '<rect x="7" y="7" width="10" height="10" rx="2"/>'
        '<path d="M9.5 3v4M14.5 3v4M9.5 17v4M14.5 17v4M3 9.5h4M3 14.5h4M17 9.5h4M17 14.5h4"/>'
    ),
    "activity": '<path d="M3 12h4l2.5-6 5 12 2.5-6h4"/>',
    "plus": '<path d="M12 5v14M5 12h14"/>',
    "close": '<path d="M6 6l12 12M18 6 6 18"/>',
    "trash": '<path d="M5 7h14M10 7V5h4v2M7 7l1 13h8l1-13"/><path d="M10 11v5M14 11v5"/>',
    "copy": (
        '<rect x="9" y="9" width="11" height="11" rx="2"/>'
        '<path d="M15 9V6a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v7a2 2 0 0 0 2 2h3"/>'
    ),
    "refresh": '<path d="M20 12a8 8 0 1 1-2.3-5.7"/><path d="M20 4v5h-5"/>',
    "check": '<path d="m5 12.5 4.5 4.5L19 7"/>',
    "alert": '<path d="M12 4 3 20h18z"/><path d="M12 9v5M12 17h.01"/>',
    "info": '<circle cx="12" cy="12" r="9"/><path d="M12 11v5M12 8h.01"/>',
    "chevron-down": '<path d="m6 9 6 6 6-6"/>',
    "chevron-up": '<path d="m6 15 6-6 6 6"/>',
    "send": '<path d="M3 11.5 21 3l-8.5 18-2-7z"/><path d="M21 3 10.5 14"/>',
    "sparkles": (
        '<path d="M12 3l1.8 5.2L19 10l-5.2 1.8L12 17l-1.8-5.2L5 10l5.2-1.8z"/>'
        '<path d="M18.5 15.5l.9 2.6 2.6.9-2.6.9-.9 2.6-.9-2.6-2.6-.9 2.6-.9z"/>'
    ),
    "eye": (
        '<path d="M2.5 12S6 6 12 6s9.5 6 9.5 6-3.5 6-9.5 6-9.5-6-9.5-6z"/>'
        '<circle cx="12" cy="12" r="3"/>'
    ),
    "play": '<path d="M8 5l11 7-11 7z"/>',
    "stop": '<rect x="6" y="6" width="12" height="12" rx="2"/>',
    "external": (
        '<path d="M14 4h6v6M20 4l-9 9"/>'
        '<path d="M19 14v5a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V6a1 1 0 0 1 1-1h5"/>'
    ),
    "sun": (
        '<circle cx="12" cy="12" r="4"/>'
        '<path d="M12 2v2M12 20v2M2 12h2M20 12h2M4.9 4.9l1.4 1.4'
        'M17.7 17.7l1.4 1.4M19.1 4.9l-1.4 1.4M6.3 17.7l-1.4 1.4"/>'
    ),
    "moon": '<path d="M20 14.5A8.5 8.5 0 0 1 9.5 4a8.5 8.5 0 1 0 10.5 10.5z"/>',
    "type": '<path d="M4 6V4h16v2M12 4v16M9 20h6"/>',
    "database": (
        '<ellipse cx="12" cy="6" rx="7" ry="3"/>'
        '<path d="M5 6v6c0 1.7 3.1 3 7 3s7-1.3 7-3V6'
        'M5 12v6c0 1.7 3.1 3 7 3s7-1.3 7-3v-6"/>'
    ),
    "monitor": '<rect x="3" y="4" width="18" height="12" rx="2"/><path d="M8 20h8M12 16v4"/>',
    "sidebar": '<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M9 4v16"/>',
    "thermometer": '<path d="M10 13V5a2 2 0 1 1 4 0v8a4 4 0 1 1-4 0z"/>',
    "memory": '<rect x="3" y="8" width="18" height="8" rx="2"/><path d="M7 8v8M11 8v8M15 8v8M19 8v8"/>',
    "gpu": (
        '<rect x="3" y="7" width="18" height="10" rx="2"/>'
        '<circle cx="9" cy="12" r="2.5"/><path d="M14 10h4M14 14h4M6 17v3M18 17v3"/>'
    ),
    "shield": '<path d="M12 3l8 3v6c0 4.5-3.4 7.6-8 9-4.6-1.4-8-4.5-8-9V6z"/><path d="m8.5 12 2.5 2.5 4.5-5"/>',
}


def svg_bytes(name: str, color: str) -> bytes:
    body = _ICONS.get(name)
    if body is None:
        body = _ICONS["info"]
    return (
        '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" '
        'viewBox="0 0 24 24" fill="none" '
        f'stroke="{color}" stroke-width="1.7" stroke-linecap="round" '
        f'stroke-linejoin="round">{body}</svg>'
    ).encode("utf-8")


@lru_cache(maxsize=512)
def pixmap(name: str, color: str, size: int = 20) -> QPixmap:
    render_size = max(8, int(size))
    supersample = 2
    canvas = QPixmap(render_size * supersample, render_size * supersample)
    canvas.fill(Qt.GlobalColor.transparent)
    painter = QPainter(canvas)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    renderer = QSvgRenderer(QByteArray(svg_bytes(name, color)))
    renderer.render(painter, QRectF(0, 0, render_size * supersample, render_size * supersample))
    painter.end()
    canvas.setDevicePixelRatio(supersample)
    return canvas


@lru_cache(maxsize=512)
def icon(name: str, color: str, size: int = 20) -> QIcon:
    return QIcon(pixmap(name, color, size))


def chevron_url(direction: str, color: str) -> str:
    points = "6 15 12 9 18 15" if direction == "up" else "6 9 12 15 18 9"
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" '
        f'fill="none" stroke="{color}" stroke-width="2.4" stroke-linecap="round" '
        f'stroke-linejoin="round"><polyline points="{points}"/></svg>'
    )
    folder = Path(tempfile.gettempdir()) / "local_office_ai_ui"
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / f"chevron-{direction}-{color.lstrip('#')}.svg"
    if not target.exists() or target.read_text(encoding="utf-8") != svg:
        target.write_text(svg, encoding="utf-8")
    return target.as_posix()


def clear_cache() -> None:
    pixmap.cache_clear()
    icon.cache_clear()
