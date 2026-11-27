"""Central design tokens and the Qt style sheet for the desktop interface.

Only Qt's own style engine is used; no third-party theme package is needed.
The palette carries every colour the widgets, icons and charts paint with, so
switching between the light and dark theme is a single re-apply.
"""

from __future__ import annotations

from dataclasses import dataclass

from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import QApplication

from local_office_ai.ui.desktop import icons


@dataclass(frozen=True)
class Palette:
    name: str
    canvas: str
    surface: str
    surface_alt: str
    surface_sunken: str
    sidebar: str
    sidebar_hover: str
    sidebar_active: str
    sidebar_text: str
    sidebar_muted: str
    text: str
    muted: str
    border: str
    border_strong: str
    accent: str
    accent_hover: str
    accent_soft: str
    on_accent: str
    success: str
    success_soft: str
    warning: str
    warning_soft: str
    danger: str
    danger_soft: str
    info: str
    info_soft: str
    chat_user: str
    chat_assistant: str
    input_bg: str
    shadow: str


LIGHT = Palette(
    name="light",
    canvas="#F3F6FB",
    surface="#FFFFFF",
    surface_alt="#F7F9FC",
    surface_sunken="#EDF2F8",
    sidebar="#0C2340",
    sidebar_hover="#173252",
    sidebar_active="#123F4C",
    sidebar_text="#DCE8F5",
    sidebar_muted="#8FA6BF",
    text="#10233A",
    muted="#5C7288",
    border="#E1E8F0",
    border_strong="#CBD6E3",
    accent="#0F766E",
    accent_hover="#0D9488",
    accent_soft="#E4F4F1",
    on_accent="#FFFFFF",
    success="#15803D",
    success_soft="#E8F6EC",
    warning="#B45309",
    warning_soft="#FCF0E1",
    danger="#B42318",
    danger_soft="#FCECEA",
    info="#1D4ED8",
    info_soft="#E9EFFD",
    chat_user="#E7F0FF",
    chat_assistant="#FFFFFF",
    input_bg="#FFFFFF",
    shadow="#10233A",
)

DARK = Palette(
    name="dark",
    canvas="#0B131E",
    surface="#131E2B",
    surface_alt="#182535",
    surface_sunken="#0F1A26",
    sidebar="#070E17",
    sidebar_hover="#152436",
    sidebar_active="#0F343B",
    sidebar_text="#DEE9F4",
    sidebar_muted="#7E93A8",
    text="#E7EEF7",
    muted="#93A6BA",
    border="#223245",
    border_strong="#324A63",
    accent="#2DD4BF",
    accent_hover="#5EEAD4",
    accent_soft="#10302F",
    on_accent="#04211F",
    success="#4ADE80",
    success_soft="#11281A",
    warning="#FBBF24",
    warning_soft="#2C2310",
    danger="#F87171",
    danger_soft="#301719",
    info="#60A5FA",
    info_soft="#13243B",
    chat_user="#17324D",
    chat_assistant="#16222F",
    input_bg="#0F1B27",
    shadow="#000000",
)


def palette_for(name: str) -> Palette:
    return DARK if str(name).lower() == "dark" else LIGHT


def _rgba(hex_color: str, alpha: float) -> str:
    color = QColor(hex_color)
    return f"rgba({color.red()}, {color.green()}, {color.blue()}, {alpha:.2f})"


def build_stylesheet(
    palette: Palette,
    *,
    family: str = "",
    base_size: int = 12,
    heading_size: int = 18,
) -> str:
    p = palette
    base = max(10, int(base_size))
    small = max(9, base - 1)
    caption = max(9, base - 2)
    heading = max(base + 3, int(heading_size))
    hero = min(heading + 8, 30)
    family_rule = f'font-family: "{family}";' if family else ""
    down_arrow = icons.chevron_url("down", p.muted)
    up_arrow = icons.chevron_url("up", p.muted)

    return f"""
* {{
    {family_rule}
    outline: none;
}}
QWidget {{
    color: {p.text};
    font-size: {base}pt;
    selection-background-color: {p.accent};
    selection-color: {p.on_accent};
}}
QMainWindow, QDialog {{
    background: {p.canvas};
}}
QWidget#Canvas, QWidget#PageBody, QWidget#PageScrollBody {{
    background: transparent;
}}
QScrollArea {{
    background: transparent;
    border: none;
}}
QScrollArea > QWidget > QWidget {{
    background: transparent;
}}
QScrollArea#ChatScroll QWidget#PageBody {{
    background: {p.surface_alt};
}}
QLabel {{
    background: transparent;
}}

/* ---------------------------------------------------------------- sidebar */
QFrame#Sidebar {{
    background: {p.sidebar};
    border: none;
}}
QLabel#BrandMark {{
    background: {p.accent};
    color: {p.on_accent};
    border-radius: 12px;
    font-size: {heading}pt;
    font-weight: 700;
}}
QLabel#BrandTitle {{
    color: {p.sidebar_text};
    font-size: {base + 1}pt;
    font-weight: 700;
}}
QLabel#BrandSub {{
    color: {p.sidebar_muted};
    font-size: {caption}pt;
}}
QLabel#SidebarSection {{
    color: {p.sidebar_muted};
    font-size: {caption}pt;
    font-weight: 600;
    padding: 0 14px;
}}
QFrame#SidebarDivider {{
    background: {_rgba(p.sidebar_text, 0.12)};
    border: none;
    max-height: 1px;
}}
QPushButton#Nav {{
    background: transparent;
    border: none;
    border-radius: 10px;
    padding: 0;
    text-align: right;
}}
QPushButton#Nav:hover {{
    background: {p.sidebar_hover};
}}
QPushButton#Nav:checked {{
    background: {p.sidebar_active};
}}
QPushButton#Nav:focus:!checked {{
    border: 1px solid {_rgba(p.accent, 0.55)};
}}
QLabel#NavLabel {{
    background: transparent;
    font-size: {base}pt;
    font-weight: 500;
}}
QLabel#NavSection {{
    background: transparent;
    color: {p.sidebar_muted};
    font-size: {caption}pt;
    font-weight: 600;
}}
QPushButton#NavToggle {{
    background: transparent;
    border: 1px solid {_rgba(p.sidebar_text, 0.14)};
    border-radius: 10px;
    padding: 0;
}}
QPushButton#NavToggle:hover {{
    background: {_rgba(p.sidebar_text, 0.10)};
}}
QPushButton#SidebarAction {{
    background: {_rgba(p.sidebar_text, 0.06)};
    border: 1px solid {_rgba(p.sidebar_text, 0.14)};
    border-radius: 11px;
    color: {p.sidebar_text};
    font-size: {small}pt;
    padding: 9px 12px;
    text-align: right;
}}
QPushButton#SidebarAction:hover {{
    background: {_rgba(p.sidebar_text, 0.12)};
}}
QLabel#SidebarNote {{
    color: {p.sidebar_muted};
    font-size: {caption}pt;
}}

/* ----------------------------------------------------------------- header */
QFrame#Header {{
    background: transparent;
    border: none;
}}
QLabel#PageTitle {{
    font-size: {heading + 3}pt;
    font-weight: 700;
}}
QLabel#PageSubtitle {{
    color: {p.muted};
    font-size: {small}pt;
}}
QLabel#Chip {{
    background: {p.surface_alt};
    border: 1px solid {p.border};
    border-radius: 10px;
    color: {p.muted};
    font-size: {small}pt;
    min-height: 20px;
    padding: 4px 12px;
}}
QLabel#Chip[tone="success"] {{
    background: {p.success_soft};
    border-color: {p.success_soft};
    color: {p.success};
}}
QLabel#Chip[tone="warning"] {{
    background: {p.warning_soft};
    border-color: {p.warning_soft};
    color: {p.warning};
}}
QLabel#Chip[tone="danger"] {{
    background: {p.danger_soft};
    border-color: {p.danger_soft};
    color: {p.danger};
}}
QLabel#Chip[tone="accent"] {{
    background: {p.accent_soft};
    border-color: {p.accent_soft};
    color: {p.accent};
}}
QLabel#Chip[tone="info"] {{
    background: {p.info_soft};
    border-color: {p.info_soft};
    color: {p.info};
}}

/* ------------------------------------------------------------------ cards */
QFrame#Card {{
    background: {p.surface};
    border: 1px solid {p.border};
    border-radius: 16px;
}}
QFrame#CardAlt {{
    background: {p.surface_alt};
    border: 1px solid {p.border};
    border-radius: 14px;
}}
QFrame#CardSunken {{
    background: {p.surface_sunken};
    border: 1px solid {p.border};
    border-radius: 14px;
}}
QFrame#Hero {{
    background: {p.surface};
    border: 1px solid {p.border};
    border-radius: 20px;
}}
QFrame#Divider {{
    background: {p.border};
    border: none;
    max-height: 1px;
}}

/* ----------------------------------------------------------------- labels */
QLabel#HeroTitle {{
    font-size: {hero}pt;
    font-weight: 700;
}}
QLabel#HeroSubtitle {{
    color: {p.muted};
    font-size: {base + 1}pt;
}}
QLabel#SectionTitle {{
    font-size: {heading}pt;
    font-weight: 700;
}}
QLabel#CardTitle {{
    font-size: {base + 1}pt;
    font-weight: 700;
}}
QLabel#Muted {{
    color: {p.muted};
    font-size: {base}pt;
}}
QLabel#Caption {{
    color: {p.muted};
    font-size: {caption}pt;
}}
QLabel#Value {{
    font-size: {heading}pt;
    font-weight: 700;
}}
QLabel#StatValue {{
    font-size: {heading + 2}pt;
    font-weight: 700;
}}
QLabel#StatLabel {{
    color: {p.muted};
    font-size: {caption}pt;
}}
QLabel#Unavailable {{
    color: {p.muted};
    font-size: {small}pt;
    font-style: italic;
}}
QLabel#ErrorText {{
    color: {p.danger};
}}
QLabel#SuccessText {{
    color: {p.success};
}}
QLabel#WarningText {{
    color: {p.warning};
}}
QLabel#Link {{
    color: {p.accent};
    font-size: {small}pt;
}}
QLabel#Badge {{
    background: {p.surface_alt};
    border: 1px solid {p.border};
    border-radius: 9px;
    color: {p.muted};
    font-size: {caption}pt;
    font-weight: 600;
    padding: 3px 9px;
}}
QLabel#Badge[tone="success"] {{
    background: {p.success_soft};
    border-color: transparent;
    color: {p.success};
}}
QLabel#Badge[tone="warning"] {{
    background: {p.warning_soft};
    border-color: transparent;
    color: {p.warning};
}}
QLabel#Badge[tone="danger"] {{
    background: {p.danger_soft};
    border-color: transparent;
    color: {p.danger};
}}
QLabel#Badge[tone="accent"] {{
    background: {p.accent_soft};
    border-color: transparent;
    color: {p.accent};
}}
QLabel#Badge[tone="info"] {{
    background: {p.info_soft};
    border-color: transparent;
    color: {p.info};
}}
QLabel#StepDot {{
    background: {p.surface_sunken};
    border-radius: 9px;
    color: {p.muted};
    font-size: {caption}pt;
    font-weight: 700;
}}
QLabel#StepDot[state="running"] {{
    background: {p.accent_soft};
    color: {p.accent};
}}
QLabel#StepDot[state="done"] {{
    background: {p.success_soft};
    color: {p.success};
}}
QLabel#StepDot[state="failed"] {{
    background: {p.danger_soft};
    color: {p.danger};
}}
QLabel#StepName[state="running"] {{
    font-weight: 700;
}}
QLabel#StepDuration {{
    color: {p.muted};
    font-size: {caption}pt;
}}

/* ---------------------------------------------------------------- buttons */
QPushButton {{
    background: {p.surface_alt};
    border: 1px solid {p.border};
    border-radius: 10px;
    color: {p.text};
    font-size: {base}pt;
    min-height: 18px;
    padding: 8px 16px;
}}
QPushButton:hover {{
    background: {p.surface_sunken};
}}
QPushButton:pressed {{
    background: {p.border};
}}
QPushButton:focus {{
    border: 1px solid {p.accent};
}}
QPushButton:disabled {{
    background: {p.surface_alt};
    border-color: {p.border};
    color: {p.muted};
}}
QPushButton#Primary {{
    background: {p.accent};
    border: 1px solid {p.accent};
    color: {p.on_accent};
    font-weight: 700;
    min-height: 20px;
    padding: 9px 22px;
}}
QPushButton#Primary:hover {{
    background: {p.accent_hover};
    border-color: {p.accent_hover};
}}
QPushButton#Primary:disabled {{
    background: {p.border_strong};
    border-color: {p.border_strong};
    color: {p.surface};
}}
QPushButton#Ghost {{
    background: transparent;
    border: 1px solid {p.border_strong};
}}
QPushButton#Ghost:hover {{
    background: {p.surface_alt};
}}
QPushButton#Ghost:disabled {{
    background: transparent;
    border-color: {p.border};
    color: {p.muted};
}}
QPushButton#Quiet {{
    background: transparent;
    border: none;
    color: {p.muted};
    padding: 6px 8px;
}}
QPushButton#Quiet:hover {{
    background: {p.surface_alt};
    color: {p.text};
}}
QPushButton#Quiet:disabled {{
    background: transparent;
    color: {p.muted};
}}
QPushButton#DangerQuiet {{
    background: transparent;
    border: none;
    color: {p.danger};
    padding: 6px 8px;
}}
QPushButton#DangerQuiet:hover {{
    background: {p.danger_soft};
}}
QPushButton#IconButton {{
    background: transparent;
    border: 1px solid transparent;
    border-radius: 9px;
    padding: 6px;
}}
QPushButton#IconButton:hover {{
    background: {p.surface_alt};
    border-color: {p.border};
}}
QPushButton#IconButton:disabled {{
    background: transparent;
    border-color: transparent;
}}
QPushButton#Segment {{
    background: transparent;
    border: 1px solid {p.border};
    border-radius: 10px;
    color: {p.muted};
    padding: 8px 16px;
}}
QPushButton#Segment:checked {{
    background: {p.accent_soft};
    border-color: {p.accent};
    color: {p.accent};
    font-weight: 700;
}}
QPushButton#Segment:disabled {{
    background: transparent;
    border-color: {p.border};
    color: {p.muted};
}}
QPushButton#PromptChip {{
    background: {p.surface_alt};
    border: 1px solid {p.border};
    border-radius: 13px;
    color: {p.muted};
    font-size: {caption}pt;
    min-height: 16px;
    padding: 5px 12px;
}}
QPushButton#PromptChip:hover {{
    background: {p.accent_soft};
    border-color: {p.accent};
    color: {p.accent};
}}

/* ----------------------------------------------------------------- inputs */
QLineEdit, QPlainTextEdit, QTextEdit, QComboBox, QSpinBox, QDoubleSpinBox {{
    background: {p.input_bg};
    border: 1px solid {p.border};
    border-radius: 9px;
    color: {p.text};
    font-size: {base}pt;
    min-height: 18px;
    padding: 8px 12px;
}}
QComboBox {{
    padding: 8px 12px 8px 32px;
}}
QComboBox::drop-down {{
    border: none;
    background: transparent;
    width: 28px;
    subcontrol-origin: border;
    subcontrol-position: center left;
}}
QComboBox::down-arrow {{
    image: url("{down_arrow}");
    width: 11px;
    height: 11px;
}}
QSpinBox, QDoubleSpinBox {{
    padding: 8px 10px 8px 26px;
    qproperty-alignment: AlignRight;
}}
QSpinBox::up-button, QDoubleSpinBox::up-button {{
    subcontrol-origin: border;
    subcontrol-position: top left;
    width: 22px;
    border: none;
    background: transparent;
}}
QSpinBox::down-button, QDoubleSpinBox::down-button {{
    subcontrol-origin: border;
    subcontrol-position: bottom left;
    width: 22px;
    border: none;
    background: transparent;
}}
QSpinBox::up-arrow, QDoubleSpinBox::up-arrow {{
    image: url("{up_arrow}");
    width: 9px;
    height: 9px;
}}
QSpinBox::down-arrow, QDoubleSpinBox::down-arrow {{
    image: url("{down_arrow}");
    width: 9px;
    height: 9px;
}}
QLineEdit:focus, QPlainTextEdit:focus, QTextEdit:focus, QComboBox:focus,
QSpinBox:focus, QDoubleSpinBox:focus {{
    border: 1px solid {p.accent};
}}
QLineEdit[readOnly="true"] {{
    background: {p.surface_alt};
    color: {p.muted};
}}
QPlainTextEdit#RequestInput {{
    border-radius: 14px;
    font-size: {base + 1}pt;
    padding: 12px 14px;
}}
QPlainTextEdit#ChatInput {{
    border-radius: 16px;
    padding: 10px 14px;
}}
QComboBox QAbstractItemView {{
    background: {p.surface};
    border: 1px solid {p.border};
    border-radius: 10px;
    padding: 4px;
    selection-background-color: {p.accent_soft};
    selection-color: {p.text};
}}
QCheckBox {{
    spacing: 8px;
    font-size: {base}pt;
}}

/* ------------------------------------------------------------------- tabs */
QTabWidget::pane {{
    border: none;
    background: transparent;
}}
QTabBar::tab {{
    background: transparent;
    border: none;
    border-bottom: 2px solid transparent;
    color: {p.muted};
    font-size: {small}pt;
    margin-left: 6px;
    padding: 8px 12px;
}}
QTabBar::tab:selected {{
    border-bottom: 2px solid {p.accent};
    color: {p.accent};
    font-weight: 700;
}}
QTabBar::tab:hover {{
    color: {p.text};
}}

/* ----------------------------------------------------------------- tables */
QTableWidget, QTableCornerButton::section {{
    background: {p.surface};
    alternate-background-color: {p.surface_alt};
    border: 1px solid {p.border};
    border-radius: 10px;
    gridline-color: transparent;
    outline: none;
    selection-background-color: {p.accent_soft};
    selection-color: {p.text};
}}
QHeaderView::section {{
    background: {p.surface_alt};
    border: none;
    border-bottom: 1px solid {p.border};
    color: {p.muted};
    font-size: {caption}pt;
    font-weight: 700;
    padding: 8px 10px;
}}
QHeaderView::section:first {{
    border-top-right-radius: 10px;
}}
QHeaderView::section:last {{
    border-top-left-radius: 10px;
}}
QTableWidget::item {{
    border: none;
    padding: 6px 8px;
}}
QTableWidget::item:selected {{
    background: {p.accent_soft};
    color: {p.text};
}}

/* --------------------------------------------------------------- progress */
QProgressBar {{
    background: {p.surface_sunken};
    border: none;
    border-radius: 4px;
    color: {p.muted};
    font-size: {caption}pt;
    max-height: 6px;
    text-align: center;
}}
QProgressBar::chunk {{
    background: {p.accent};
    border-radius: 4px;
}}
QProgressBar#Busy {{
    max-height: 4px;
}}

/* -------------------------------------------------------------- scrollbar */
QScrollBar:vertical {{
    background: transparent;
    border: none;
    margin: 2px 2px 2px 0;
    width: 9px;
}}
QScrollBar::handle:vertical {{
    background: {_rgba(p.text, 0.20)};
    border-radius: 4px;
    min-height: 44px;
}}
QScrollBar::handle:vertical:hover {{
    background: {_rgba(p.text, 0.34)};
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0;
}}
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{
    background: transparent;
}}
QScrollBar:horizontal {{
    background: transparent;
    border: none;
    height: 9px;
    margin: 0 2px 2px 2px;
}}
QScrollBar::handle:horizontal {{
    background: {_rgba(p.text, 0.20)};
    border-radius: 4px;
    min-width: 44px;
}}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
    width: 0;
}}
QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {{
    background: transparent;
}}

/* ------------------------------------------------------------------ chat */
QFrame#ChatBubble {{
    background: {p.chat_assistant};
    border: 1px solid {p.border};
    border-radius: 16px;
}}
QFrame#ChatBubble[role="user"] {{
    background: {p.chat_user};
    border-color: transparent;
}}
QFrame#ChatError {{
    background: {p.danger_soft};
    border: 1px solid {p.danger_soft};
    border-radius: 14px;
}}
QLabel#BubbleAuthor {{
    color: {p.muted};
    font-size: {caption}pt;
    font-weight: 700;
}}
QScrollArea#ChatScroll {{
    background: {p.surface_alt};
    border: 1px solid {p.border};
    border-radius: 16px;
}}

/* -------------------------------------------------------------- hardware */
QFrame#HardwarePanel {{
    background: {p.surface};
    border: none;
    border-top: 1px solid {p.border};
}}
QFrame#HardwareTile {{
    background: {p.surface_alt};
    border: 1px solid {p.border};
    border-radius: 12px;
}}
QLabel#HardwareValue {{
    font-size: {base + 2}pt;
    font-weight: 700;
}}
QLabel#HardwareName {{
    font-size: {caption}pt;
    color: {p.muted};
}}
QFrame#Banner {{
    background: {p.info_soft};
    border: 1px solid {p.info_soft};
    border-radius: 12px;
}}
QFrame#Banner[tone="success"] {{
    background: {p.success_soft};
    border-color: {p.success_soft};
}}
QFrame#Banner[tone="warning"] {{
    background: {p.warning_soft};
    border-color: {p.warning_soft};
}}
QFrame#Banner[tone="danger"] {{
    background: {p.danger_soft};
    border-color: {p.danger_soft};
}}
QLabel#BannerText {{
    font-size: {small}pt;
}}

/* ------------------------------------------------------- floating widget */
QFrame#FloatingCard {{
    background: {p.surface};
    border: 1px solid {p.border_strong};
    border-radius: 14px;
}}
QLabel#FloatingTitle {{
    font-size: {base}pt;
    font-weight: 700;
    color: {p.text};
}}
QLabel#FloatingProvider {{
    font-size: {caption}pt;
    color: {p.muted};
}}
QFrame#FloatingRow {{
    background: {p.surface_alt};
    border: 1px solid {p.border};
    border-radius: 10px;
}}
QLabel#FloatingName {{
    font-size: {caption}pt;
    color: {p.muted};
}}
QLabel#FloatingValue {{
    font-size: {base + 1}pt;
    font-weight: 700;
    color: {p.text};
}}
QLabel#FloatingDetail {{
    font-size: {caption}pt;
    color: {p.muted};
}}
QLabel#FloatingFootnote {{
    font-size: {caption}pt;
    color: {p.muted};
}}

/* --------------------------------------------------------------- tooltip */
QToolTip {{
    background: {p.sidebar};
    border: none;
    border-radius: 8px;
    color: {p.sidebar_text};
    font-size: {small}pt;
    padding: 6px 9px;
}}
"""


def apply_palette(app: QApplication, palette: Palette) -> None:
    """Keep native dialogs (message boxes, file pickers) in the same theme."""
    qt_palette = QPalette()
    qt_palette.setColor(QPalette.ColorRole.Window, QColor(palette.canvas))
    qt_palette.setColor(QPalette.ColorRole.WindowText, QColor(palette.text))
    qt_palette.setColor(QPalette.ColorRole.Base, QColor(palette.input_bg))
    qt_palette.setColor(QPalette.ColorRole.AlternateBase, QColor(palette.surface_alt))
    qt_palette.setColor(QPalette.ColorRole.Text, QColor(palette.text))
    qt_palette.setColor(QPalette.ColorRole.Button, QColor(palette.surface_alt))
    qt_palette.setColor(QPalette.ColorRole.ButtonText, QColor(palette.text))
    qt_palette.setColor(QPalette.ColorRole.Highlight, QColor(palette.accent))
    qt_palette.setColor(QPalette.ColorRole.HighlightedText, QColor(palette.on_accent))
    qt_palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(palette.sidebar))
    qt_palette.setColor(QPalette.ColorRole.ToolTipText, QColor(palette.sidebar_text))
    qt_palette.setColor(QPalette.ColorRole.PlaceholderText, QColor(palette.muted))
    app.setPalette(qt_palette)
