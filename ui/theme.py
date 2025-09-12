# ui/theme.py
from __future__ import annotations
import platform
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import QApplication, QWidget

# ---------- Paleta ----------
PALETTE = {
    # Primary
    "primary_700": "#1565C0",
    "primary_600": "#1E88E5",
    "primary_800": "#0D47A1",
    "primary_50":  "#E3F2FD",
    "primary_txt_on": "#FFFFFF",

    # Accent (CTA puntual)
    "accent_700": "#F57C00",

    # Semantic
    "success_600": "#2E7D32",
    "error_600":   "#C62828",
    "error_50":    "#FFEBEE",
    "warn_600":    "#ED6C02",
    "info_600":    "#0288D1",

    # Neutrals / Surfaces
    "neutral_900": "#111827",
    "neutral_700": "#374151",
    "neutral_600": "#4B5563",
    "neutral_500": "#6B7280",
    "neutral_400": "#9CA3AF",
    "neutral_300": "#D1D5DB",
    "neutral_200": "#E5E7EB",
    "neutral_100": "#F3F4F6",
    "neutral_50":  "#F9FAFB",
    "panel_bg":    "#F5F7FA",
    "white":       "#FFFFFF",
}

SPACING = 8  # grid de 8 px

def _base_font() -> QFont:
    sys = platform.system()
    family = "Segoe UI" if sys == "Windows" else ("SF Pro Text" if sys == "Darwin" else "Sans Serif")
    return QFont(family, 10)

def apply_app_theme(app: QApplication) -> None:
    """Aplica tipografía base y la hoja QSS global."""
    app.setFont(_base_font())
    app.setStyleSheet(build_qss())

def mark_error(widget: QWidget, on: bool) -> None:
    """Marca un control con estado de error (borde rojo + fondo sutil)."""
    widget.setProperty("hasError", bool(on))
    widget.style().unpolish(widget)
    widget.style().polish(widget)
    widget.update()

def build_qss() -> str:
    p = PALETTE
    return f"""
/* --------- Base --------- */
* {{
  font-family: "{_base_font().family()}", "Segoe UI", "SF Pro Text", sans-serif;
  color: {p['neutral_900']};
}}

QWidget {{
  background: {p['neutral_50']};
}}
QGroupBox {{
  background: {p['panel_bg']};
  border: 1px solid {p['neutral_200']};
  border-radius: 6px;
  margin-top: 16px; /* espacio para el título */
}}
QGroupBox::title {{
  subcontrol-origin: margin;
  subcontrol-position: top left;
  padding: 0 8px;
  margin-left: 6px;
  color: {p['neutral_700']};
  font-weight: 600;
  text-transform: none;
}}

QLabel {{
  color: {p['neutral_700']};
}}

QTabWidget::pane {{
  border: 1px solid {p['neutral_200']};
  border-radius: 6px;
  top: -2px;
  background: {p['white']};
}}
QTabBar::tab {{
  background: {p['neutral_100']};
  color: {p['neutral_700']};
  padding: 8px 12px;
  margin-right: 2px;
  border-top-left-radius: 6px;
  border-top-right-radius: 6px;
}}
QTabBar::tab:selected {{
  background: {p['white']};
  color: {p['neutral_900']};
  border: 1px solid {p['neutral_200']};
  border-bottom-color: {p['white']};
}}
QTabBar::tab:hover {{
  background: {p['neutral_100']};
}}

/* --------- Inputs --------- */
QLineEdit, QComboBox, QDateEdit, QTimeEdit {{
  background: {p['white']};
  border: 1px solid {p['neutral_300']};
  border-radius: 6px;
  padding: 6px 8px;
}}
QLineEdit:focus, QComboBox:focus, QDateEdit:focus, QTimeEdit:focus {{
  border: 1px solid {p['primary_700']};
}}
QLineEdit[hasError="true"], QComboBox[hasError="true"], QDateEdit[hasError="true"], QTimeEdit[hasError="true"] {{
  border: 1px solid {p['error_600']};
  background: {p['error_50']};
}}

QComboBox::drop-down {{
  border: 0px;
  width: 20px;
  margin-right: 4px;
}}
QComboBox QAbstractItemView {{
  background: {p['white']};
  border: 1px solid {p['neutral_300']};
  selection-background-color: {p['primary_50']};
  selection-color: {p['neutral_900']};
}}

/* --------- Buttons --------- */
QPushButton {{
  background: {p['neutral_100']};
  border: 1px solid {p['neutral_300']};
  border-radius: 6px;
  padding: 6px 12px;
}}
QPushButton:disabled {{
  color: {p['neutral_400']};
  background: {p['neutral_100']};
}}
/* Variants */
QPushButton[variant="primary"] {{
  background: {p['primary_700']};
  border: 1px solid {p['primary_700']};
  color: {p['primary_txt_on']};
}}
QPushButton[variant="primary"]:hover {{
  background: {p['primary_600']};
}}
QPushButton[variant="primary"]:pressed {{
  background: {p['primary_800']};
}}

QPushButton[variant="secondary"] {{
  background: {p['neutral_100']};
  border: 1px solid {p['neutral_300']};
  color: {p['neutral_900']};
}}
QPushButton[variant="secondary"]:hover {{
  background: {p['neutral_100']};
  border-color: {p['neutral_400']};
}}
QPushButton[variant="secondary"]:pressed {{
  background: {p['neutral_200']};
}}

QPushButton[variant="accent"] {{
  background: {p['accent_700']};
  border: 1px solid {p['accent_700']};
  color: {p['white']};
}}
QPushButton[variant="text"] {{
  background: transparent;
  border: none;
  color: {p['primary_700']};
  padding: 6px 8px;
}}
QPushButton[danger="true"] {{
  background: {p['error_600']};
  border-color: {p['error_600']};
  color: {p['white']};
}}
QPushButton[danger="true"]:hover {{
  background: #B71C1C;
}}

/* --------- Tables --------- */
QTableView {{
  background: {p['white']};
  alternate-background-color: {p['neutral_50']};
  gridline-color: {p['neutral_200']};
  selection-background-color: {p['primary_50']};
  selection-color: {p['neutral_900']};
  border: 1px solid {p['neutral_200']};
  border-radius: 6px;
}}
QTableView::item {{
  padding: 4px 6px; /* densidad compacta */
}}
QHeaderView::section {{
  background: {p['panel_bg']};
  color: {p['neutral_700']};
  padding: 6px 8px;
  border: 1px solid {p['neutral_200']};
  font-weight: 600;
  text-transform: uppercase;
}}

/* --------- Scrollbars --------- */
QScrollBar:vertical {{
  background: transparent;
  width: 10px;
  margin: 0px;
}}
QScrollBar::handle:vertical {{
  background: {p['neutral_300']};
  border-radius: 4px;
  min-height: 20px;
}}
QScrollBar::handle:vertical:hover {{
  background: {p['neutral_400']};
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
  height: 0px;
}}

QScrollBar:horizontal {{
  background: transparent;
  height: 10px;
  margin: 0px;
}}
QScrollBar::handle:horizontal {{
  background: {p['neutral_300']};
  border-radius: 4px;
  min-width: 20px;
}}
QScrollBar::handle:horizontal:hover {{
  background: {p['neutral_400']};
}}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
  width: 0px;
}}

/* --------- ProgressBar --------- */
QProgressBar {{
  border: 1px solid {p['neutral_200']};
  background: {p['neutral_100']};
  border-radius: 6px;
  text-align: center;
}}
QProgressBar::chunk {{
  background-color: {p['primary_700']};
  border-radius: 6px;
}}
"""
