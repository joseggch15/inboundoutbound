# ui/theme.py
# Full UI theme with palette, including warn_50 for soft warning highlights.
# Based on the original theme module and extended to add the pale warning color. :contentReference[oaicite:0]{index=0}

from __future__ import annotations
import platform
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import QApplication, QWidget

# ---------- Palette ----------
PALETTE = {
    # Primary
    "primary_700": "#1565C0",
    "primary_600": "#1E88E5",
    "primary_800": "#0D47A1",
    "primary_50": "#E3F2FD",
    "primary_txt_on": "#FFFFFF",
    # Accent (CTA)
    "accent_700": "#F57C00",
    # Semantic
    "success_600": "#2E7D32",
    "error_600": "#C62828",
    "error_50": "#FFEBEE",
    "warn_600": "#ED6C02",
    "warn_50": "#FFFBEA",  # NEW: pale warning used to highlight OFF→ON changes
    "info_600": "#0288D1",
    # Neutrals / Surfaces
    "neutral_900": "#111827",
    "neutral_700": "#374151",
    "neutral_600": "#4B5563",
    "neutral_500": "#6B7280",
    "neutral_400": "#9CA3AF",
    "neutral_300": "#D1D5DB",
    "neutral_200": "#E5E7EB",
    "neutral_100": "#F3F4F6",
    "neutral_50": "#F9FAFB",
    "panel_bg": "#F5F7FA",
    "white": "#FFFFFF",
}

SPACING = 8  # 8px grid


def _base_font() -> QFont:
    system = platform.system()
    family = (
        "Segoe UI"
        if system == "Windows"
        else ("SF Pro Text" if system == "Darwin" else "Sans Serif")
    )
    return QFont(family, 10)


def apply_app_theme(app: QApplication) -> None:
    """Apply base font and the global QSS sheet."""
    app.setFont(_base_font())
    app.setStyleSheet(build_qss())


def mark_error(widget: QWidget, on: bool) -> None:
    """Add/remove an error visual state (red border + subtle background)."""
    widget.setProperty("hasError", bool(on))
    widget.style().unpolish(widget)
    widget.style().polish(widget)
    widget.update()


def build_qss() -> str:
    p = PALETTE
    return f"""
/* ---------- Base ---------- */
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
  margin-top: 16px;
}}
QGroupBox::title {{
  subcontrol-origin: margin;
  subcontrol-position: top left;
  padding: 0 8px;
  margin-left: 6px;
  color: {p['neutral_700']};
  font-weight: 600;
}}

/* ---------- Tabs ---------- */
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

/* ---------- Inputs ---------- */
/* 1. Estilos compartidos base (MANTENER IGUAL) */
QLineEdit, QComboBox, QDateEdit, QTimeEdit {{
  background: {p['white']};
  border: 1px solid {p['neutral_300']};
  border-radius: 6px;
  padding: 6px 8px;
}}

/* 2. Focus states (MANTENER IGUAL) */
QLineEdit:focus, QComboBox:focus, QDateEdit:focus, QTimeEdit:focus {{
  border: 1px solid {p['primary_700']};
}}

/* 3. Error states (MANTENER IGUAL) */
QLineEdit[hasError="true"], QComboBox[hasError="true"], QDateEdit[hasError="true"], QTimeEdit[hasError="true"] {{
  border: 1px solid {p['error_600']};
  background: {p['error_50']};
}}

/* --- AQUÍ EMPIEZA EL CAMBIO (NUEVO ESTILO DE DROPDOWN) --- */

/* A. Padding extra a la derecha para que el texto no se monte sobre la flecha */
QComboBox {{
  padding-right: 32px; 
}}

/* B. Estilo del botón de la flecha (Drop-down subcontrol) */
QComboBox::drop-down {{
  subcontrol-origin: padding;
  subcontrol-position: top right;
  width: 28px;
  
  /* Borde izquierdo y fondo para que parezca un botón físico */
  border-left: 1px solid {p['neutral_200']}; 
  background: {p['neutral_100']};           
  
  /* Radios solo a la derecha para encajar en el contenedor */
  border-top-right-radius: 6px;
  border-bottom-right-radius: 6px;
}}

/* C. Feedback visual al pasar el mouse sobre la flecha */
QComboBox::drop-down:hover {{
  background: {p['neutral_200']}; 
}}



/* D. Feedback visual cuando el combo tiene el foco (borde azul) */
QComboBox:focus::drop-down {{
  border-left: 1px solid {p['primary_700']};
}}

/* E. Feedback visual si hay error (borde rojo) */
QComboBox[hasError="true"]::drop-down {{
  border-left: 1px solid {p['error_600']};
}}

/* C.2. Flecha (Icono) dentro del botón */
QComboBox::down-arrow {{
    /* Triángulo sólido relleno (fill) color #6B7280, sin comillas en la URL para evitar conflictos */
    image: url(data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNCAyNCIgZmlsbD0iIzZCNzI4MCI+PHBhdGggZD0iTTcgMTBsNSA1IDUtNXoiLz48L3N2Zz4=);
    
    width: 12px;
    height: 12px;
    
    /* CRÍTICO: Centra la flecha exactamente en medio del botón gris */
    subcontrol-position: center;
    subcontrol-origin: padding;
}}

/* F. Estilo del menú desplegable (Item View) - MANTENER IGUAL O AJUSTAR */
QComboBox QAbstractItemView {{
  background: {p['white']};
  border: 1px solid {p['neutral_300']};
  selection-background-color: {p['primary_50']};
  selection-color: {p['neutral_900']};
  outline: 0; /* Quita el borde punteado feo al seleccionar */
}}

/* ---------- Buttons ---------- */
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
QPushButton[variant="primary"] {{
  background: {p['primary_700']};
  border: 1px solid {p['primary_700']};
  color: {p['primary_txt_on']};
}}
QPushButton[variant="primary"]:hover {{ background: {p['primary_600']}; }}
QPushButton[variant="primary"]:pressed {{ background: {p['primary_800']}; }}

QPushButton[variant="secondary"] {{
  background: {p['neutral_100']};
  border: 1px solid {p['neutral_300']};
  color: {p['neutral_900']};
}}
QPushButton[variant="secondary"]:hover {{ border-color: {p['neutral_400']}; }}
QPushButton[variant="secondary"]:pressed {{ background: {p['neutral_200']}; }}

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

/* ---------- Tables ---------- */
QHeaderView::section {{
  background: {p['white']};
  color: {p['neutral_700']};
  padding: 6px 8px;
  border: 1px solid {p['neutral_200']};
  font-weight: 600;
}}

/* Custom Header Styles for Schedule Preview */
QHeaderView#fixedHeaders::section, QHeaderView#dateHeaders::section {{
  background-color: {p['primary_700']};
  color: {p['primary_txt_on']};
  border: none;
  border-right: 1px solid {p['white']};
  padding: 8px;
  font-weight: 600;
}}

QTableWidget {{
  gridline-color: {p['neutral_200']};
  selection-background-color: {p['primary_50']};
  selection-color: {p['neutral_900']};
  alternate-background-color: {p['neutral_100']};
}}
QTableWidget::item {{
  padding: 2px;
}}
"""
