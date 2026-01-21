# ui/theme.py

from PyQt6.QtGui import QPalette, QColor
from PyQt6.QtWidgets import QWidget
from PyQt6.QtCore import Qt

# ==========================================
# PALETA DE COLORES MODERNA (Estilo SaaS)
# ==========================================
class ModernPalette:
    # Colores Primarios (Marca/Acción)
    PRIMARY_MAIN = "#2563EB"    # Azul intenso
    PRIMARY_HOVER = "#1D4ED8"   # Azul más oscuro
    PRIMARY_TEXT = "#FFFFFF"    # Texto sobre primario

    # Colores de Superficie y Fondo
    BACKGROUND_APP = "#F1F5F9"  # Gris muy claro fondo app
    SURFACE_MAIN = "#FFFFFF"    # Blanco puro paneles
    SURFACE_HOVER = "#F8FAFC"   # Blanco hueso
    
    # Texto y Bordes
    TEXT_PRIMARY = "#1E293B"    # Gris oscuro
    TEXT_SECONDARY = "#64748B"  # Gris medio
    BORDER_LIGHT = "#E2E8F0"    # Gris claro bordes
    BORDER_FOCUS = PRIMARY_MAIN 
    BORDER_SECONDARY_HOVER = "#CBD5E1" 

    # Estados
    SUCCESS = "#22C55E"
    DANGER = "#EF4444"
    WARNING = "#F59E0B"

    # Componentes
    TABLE_HEADER_BG = "#F8FAFC"
    INPUT_BG_READONLY = "#F1F5F9"

# ==========================================
# HOJA DE ESTILOS GLOBAL (QSS)
# ==========================================
MODERN_STYLESHEET = f"""
    /* --- APLICACIÓN GLOBAL --- */
    QWidget {{
        color: {ModernPalette.TEXT_PRIMARY};
        font-family: 'Segoe UI', Roboto, 'Helvetica Neue', sans-serif;
        font-size: 12px;
    }}

    QMainWindow, QDialog {{
        background-color: {ModernPalette.BACKGROUND_APP};
    }}
    
    QTabWidget::pane {{
        border: 1px solid {ModernPalette.BORDER_LIGHT};
        background: {ModernPalette.SURFACE_MAIN};
        border-radius: 4px;
    }}

    /* --- INPUTS --- */
    QLineEdit, QComboBox, QDateEdit, QTimeEdit {{
        background-color: {ModernPalette.SURFACE_MAIN};
        border: 1px solid {ModernPalette.BORDER_LIGHT};
        border-radius: 4px;
        padding: 4px 8px;
        min-height: 22px;
    }}

    QLineEdit:focus, QComboBox:focus, QDateEdit:focus, QTimeEdit:focus {{
        border: 1px solid {ModernPalette.BORDER_FOCUS};
    }}

    QLineEdit[readOnly="true"], QLineEdit:disabled, QComboBox:disabled {{
        background-color: {ModernPalette.INPUT_BG_READONLY};
        color: {ModernPalette.TEXT_SECONDARY};
        border-color: {ModernPalette.BORDER_LIGHT};
    }}

    QComboBox::drop-down {{
        subcontrol-origin: padding;
        subcontrol-position: top right;
        width: 20px;
        border-left-width: 0px;
    }}

    /* --- BOTONES --- */
    QPushButton {{
        background-color: {ModernPalette.SURFACE_MAIN};
        border: 1px solid {ModernPalette.BORDER_LIGHT};
        border-radius: 4px;
        padding: 5px 12px;
        font-weight: 600;
        color: {ModernPalette.TEXT_PRIMARY};
    }}
    QPushButton:hover {{
        background-color: {ModernPalette.SURFACE_HOVER};
        border-color: {ModernPalette.BORDER_SECONDARY_HOVER};
    }}
    QPushButton:pressed {{
        background-color: {ModernPalette.BORDER_LIGHT};
    }}

    QPushButton[variant="primary"] {{
        background-color: {ModernPalette.PRIMARY_MAIN};
        color: {ModernPalette.PRIMARY_TEXT};
        border: 1px solid {ModernPalette.PRIMARY_MAIN};
    }}
    QPushButton[variant="primary"]:hover {{
        background-color: {ModernPalette.PRIMARY_HOVER};
        border-color: {ModernPalette.PRIMARY_HOVER};
    }}

    QPushButton[variant="text"] {{
        background-color: transparent;
        border: none;
        color: {ModernPalette.TEXT_SECONDARY};
        padding: 4px 8px;
    }}
    QPushButton[variant="text"]:hover {{
        background-color: {ModernPalette.BORDER_LIGHT};
        color: {ModernPalette.TEXT_PRIMARY};
    }}

    /* --- TABLAS --- */
    QTableWidget {{
        background-color: {ModernPalette.SURFACE_MAIN};
        border: 1px solid {ModernPalette.BORDER_LIGHT};
        gridline-color: {ModernPalette.BORDER_LIGHT};
        selection-background-color: {ModernPalette.PRIMARY_MAIN}33; 
        selection-color: {ModernPalette.TEXT_PRIMARY};
        alternate-background-color: {ModernPalette.BACKGROUND_APP};
    }}

    QHeaderView::section {{
        background-color: {ModernPalette.TABLE_HEADER_BG};
        color: {ModernPalette.TEXT_SECONDARY};
        padding: 6px;
        border: none;
        border-bottom: 1px solid {ModernPalette.BORDER_LIGHT};
        border-right: 1px solid {ModernPalette.BORDER_LIGHT};
        font-weight: 700;
        font-size: 11px;
        text-transform: uppercase;
    }}
    
    QTableCornerButton::section {{
        background-color: {ModernPalette.TABLE_HEADER_BG};
        border: none;
    }}

    /* --- OTROS --- */
    QGroupBox {{
        border: 1px solid {ModernPalette.BORDER_LIGHT};
        border-radius: 6px;
        margin-top: 20px;
        background-color: {ModernPalette.SURFACE_MAIN};
    }}
    QGroupBox::title {{
        subcontrol-origin: margin;
        subcontrol-position: top left;
        padding: 0 5px;
        color: {ModernPalette.TEXT_PRIMARY};
        font-weight: 700;
    }}
    
    QLabel[role="sectionTitle"] {{
         font-size: 11px;
         font-weight: 800;
         color: {ModernPalette.TEXT_SECONDARY};
         text-transform: uppercase;
         letter-spacing: 0.5px;
    }}
"""

# AQUÍ ESTABA EL ERROR DE NOMBRE:
def apply_app_theme(app):  # <-- Renombrado a apply_app_theme
    """Aplica la paleta y la hoja de estilos moderna."""
    app.setStyle("Fusion")

    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor(ModernPalette.BACKGROUND_APP))
    palette.setColor(QPalette.ColorRole.WindowText, QColor(ModernPalette.TEXT_PRIMARY))
    palette.setColor(QPalette.ColorRole.Base, QColor(ModernPalette.SURFACE_MAIN))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor(ModernPalette.BACKGROUND_APP))
    palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(ModernPalette.TEXT_PRIMARY))
    palette.setColor(QPalette.ColorRole.ToolTipText, QColor(ModernPalette.SURFACE_MAIN))
    palette.setColor(QPalette.ColorRole.Text, QColor(ModernPalette.TEXT_PRIMARY))
    palette.setColor(QPalette.ColorRole.Button, QColor(ModernPalette.SURFACE_MAIN))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor(ModernPalette.TEXT_PRIMARY))
    palette.setColor(QPalette.ColorRole.BrightText, QColor(ModernPalette.DANGER))
    palette.setColor(QPalette.ColorRole.Link, QColor(ModernPalette.PRIMARY_MAIN))
    palette.setColor(QPalette.ColorRole.Highlight, QColor(ModernPalette.PRIMARY_MAIN))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor(ModernPalette.PRIMARY_TEXT))

    app.setPalette(palette)
    app.setStyleSheet(MODERN_STYLESHEET)

# FUNCIONES AUXILIARES NECESARIAS
def mark_error(widget: QWidget):
    widget.setStyleSheet(f"border: 1px solid {ModernPalette.DANGER}; background-color: #FEF2F2;")

def mark_success(widget: QWidget):
    widget.setStyleSheet(f"border: 1px solid {ModernPalette.SUCCESS}; background-color: #F0FDF4;")

def mark_normal(widget: QWidget):
    widget.setStyleSheet("")