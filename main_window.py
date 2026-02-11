# main_window.py
# Implements:
#  - REQ-001 (OFF→ON inline confirmation & warning highlight)
#  - REQ-002 (auto-center today in Schedule Preview)
#  - REQ-003 (date headers with weekday)
#  - Locations module (CRUD) and dropdowns for Pick Up / Drop Off in the register form
#  - Restores the blue “Save Changes to DB Excel” button (centered action bar)
#  - Hides the “User Pick Up / Drop Off (inline, saves immediately)” panel
#  - Minor UX refinements (responsive form layout, headers alignment)
#
# UI content is in English end-to-end.
# --- INJECTED FEATURE: Hover-card with IN/OUT times and Pick Up/Drop Off locations. ---
# --- MODIFIED: The hover-card (ShiftInfoCard) has been redesigned for compactness and better UX as per user requirements. ---
# --- UPDATED: The ShiftInfoCard background is now opaque with a shadow for better visibility as per technical requirements. ---
# --- NEW: Added filter panels to all relevant tabs as per specifications. ---
# --- MODIFICATION: Added color chips to Shift Types table and Status/Shift dropdown for better color visibility. ---
# --- NEW: Added "Remarks" feature to registration form and ShiftInfoCard. ---
# --- MODIFIED: Added 'created_by' tracking for rotation history, filtering the view based on the logged-in user.
# --- NEW: Added separate entry/exit date feature in registration form, with UI toggle and updated hover card info.
# --- MODIFICATION: Added time inputs for Entry/Exit dates and updated Hover Card to display them.
# --- Integración Final: Se añade el botón y la lógica para "Generate Onsite Stay Report" en el PlanStaffWidget.

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment
import json
import os
from PyQt6.QtWidgets import (
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QLabel,
    QLineEdit,
    QComboBox,
    QDateEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QGroupBox,
    QMessageBox,
    QFileDialog,
    QTabWidget,
    QApplication,
    QColorDialog,
    QTimeEdit,
    QSizePolicy,
    QToolButton,
    QAbstractItemView,
    QFontComboBox,
    QGraphicsDropShadowEffect,
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QFrame,
)
from PyQt6.QtCore import (
    QDate,
    Qt,
    pyqtSignal,
    QTime,
    QTimer,
    QSignalBlocker,
    QEvent,
    QRect,
    QSize,
)
from PyQt6.QtGui import QColor, QFont, QCursor, QIcon, QPixmap, QPainter, QPen, QKeySequence
from datetime import datetime, date as pydate, timedelta
from PyQt6.QtWidgets import QStyledItemDelegate
from clipboard_logic import ScheduleClipboardService


# App logic (unchanged)
import database_logic as db
import excel_logic as excel

# Theme helper (for visual error state)
from ui.theme import mark_error

# ---------- constants ----------
WARN_BG_HEX = "#FFFBEA"  # soft warning highlight
FROZEN_COLUMN_COUNT = 3  # ROLE, NAME, BADGE
NEWMONT_REPORT_HEADERS = [
    "#",
    "NAME",
    "FIRST NAME",
    "GID",
    "COMPANY",
    "DEPT",
    "FROM",
    "TO",
    "DATE",
    "TIME",
]
RGM_REPORT_HEADERS = [
    "NR",
    "NAME (Last, First Name)",
    "DEPARTMENT",
    "BADGE #",
    "POSITION / TITLE",
    "CREW A/B/C",
    "PICK UP LOCATION",
    "IN BOUND DATE",
    "Method Of Transport",
    "Location",
    "DEPT TIME",
    "ROSEBEL SITE OUT BOUND DATE",
]
DEBOUNCE_MS = 200

WEEKEND_HEADER_YELLOW = "#FFEB3B"  # Amarillo vibrante para cabeceras


def _is_off_like_payload(payload: dict) -> bool:
    """
    Determina si un turno representa un estado NO laborable.
    Es la Fuente Única de Verdad (SSoT) para:
      1. 'Do Not Mark Days' (kind='none')
      2. Turnos Base 'OFF'
      3. Turnos Custom con la bandera is_off=1 (ej. SICK, VACATION)
    """
    if not isinstance(payload, dict):
        return False

    # 1. Chequear 'Do Not Mark Days'
    kind = payload.get('kind')
    if kind == 'none':
        return True

    # 2. Chequear Turnos Custom marcados como 'Treat as OFF' (flag is_off)
    # SQLite a veces retorna 1/0, forzamos conversión a bool
    if bool(payload.get('is_off', False)):
        return True

    # 3. Fallback: El status base se llama estrictamente 'OFF'
    if kind == 'base' and (payload.get('status') or '').strip().upper() == 'OFF':
        return True

    return False

class ShiftCellDelegate(QStyledItemDelegate):
    def __init__(self, parent, get_options_callback):
        super().__init__(parent)
        self.get_options_callback = get_options_callback

    def createEditor(self, parent, option, index):
        combo = QComboBox(parent)

        # Reusar la función self._status_options_for_dialog()
        options = self.get_options_callback()
        for icon, text, data in options:
            combo.addItem(icon, text, data)

        combo.setEditable(False)
        return combo

    def setEditorData(self, editor, index):
        value = index.model().data(index, Qt.ItemDataRole.DisplayRole)
        i = editor.findText(value)
        if i >= 0:
            editor.setCurrentIndex(i)

    def setModelData(self, editor, model, index):
        selected_text = editor.currentText()
        model.setData(index, selected_text)

class RoleAdminWidget(QWidget):
    """
    Manages the master list of Roles. 
    Similar to LocationAdminWidget but for 'roles' table.
    """
    roles_changed = pyqtSignal()

    def __init__(self, scope_source: str | None = None):
        super().__init__()
        self.scope_source = scope_source 
        self.role_id = None
        self._debounce_timer = QTimer(self)
        self._debounce_timer.setSingleShot(True)
        self._debounce_timer.timeout.connect(self._reload_table)

        layout = QHBoxLayout(self)

        # --- Form ---
        form = QGridLayout()
        row = 0
        self.role_input = QLineEdit()
        form.addWidget(QLabel("Role / Dept Name:"), row, 0)
        form.addWidget(self.role_input, row, 1)
        row += 1

        # Admin Logic (Optional owner selection)
        self.owner_combo = None
        if self.scope_source is None:
            self.owner_combo = QComboBox()
            self.owner_combo.addItems(["RGM", "Newmont"])
            form.addWidget(QLabel("Owner:"), row, 0)
            form.addWidget(self.owner_combo, row, 1)
            row += 1

        btn_new = QPushButton("✨ New")
        btn_save = QPushButton("💾 Save")
        btn_del = QPushButton("❌ Delete")
        btn_save.setProperty("variant", "primary")
        btn_del.setProperty("danger", True)
        
        h = QHBoxLayout()
        #h.addWidget(btn_new);
        h.addWidget(btn_save); h.addWidget(btn_del)
        form.addLayout(h, row, 0, 1, 2)

        form_group = create_group_box("Manage Roles", form)
        form_group.setFixedWidth(400)

        # --- Table ---
        table_panel = QWidget()
        t_layout = QVBoxLayout(table_panel)
        
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search role...")
        self.search_input.textChanged.connect(lambda: self._debounce_timer.start(DEBOUNCE_MS))
        t_layout.addWidget(self.search_input)

        self.role_table = QTableWidget()
        self.role_table.setAlternatingRowColors(True)
        self.role_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.role_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.role_table.itemClicked.connect(self._load_to_form)
        t_layout.addWidget(self.role_table)

        table_group = create_group_box("Existing Roles", t_layout)

        layout.addWidget(form_group)
        layout.addWidget(table_group)

        # Connects
        btn_new.clicked.connect(self._new_role)
        btn_save.clicked.connect(self._save_role)
        btn_del.clicked.connect(self._delete_role)

        self._reload_table()

    def _reload_table(self):
        text = self.search_input.text()
        rows = db.get_roles_filtered(self.scope_source, text)
        
        headers = ["ID", "Source", "Role Name"] if self.scope_source is None else ["ID", "Role Name"]
        self.role_table.setRowCount(len(rows))
        self.role_table.setColumnCount(len(headers))
        self.role_table.setHorizontalHeaderLabels(headers)
        
        for r, row in enumerate(rows):
            self.role_table.setItem(r, 0, QTableWidgetItem(str(row["id"])))
            if self.scope_source is None:
                self.role_table.setItem(r, 1, QTableWidgetItem(row["source"]))
                self.role_table.setItem(r, 2, QTableWidgetItem(row["name"]))
            else:
                self.role_table.setItem(r, 1, QTableWidgetItem(row["name"]))
        
        self.role_table.setColumnHidden(0, True)
        self.role_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

    def _new_role(self):
        self.role_id = None
        self.role_input.clear()
        self.role_table.clearSelection()

    def _save_role(self):
        name = self.role_input.text().strip()
        if not name: return
        
        src = self.scope_source
        if src is None and self.owner_combo:
            src = self.owner_combo.currentText()

        if self.role_id:
            ok, msg = db.update_role(self.role_id, name, src)
        else:
            ok, msg = db.create_role(name, src)
        
        QMessageBox.information(self, "Role", msg)
        self._reload_table()
        self.roles_changed.emit()
        self._new_role()

    def _delete_role(self):
        if not self.role_id: return
        src = self.scope_source
        if src is None and self.owner_combo: # Admin case fallback
             src = self.role_table.item(self.role_table.currentRow(), 1).text()

        confirm = QMessageBox.question(self, "Confirm", "Delete this role?")
        if confirm == QMessageBox.StandardButton.Yes:
            ok, msg = db.delete_role(self.role_id, src)
            QMessageBox.information(self, "Role", msg)
            self._reload_table()
            self.roles_changed.emit()
            self._new_role()

    def _load_to_form(self, item):
        row = item.row()
        self.role_id = int(self.role_table.item(row, 0).text())
        if self.scope_source is None:
            self.role_input.setText(self.role_table.item(row, 2).text())
            src = self.role_table.item(row, 1).text()
            if self.owner_combo: self.owner_combo.setCurrentText(src)
        else:
            self.role_input.setText(self.role_table.item(row, 1).text())
# -------------------------------------------------------------
# MODIFIED: Hover card widget
# -------------------------------------------------------------
class ShiftInfoCard(QLabel):
    """
    A custom widget to display shift details in a styled card that repositions
    itself to stay within the screen viewport, based on detailed UI/UX rules.
    It now has an opaque background and shadow for readability.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.WindowType.ToolTip | Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)

        # Style according to the new design rules
        # Using light gray background for a softer, modern look as per recommendations.
        self.setStyleSheet(
            """
            QLabel {
                background-color: #374151; /* Gris oscuro */
                color: #FFFFFF;             /* LETRA BLANCA PARA EL TEXTO */
                border: 1px solid #4B5563;  /* Borde ligeramente más claro */
                border-radius: 8px;
                padding: 12px;
                font-size: 13px;
            }
        """
        )

        # Add a subtle box-shadow for depth, as requested.
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(12)
        # Corresponds to: box-shadow: 0 4px 12px rgba(0, 0, 0, 0.12)
        shadow.setColor(QColor(0, 0, 0, 30))  # ~12% alpha
        shadow.setOffset(0, 4)
        self.setGraphicsEffect(shadow)

        self.adjustSize()

    def show_info(self, anchor_rect: QRect, text: str):
        """
        Shows the card with rich text, positioning it relative to an anchor rectangle
        (e.g., a table cell) and ensuring it stays on screen by flipping vertically.
        """
        self.setText(text)
        # Set a max width to control wrapping and prevent the card from becoming too wide.
        self.setWordWrap(True)
        self.setMaximumWidth(320)
        self.adjustSize()

        # Positioning logic
        screen_geom = self.screen().availableGeometry()
        card_size = self.sizeHint()
        offset = 8  # 8-10px offset from the cell

        # Prefer top position, centered horizontally relative to the anchor
        top_pos = anchor_rect.topLeft()
        top_pos.setY(top_pos.y() - card_size.height() - offset)
        top_pos.setX(top_pos.x() + (anchor_rect.width() - card_size.width()) // 2)

        # Clamp horizontal position to stay within the screen
        if top_pos.x() < screen_geom.x():
            top_pos.setX(screen_geom.x() + offset)
        if top_pos.x() + card_size.width() > screen_geom.right():
            top_pos.setX(screen_geom.right() - card_size.width() - offset)

        # Check if top position is on-screen vertically
        if top_pos.y() >= screen_geom.y():
            self.move(top_pos)
        else:
            # Fallback to bottom position
            bottom_pos = anchor_rect.bottomLeft()
            bottom_pos.setY(bottom_pos.y() + offset)
            bottom_pos.setX(top_pos.x())  # Use the same clamped horizontal position

            # Ensure bottom position doesn't go off-screen either
            if bottom_pos.y() + card_size.height() > screen_geom.bottom():
                self.move(top_pos)  # If both fail, default to top
            else:
                self.move(bottom_pos)

        self.show()


# -------------------------------------------------------------
# Common helpers
# -------------------------------------------------------------
def _create_color_icon(hex_code: str, size: QSize = QSize(24, 12)) -> QIcon:
    """Creates a QIcon with a solid color rectangle (chip) for color previews."""
    if not hex_code or not hex_code.startswith("#"):
        hex_code = "#D1D5DB"  # Default gray for invalid/empty codes

    pixmap = QPixmap(size)
    pixmap.fill(Qt.GlobalColor.transparent)

    painter = QPainter(pixmap)
    try:
        color = QColor(hex_code)
        if not color.isValid():
            color = QColor("#D1D5DB")  # Gray fallback for invalid hex
    except Exception:
        color = QColor("#D1D5DB")

    painter.setBrush(color)

    # Draw border if color is very light or for fallback gray
    pen = QPen()
    if color.lightness() > 230 or not color.isValid():
        pen.setColor(
            QColor("#A9A9A9")
        )  # Darker gray for visibility on light backgrounds
        pen.setWidth(1)
    else:
        # For darker colors, use a transparent pen (no border)
        pen.setColor(Qt.GlobalColor.transparent)

    painter.setPen(pen)

    # Draw rounded rectangle slightly inset so border is fully visible
    rect = pixmap.rect().adjusted(0, 0, -1, -1)
    painter.drawRoundedRect(rect, 3.0, 3.0)
    painter.end()

    return QIcon(pixmap)


def create_group_box(title: str, inner_layout) -> QGroupBox:
    box = QGroupBox(title)
    font = box.font()
    font.setBold(True)
    box.setFont(font)
    box.setLayout(inner_layout)
    return box


def _clean(value) -> str:
    """Cleans cells for the UI: None/NaN/'nan'/'null' -> ''."""
    if value is None:
        return ""
    s = str(value).strip()
    if s.lower() in ("nan", "none", "null"):
        return ""
    return s


def _weekday_full_en(d: pydate) -> str:
    """English full weekday names."""
    names = [
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
        "Sunday",
    ]
    return names[d.weekday()]


from datetime import time as dtime


def _parse_hhmm_to_time(value, default: dtime) -> dtime:
    """
    Convierte 'HH:MM' o 'HH:MM:SS' a datetime.time.
    - Si ya es datetime.time, lo devuelve tal cual.
    - Si es None o cadena vacía, devuelve 'default'.
    - Si falla el parseo, también devuelve 'default'.
    """
    if isinstance(value, dtime):
        return value

    if not value:
        return default

    try:
        parts = str(value).strip().split(":")
        hour = int(parts[0])
        minute = int(parts[1]) if len(parts) > 1 else 0
        second = int(parts[2]) if len(parts) > 2 else 0
        return dtime(hour, minute, second)
    except Exception:
        return default


# -------------------------------------------------------------
# Collapsible group used to free vertical space by default
# -------------------------------------------------------------
class CollapsibleGroupBox(QWidget):
    """
    Simple collapsible container with a header button.
    - setContentLayout(layout) to attach the inner layout
    - setCollapsed(True/False) to toggle visibility
    """

    def __init__(self, title: str, collapsed: bool = True, parent=None):
        super().__init__(parent)
        self._collapsed = bool(collapsed)

        self._root = QVBoxLayout(self)
        self._root.setContentsMargins(0, 0, 0, 0)
        self._root.setSpacing(6)

        # Header
        header = QHBoxLayout()
        header.setContentsMargins(4, 0, 4, 0)

        self.toggle_btn = QToolButton()
        self.toggle_btn.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        self.toggle_btn.setArrowType(
            Qt.ArrowType.RightArrow if self._collapsed else Qt.ArrowType.DownArrow
        )
        self.toggle_btn.setText(title)
        self.toggle_btn.setCheckable(True)
        self.toggle_btn.setChecked(not self._collapsed)
        self.toggle_btn.clicked.connect(self._on_toggle)

        header.addWidget(self.toggle_btn)
        header.addStretch()
        self._root.addLayout(header)

        # Content
        self._content = QWidget()
        self._content.setVisible(not self._collapsed)
        self._content.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum
        )
        self._content_layout = QVBoxLayout(
            self._content
        )  # Use a layout that can hold a widget
        self._root.addWidget(self._content)

    def _on_toggle(self, checked: bool):
        self.setCollapsed(not checked)

    def setCollapsed(self, collapsed: bool):
        self._collapsed = bool(collapsed)
        self._content.setVisible(not self._collapsed)
        self.toggle_btn.setArrowType(
            Qt.ArrowType.RightArrow if self._collapsed else Qt.ArrowType.DownArrow
        )

    def setContentWidget(self, widget: QWidget):
        # Clear existing content
        while self._content_layout.count():
            item = self._content_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self._content_layout.addWidget(widget)

    def content(self) -> QWidget:
        return self._content


# -------------------------------------------------------------
# NEW Widget: Report Layout Settings
# -------------------------------------------------------------
class ReportSettingsWidget(QWidget):
    def __init__(self, username: str, source: str):
        super().__init__()
        self.username = username
        self.source = source
        self._color_buttons = {}

        layout = QGridLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setHorizontalSpacing(15)
        layout.setVerticalSpacing(10)

        # Font Name
        self.font_name_combo = QFontComboBox()
        layout.addWidget(QLabel("Font Name:"), 0, 0)
        layout.addWidget(self.font_name_combo, 0, 1)

        # Font Color
        self.font_color_btn = self._create_color_button("font_color")
        layout.addWidget(QLabel("Body Font Color:"), 1, 0)
        layout.addWidget(self.font_color_btn, 1, 1)

        # NEW: Date Format
        self.date_format_input = QLineEdit()
        self.date_format_input.setPlaceholderText("e.g., dd/mm/yyyy, yyyy-mm-dd, etc.")
        layout.addWidget(QLabel("Report Date Format:"), 2, 0)
        layout.addWidget(self.date_format_input, 2, 1)

        # Header BG Color
        self.header_bg_btn = self._create_color_button("header_bg_color")
        layout.addWidget(QLabel("Default Header Background:"), 3, 0)
        layout.addWidget(self.header_bg_btn, 3, 1)

        # Header Font Color
        self.header_font_btn = self._create_color_button("header_font_color")
        layout.addWidget(QLabel("Header Font Color:"), 4, 0)
        layout.addWidget(self.header_font_btn, 4, 1)

        # Column specific colors
        self.col_table = QTableWidget(0, 2)
        self.col_table.setHorizontalHeaderLabels(["Column Name", "Header Color"])
        self.col_table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.Stretch
        )
        self.col_table.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.ResizeToContents
        )
        self.col_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.col_table.cellDoubleClicked.connect(self._pick_column_color)
        layout.addWidget(
            QLabel("Column-specific Header Colors (double-click to change):"),
            5,
            0,
            1,
            2,
        )
        layout.addWidget(self.col_table, 6, 0, 1, 2)

        # Action Buttons
        self.save_btn = QPushButton("💾 Save Report Settings")
        self.save_btn.setProperty("variant", "primary")
        self.save_btn.clicked.connect(self.save_settings)

        self.reset_btn = QPushButton("🔄 Reset to Default")
        self.reset_btn.setProperty("variant", "secondary")
        self.reset_btn.clicked.connect(self._reset_settings)

        button_layout = QHBoxLayout()
        button_layout.addStretch()
        button_layout.addWidget(self.reset_btn)
        button_layout.addWidget(self.save_btn)
        button_layout.addStretch()

        layout.addLayout(button_layout, 7, 0, 1, 2)

        self.load_settings()

    def _create_color_button(self, key):
        btn = QPushButton("Pick Color...")
        btn.setProperty("key", key)
        btn.clicked.connect(self._pick_color)
        self._color_buttons[key] = btn
        return btn

    def _update_button_color(self, btn, color_hex):
        btn.setText(color_hex.upper())
        btn.setStyleSheet(
            f"background-color: {color_hex}; color: {'#FFFFFF' if QColor(color_hex).lightness() < 128 else '#000000'};"
        )

    def _pick_color(self):
        sender_btn = self.sender()
        key = sender_btn.property("key")
        current_color = sender_btn.text()
        color = QColorDialog.getColor(QColor(current_color), self, "Pick a Color")
        if color.isValid():
            self._update_button_color(sender_btn, color.name())

    def _pick_column_color(self, row, column):
        if column != 1:
            return
        header_item = self.col_table.item(row, 0)
        color_item = self.col_table.item(row, 1)
        if not header_item or not color_item:
            return

        current_color_hex = color_item.text()
        color = QColorDialog.getColor(
            QColor(current_color_hex), self, f"Color for {header_item.text()}"
        )
        if color.isValid():
            color_item.setText(color.name())
            color_item.setBackground(color)
            color_item.setForeground(
                QColor("#FFFFFF" if color.lightness() < 128 else "#000000")
            )

    def load_settings(self):
        settings = db.get_report_settings(self.username, self.source)
        self.font_name_combo.setCurrentFont(QFont(settings["font_name"]))
        self._update_button_color(self.font_color_btn, settings["font_color"])
        self.date_format_input.setText(settings.get("date_format", "dd/mm/yyyy"))
        self._update_button_color(self.header_bg_btn, settings["header_bg_color"])
        self._update_button_color(self.header_font_btn, settings["header_font_color"])

        headers = RGM_REPORT_HEADERS if self.source == "RGM" else NEWMONT_REPORT_HEADERS
        self.col_table.setRowCount(0)
        col_colors = settings.get("column_colors", {})
        for header in headers:
            row_pos = self.col_table.rowCount()
            self.col_table.insertRow(row_pos)
            self.col_table.setItem(row_pos, 0, QTableWidgetItem(header))
            color_hex = col_colors.get(header, settings["header_bg_color"])
            color_item = QTableWidgetItem(color_hex)
            color_item.setBackground(QColor(color_hex))
            color_item.setForeground(
                QColor("#FFFFFF" if QColor(color_hex).lightness() < 128 else "#000000")
            )
            self.col_table.setItem(row_pos, 1, color_item)

    def save_settings(self):
        col_colors = {}
        for row in range(self.col_table.rowCount()):
            header = self.col_table.item(row, 0).text()
            color = self.col_table.item(row, 1).text()
            col_colors[header] = color

        settings = {
            "font_name": self.font_name_combo.currentFont().family(),
            "font_color": self.font_color_btn.text(),
            "date_format": self.date_format_input.text().strip(),
            "header_bg_color": self.header_bg_btn.text(),
            "header_font_color": self.header_font_btn.text(),
            "column_colors": col_colors,
        }
        db.save_report_settings(self.username, self.source, settings)
        db.log_event(
            self.username,
            self.source,
            "SETTINGS_UPDATE",
            f"Report layout changed: {json.dumps(settings)}",
        )
        QMessageBox.information(
            self,
            "Settings Saved",
            "Report layout settings have been saved successfully.",
        )

    def _reset_settings(self):
        reply = QMessageBox.question(
            self,
            "Confirm Reset",
            "Are you sure you want to reset the report settings for this profile to their defaults?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )

        if reply == QMessageBox.StandardButton.Yes:
            db.delete_report_settings(self.username, self.source)
            self.load_settings()  # Reload to get defaults
            QMessageBox.information(
                self,
                "Settings Reset",
                "Report layout settings have been reset to default.",
            )


class DayScheduleEditor(QDialog):
    """
    Diálogo para editar un día: Status/Shift + Pick Up + Drop Off + Remark.
    Opcionalmente permite aplicar el cambio a varias celdas hacia la derecha.
    """

    def __init__(self, parent, status_options, locations, initial=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Day Schedule")
        self.setModal(True)
        initial = initial or {}

        layout = QVBoxLayout(self)

        # Status / Shift
        self.status_combo = QComboBox()
        for icon, text, payload in status_options:
            self.status_combo.addItem(icon, text, payload)

        # FIX → Preselect using payload, not label text
        status_code = (initial.get("status_text") or "").upper()
        if status_code:
            for i in range(self.status_combo.count()):
                data = self.status_combo.itemData(i)
                if not isinstance(data, dict):
                    continue

                kind = data.get("kind")
                if kind in ("none", "separator"):
                    continue

                # Base statuses: OFF / ON / ON NS
                if kind == "base":
                    if (data.get("status") or "").upper() == status_code:
                        self.status_combo.setCurrentIndex(i)
                        break

                # Custom shift types: use the 'code'
                if kind == "custom":
                    if (data.get("code") or "").upper() == status_code:
                        self.status_combo.setCurrentIndex(i)
                        break

        # Locations
        self.pickup_combo = QComboBox()
        self.dropoff_combo = QComboBox()
        self.pickup_combo.addItem("— Select location —", None)
        self.dropoff_combo.addItem("— Select location —", None)
        for loc in locations:
            self.pickup_combo.addItem(loc, loc)
            self.dropoff_combo.addItem(loc, loc)

        if initial.get("pickup"):
            i = self.pickup_combo.findData(initial["pickup"])
            if i >= 0:
                self.pickup_combo.setCurrentIndex(i)

        if initial.get("dropoff"):
            i = self.dropoff_combo.findData(initial["dropoff"])
            if i >= 0:
                self.dropoff_combo.setCurrentIndex(i)

        # Remarks
        self.remark_edit = QLineEdit(initial.get("remark", ""))
        self.remark_edit.setPlaceholderText("Optional: add a note for this day...")

        # ---------------------------------------------------------------------
        # MODIFICACIÓN: Instanciamos los checkboxes pero NO los mostramos
        # Mantenemos los objetos en memoria para no romper la lógica interna.
        # ---------------------------------------------------------------------
        self.apply_to_range_chk = QCheckBox("Apply to all selected cells to the right")
        self.apply_to_range_chk.setVisible(False)  # Aseguramos que sea invisible

        # --- Travel dates opcionales ---
        self.travel_diff_chk = QCheckBox("Travel dates are different from work day")
        self.travel_diff_chk.setVisible(False)  # Aseguramos que sea invisible

        # Valores por defecto (puedes ajustarlos luego si quieres)
        from PyQt6.QtCore import QDate, QTime

        today = QDate.currentDate()
        self.entry_date_edit = QDateEdit(today)
        self.entry_date_edit.setCalendarPopup(True)
        self.entry_date_edit.setDisplayFormat("dd/MM/yyyy")

        self.entry_time_edit = QTimeEdit(QTime(6, 0))
        self.entry_time_edit.setDisplayFormat("HH:mm")

        self.exit_date_edit = QDateEdit(today)
        self.exit_date_edit.setCalendarPopup(True)
        self.exit_date_edit.setDisplayFormat("dd/MM/yyyy")

        self.exit_time_edit = QTimeEdit(QTime(7, 0))
        self.exit_time_edit.setDisplayFormat("HH:mm")

        # Contenedor horizontal para Entry/Exit
        self._travel_container = QWidget()
        travel_layout = QHBoxLayout(self._travel_container)
        travel_layout.setContentsMargins(0, 0, 0, 0)
        travel_layout.addWidget(QLabel("Entry Date:"))
        travel_layout.addWidget(self.entry_date_edit)
        travel_layout.addWidget(QLabel("Time:"))
        travel_layout.addWidget(self.entry_time_edit)
        travel_layout.addSpacing(12)
        travel_layout.addWidget(QLabel("Exit Date:"))
        travel_layout.addWidget(self.exit_date_edit)
        travel_layout.addWidget(QLabel("Time:"))
        travel_layout.addWidget(self.exit_time_edit)
        travel_layout.addStretch()

        # Ocultamos el bloque por defecto (y como el checkbox no es visible, nunca se mostrará)
        self._travel_container.setVisible(False)
        self.travel_diff_chk.toggled.connect(self._travel_container.setVisible)

        # Form layout
        form = QFormLayout()
        form.addRow("Status / Shift:", self.status_combo)
        form.addRow("Pick Up:", self.pickup_combo)
        form.addRow("Drop Off:", self.dropoff_combo)
        form.addRow("Remarks:", self.remark_edit)
        self.status_combo.currentIndexChanged.connect(self._on_status_changed)
        # NOTA: NO agregamos travel_diff_chk ni _travel_container al layout visual
        # form.addRow("", self.travel_diff_chk)
        # form.addRow("", self._travel_container)

        # Botones
        btn_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)

        layout.addLayout(form)

        # NOTA: NO agregamos apply_to_range_chk al layout visual
        # layout.addWidget(self.apply_to_range_chk)

        layout.addWidget(btn_box)
        self._on_status_changed(self.status_combo.currentIndex())
    
    def _on_status_changed(self, index):
        """
        Refactorizado para soportar Turnos Custom marcados como OFF (SICK, etc).
        """
        data = self.status_combo.itemData(index)
        
        # Validación de seguridad
        if not isinstance(data, dict):
            return

        # LÓGICA: Si el estatus es tipo OFF, limpiar logística inmediatamente.
        if _is_off_like_payload(data):
            # Bloquear señales para evitar recursividad o loops de actualización
            with QSignalBlocker(self.pickup_combo), QSignalBlocker(self.dropoff_combo):
                self.pickup_combo.setCurrentIndex(0) # Asume que el índice 0 es vacío/None
                self.dropoff_combo.setCurrentIndex(0)
                
                # OPCIONAL: Feedback Visual - Deshabilitar campos
                # self.pickup_combo.setEnabled(False)
                # self.dropoff_combo.setEnabled(False)
        else:
            # OPCIONAL: Rehabilitar si se deshabilitaron
            # self.pickup_combo.setEnabled(True)
            # self.dropoff_combo.setEnabled(True)
            pass

    def result_payload(self):
        from datetime import datetime

        selection = self.status_combo.currentData() or {}

        entry_dt = None
        exit_dt = None

        # Esta condición siempre será False porque el checkbox está oculto y desmarcado,
        # lo cual es correcto para ocultar la funcionalidad.
        if self.travel_diff_chk.isChecked():
            entry_date = self.entry_date_edit.date().toPyDate()
            entry_time = self.entry_time_edit.time().toPyTime()
            exit_date = self.exit_date_edit.date().toPyDate()
            exit_time = self.exit_time_edit.time().toPyTime()

            entry_dt = datetime.combine(entry_date, entry_time)
            exit_dt = datetime.combine(exit_date, exit_time)

        return {
            "selection": selection,
            "pickup": self.pickup_combo.currentData(),
            "dropoff": self.dropoff_combo.currentData(),
            "remark": self.remark_edit.text().strip(),
            "apply_to_range": self.apply_to_range_chk.isChecked(),  # Será False
            "entry_datetime": entry_dt,  # Será None (se calculará por defecto fuera)
            "exit_datetime": exit_dt,  # Será None (se calculará por defecto fuera)
        }
    
    


# -------------------------------------------------------------
# Widget: Plan Staff (Preview, Register, Reports)
# -------------------------------------------------------------

class WeekendHeader(QHeaderView):
    """Cabecera personalizada que resalta los fines de semana en amarillo."""
    def __init__(self, orientation, parent=None, date_list=None):
        super().__init__(orientation, parent)
        self._date_list = date_list or []

    def set_dates(self, dates):
        self._date_list = dates
        self.viewport().update()

    def paintSection(self, painter, rect, logicalIndex):
        # Verificamos si el índice corresponde a un fin de semana
        is_weekend = False
        if 0 <= logicalIndex < len(self._date_list):
            d = self._date_list[logicalIndex]
            is_weekend = d.weekday() >= 5  # 5=Saturday, 6=Sunday

        if is_weekend:
            painter.save()
            # Pintamos el fondo amarillo
            painter.fillRect(rect, QColor(WEEKEND_HEADER_YELLOW))
            
            # Dibujamos el borde (para mantener la estética de la grilla)
            painter.setPen(QPen(QColor("#CCCCCC")))
            painter.drawRect(rect.adjusted(0, 0, -1, -1))
            
            # Pintamos el texto (Date + Day)
            painter.setPen(QPen(QColor("#000000"))) # Texto negro sobre amarillo
            text = self.model().headerData(logicalIndex, Qt.Orientation.Horizontal, Qt.ItemDataRole.DisplayRole)
            painter.drawText(rect, Qt.AlignmentFlag.AlignCenter, text)
            painter.restore()
        else:
            # Comportamiento normal para días de semana
            super().paintSection(painter, rect, logicalIndex)
            
class PlanStaffWidget(QWidget):
    # Emitted after saving a change so the Rotation History tab can refresh
    rotation_changed = pyqtSignal()
    excel_path_changed = pyqtSignal(str, str)

    def __init__(self, source: str, excel_file: str, logged_username: str, preloaded_data=None):
        super().__init__()
        self.source = source  # "RGM" | "Newmont"
        self.excel_file = excel_file
        self.logged_username = logged_username or "Unknown"
        self._is_internal_update = False
        self._last_excel_mtime = None
        self._missing_prompt_shown = False
        self._custom_shift_map = db.get_shift_type_map(self.source)
        self._initial_preloaded_df = preloaded_data

        # For REQ-001 tracking
        self._loading_preview = False
        self._cell_original_values = {}  # (row, col) -> original text
        self._row_identities = []  # index -> {"name":..., "badge":...}
        self._date_col_dates = []  # schedule_table column index -> pydate
        self._warn_highlight_keys = set()  # {"<badge>|YYYY-MM-DD", ...}
        self._bulk_editing = False  # para detectar relleno masivo por arrastre

        # --- NUEVO: Bandera para evitar doble apertura de diálogo ---
        self._is_handling_change = False

        # ---------- root layout ----------
        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(8)

        # --- File status (health / SSoT) ---
        status_layout = QHBoxLayout()
        status_layout.setContentsMargins(8, 4, 8, 4)

        self.excel_health_label = QLabel("Excel status: checking.")
        self.excel_health_label.setStyleSheet("font-weight: bold;")

        self.validate_button = QPushButton("🧪 Validate Excel Structure")
        self.validate_button.clicked.connect(self.validate_excel_structure_ui)
        self.validate_button.setProperty("variant", "text")

        self.compare_button = QPushButton("🔎 Compare Excel vs DB")
        self.compare_button.clicked.connect(self.compare_excel_db_ui)
        self.compare_button.setProperty("variant", "text")

        self.refresh_button = QPushButton("🔄 Refresh Excel from DB")
        self.refresh_button.clicked.connect(self.refresh_excel_from_db_ui)
        self.refresh_button.setProperty("variant", "text")

        self.regen_button = QPushButton("🛠️ Regenerate Plan Staff from DB")
        self.regen_button.clicked.connect(self.regenerate_excel_from_db)
        self.regen_button.setProperty("variant", "secondary")

        status_layout.addWidget(self.excel_health_label)
        status_layout.addStretch()
        status_layout.addWidget(self.validate_button)
        status_layout.addWidget(self.compare_button)
        status_layout.addWidget(self.refresh_button)
        status_layout.addWidget(self.regen_button)
        root.addWidget(create_group_box("File Status  SSoT", status_layout), 0)

        # --- Schedule preview (enlarged) ---
        preview_container = QVBoxLayout()
        preview_container.setContentsMargins(0, 0, 0, 0)

        tables_layout = QHBoxLayout()
        tables_layout.setSpacing(0)
        tables_layout.setContentsMargins(0, 0, 0, 0)

        self.frozen_table = QTableWidget()
        self.schedule_table = QTableWidget()
        # Inyectamos nuestra cabecera personalizada
        self.weekend_header = WeekendHeader(Qt.Orientation.Horizontal, self.schedule_table)
        self.schedule_table.setHorizontalHeader(self.weekend_header)

        # Delegate para que las celdas de schedule usen el combo de Status/Shift
        self.shift_delegate = ShiftCellDelegate(
            self.schedule_table,
            self._status_options_for_dialog,
        )
        self.schedule_table.setItemDelegate(self.shift_delegate)

        # Set object names for styling headers
        self.frozen_table.horizontalHeader().setObjectName("fixedHeaders")
        self.schedule_table.horizontalHeader().setObjectName("dateHeaders")

        # Freeze (left) table is read-only; main table is editable for inline changes
        self.frozen_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)

        # Scroll lock between frozen & main tables
        self.frozen_table.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self.schedule_table.verticalScrollBar().valueChanged.connect(
            self.frozen_table.verticalScrollBar().setValue
        )
        self.frozen_table.verticalScrollBar().valueChanged.connect(
            self.schedule_table.verticalScrollBar().setValue
        )

        # Dense layout & alternating rows for readability
        self.frozen_table.setAlternatingRowColors(True)
        self.schedule_table.setAlternatingRowColors(True)
        self.frozen_table.setObjectName("FrozenTable")
        self.schedule_table.setObjectName("ScheduleTable")

        # REQ-001: detect inline edits
        self.schedule_table.itemChanged.connect(self._on_schedule_cell_changed)

        # Center headers
        self.schedule_table.horizontalHeader().setDefaultAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        # --- Frozen panel width policy (ensure 3 fixed columns visible) ---
        self.frozen_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.ResizeToContents
        )
        self.frozen_table.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOn
        )
        self.frozen_table.setSizePolicy(
            QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Expanding
        )
        self.schedule_table.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        tables_layout.addWidget(self.frozen_table)
        tables_layout.addWidget(self.schedule_table, 1)
        preview_container.addLayout(tables_layout)

        preview_title = f"🗓️ Schedule Preview ({os.path.basename(self.excel_file)})"
        preview_group = create_group_box(preview_title, preview_container)
        root.addWidget(preview_group, 12)  # give the preview most of the space

        # --- Register Employee Schedule (compact, multi-column, collapsible) ---
        self.registration_section = CollapsibleGroupBox(
            "1. Register Employee Schedule (DB is SSoT)", collapsed=True
        )
        register_layout = self._build_registration_form()
        self.registration_section.setContentWidget(register_layout)
        root.addWidget(self.registration_section, 1)

        # --- Transportation Report & Export ---
        report_layout = QHBoxLayout()
        report_layout.setContentsMargins(8, 4, 8, 4)

        report_layout.addWidget(QLabel("START Date:"))
        self.report_start_date = QDateEdit(QDate.currentDate())
        self.report_start_date.setCalendarPopup(True)
        self.report_start_date.setDisplayFormat("dd/MM/yyyy")
        report_layout.addWidget(self.report_start_date)

        report_layout.addWidget(QLabel("END Date:"))
        self.report_end_date = QDateEdit(QDate.currentDate().addDays(7))
        self.report_end_date.setCalendarPopup(True)
        self.report_end_date.setDisplayFormat("dd/MM/yyyy")
        report_layout.addWidget(self.report_end_date)

        report_button = QPushButton("🚀 Generate Inbound Outbound Report")
        report_button.clicked.connect(self.generate_report)
        report_button.setProperty("variant", "primary")
        report_layout.addWidget(report_button)

        # =======================================================
        # NUEVA FUNCIONALIDAD: Botón para reporte de estadías
        # =======================================================
        self.stay_report_btn = QPushButton("🏕️ Onsite Stay Report")
        self.stay_report_btn.clicked.connect(self._generate_stay_report)
        self.stay_report_btn.setProperty("variant", "secondary")
        report_layout.addWidget(self.stay_report_btn)

        export_button = QPushButton("📤 Export Plan Staff (.xlsx) from DB")
        export_button.clicked.connect(self.export_plan_from_db)
        export_button.setProperty("variant", "secondary")
        report_layout.addWidget(export_button)

        report_group = create_group_box(
            f"2. Transportation Report & Export (from {os.path.basename(self.excel_file)})",
            report_layout,
        )
        root.addWidget(report_group, 1)

        # --- Hover card setup ---
        self._shift_info_card = ShiftInfoCard(self)
        self.schedule_table.setMouseTracking(True)
        self.schedule_table.cellEntered.connect(self._show_shift_tooltip)
        self.schedule_table.viewport().installEventFilter(self)

        # Initial data load
        self.refresh_ui_data(use_preloaded=True)

        # --- File monitor: detect moved/deleted/renamed file ---
        self.file_watch_timer = QTimer(self)
        self.file_watch_timer.setInterval(2000)  # 2s
        self.file_watch_timer.timeout.connect(self.check_excel_health)
        self.file_watch_timer.start()
        self.check_excel_health()

        # Track current responsive columns for the register grid
        self._current_form_cols = 3
        self._rebuild_registration_grid(self._current_form_cols)

    def _smart_reflow(self, cursor, badge: str, center_date: datetime.date):
        """
        Recalcula y corrige los campos start_date y end_date para asegurar continuidad
        alrededor de una fecha modificada. Soluciona la fragmentación de rangos.
        """
        # 1. Definir zona de "reconstrucción" (60 días antes y después para seguridad)
        margin = timedelta(days=60)
        search_start = center_date - margin
        search_end = center_date + margin

        # 2. Obtener la verdad atómica (Fecha + Tipo de Turno) ordenados cronológicamente
        cursor.execute("""
            SELECT date, shift_type 
            FROM schedules 
            WHERE badge = ? AND date BETWEEN ? AND ?
            ORDER BY date ASC
        """, (badge, search_start, search_end))
        
        rows = cursor.fetchall()
        if not rows:
            return

        # Convertir a objetos manipulables
        # data_points = lista de {'date': obj_date, 'type': str}
        data_points = []
        for r in rows:
            d_val = r[0]
            # Manejo robusto de fechas (si viene como string o como objeto)
            if isinstance(d_val, str):
                try:
                    d_obj = datetime.strptime(d_val, "%Y-%m-%d").date()
                except ValueError:
                    continue # Saltar fechas inválidas
            else:
                d_obj = d_val
            data_points.append({'date': d_obj, 'type': r[1]})

        if not data_points:
            return

        # 3. Algoritmo de Agrupación (Clustering)
        # Recorre los días y agrupa los que sean consecutivos Y tengan el mismo shift_type
        updates = []
        
        # Iniciamos el primer grupo
        current_block_start = data_points[0]['date']
        # current_type = data_points[0]['type'] # No usado directamente en el loop, solo comparativa
        block_members = [data_points[0]['date']] # Lista de fechas en el bloque actual

        for i in range(1, len(data_points)):
            prev = data_points[i-1]
            curr = data_points[i]
            
            is_consecutive = (curr['date'] - prev['date']).days == 1
            is_same_type = (curr['type'] == prev['type'])

            if is_consecutive and is_same_type:
                # Continuar el bloque actual
                block_members.append(curr['date'])
            else:
                # El bloque se rompió. Cerrar el bloque anterior y guardar sus actualizaciones.
                start_str = current_block_start.strftime("%Y-%m-%d")
                end_str = block_members[-1].strftime("%Y-%m-%d")
                
                for member_date in block_members:
                    updates.append((start_str, end_str, badge, member_date))
                
                # Iniciar nuevo bloque
                current_block_start = curr['date']
                block_members = [curr['date']]

        # Cerrar el último bloque pendiente
        if block_members:
            start_str = current_block_start.strftime("%Y-%m-%d")
            end_str = block_members[-1].strftime("%Y-%m-%d")
            for member_date in block_members:
                updates.append((start_str, end_str, badge, member_date))

        # 4. Escritura Masiva (Bulk Update)
        # Actualizamos start_date y end_date para todas las filas procesadas
        if updates:
            cursor.executemany("""
                UPDATE schedules 
                SET start_date = ?, end_date = ?
                WHERE badge = ? AND date = ?
            """, updates)
    
    
    def eventFilter(self, source, event):
        # Comportamiento extra para la tabla de horario:
        #  - Ocultar la tarjetita al salir
        #  - Al soltar el mouse después de un arrastre, copiar el valor
        #    de la celda actual al resto de celdas seleccionadas de esa fila.
        if source is self.schedule_table.viewport():
            if event.type() == QEvent.Type.Leave:
                self._shift_info_card.hide()
            elif (
                event.type() == QEvent.Type.MouseButtonRelease
                and event.button() == Qt.MouseButton.LeftButton
            ):
                self._apply_fill_from_anchor()
        return super().eventFilter(source, event)

    # ---------- registration form (compact) ----------
    def _build_registration_form(self) -> QWidget:
        container = QWidget()
        self.main_form_layout = QVBoxLayout(container)
        
        # 1. MODIFICACIÓN: Márgenes mínimos y espaciado reducido
        self.main_form_layout.setContentsMargins(2, 4, 2, 2)
        self.main_form_layout.setSpacing(4) 

        # --- Inicialización de controles (Lógica original intacta) ---
        self.user_selector_combo = QComboBox()
        self.user_selector_combo.currentIndexChanged.connect(self.autofill_user_data)
        self.user_selector_combo.setFixedHeight(26) # Altura forzada compacta

        self.role_display = QLineEdit()
        self.role_display.setReadOnly(True)
        self.role_display.setStyleSheet("background-color: #F3F4F6; color: #6B7280;")
        self.role_display.setFixedHeight(26)

        self.badge_display = QLineEdit()
        self.badge_display.setReadOnly(True)
        self.badge_display.setStyleSheet("background-color: #F3F4F6; color: #6B7280;")
        self.badge_display.setFixedHeight(26)

        self.status_selector = QComboBox()
        self.status_selector.setFixedHeight(26)
        self.status_selector.currentIndexChanged.connect(self._on_register_status_changed)
        self.start_date_edit = QDateEdit(QDate.currentDate())
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDisplayFormat("dd/MM/yyyy")
        self.start_date_edit.setFixedHeight(26)

        self.end_date_edit = QDateEdit(QDate.currentDate().addDays(7))
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setDisplayFormat("dd/MM/yyyy")
        self.end_date_edit.setFixedHeight(26)

        self.pickup_combo = QComboBox()
        self.pickup_combo.setFixedHeight(26)
        self.dropoff_combo = QComboBox()
        self.dropoff_combo.setFixedHeight(26)
        
        self.remarks_input = QLineEdit()
        self.remarks_input.setPlaceholderText("Optional remarks...")
        self.remarks_input.setFixedHeight(26)

        # Checkbox auxiliar lógico (mantenido pero oculto)
        self.apply_to_range_chk = QCheckBox("Apply to all selected cells")
        self.apply_to_range_chk.setVisible(False)

        # 2. MODIFICACIÓN: Botón más compacto
        self.save_button = QPushButton("💾 Save Changes") 
        self.save_button.clicked.connect(self.save_plan_changes)
        self.save_button.setProperty("variant", "primary")
        self.save_button.setFixedSize(120, 28) # Tamaño fijo y pequeño

        # Toggle de fechas de viaje
        self.travel_dates_check = QCheckBox("Travel dates ≠ Work period")
        self.travel_dates_check.toggled.connect(self._toggle_travel_dates_visibility)
        self.travel_dates_check.setStyleSheet("font-size: 11px;")

        # --- Contenedor de Viaje (Entry/Exit) ---
        self.entry_date_edit = QDateEdit(QDate.currentDate())
        self.entry_date_edit.setCalendarPopup(True)
        self.entry_date_edit.setDisplayFormat("dd/MM")
        self.entry_time_edit = QTimeEdit(QTime(6, 0))
        self.entry_time_edit.setDisplayFormat("HH:mm")

        self.exit_date_edit = QDateEdit(QDate.currentDate().addDays(14))
        self.exit_date_edit.setCalendarPopup(True)
        self.exit_date_edit.setDisplayFormat("dd/MM")
        self.exit_time_edit = QTimeEdit(QTime(18, 0))
        self.exit_time_edit.setDisplayFormat("HH:mm")

        self._travel_dates_container = QWidget()
        self._travel_dates_container.setStyleSheet("background-color: #F9FAFB; border: 1px solid #E5E7EB; border-radius: 4px;")
        travel_layout = QHBoxLayout(self._travel_dates_container)
        travel_layout.setContentsMargins(4, 2, 4, 2)
        travel_layout.setSpacing(6)
        
        # Widgets de viaje en una sola línea compacta
        travel_layout.addWidget(QLabel("IN:"))
        travel_layout.addWidget(self.entry_date_edit)
        travel_layout.addWidget(self.entry_time_edit)
        travel_layout.addWidget(QLabel("|"))
        travel_layout.addWidget(QLabel("OUT:"))
        travel_layout.addWidget(self.exit_date_edit)
        travel_layout.addWidget(self.exit_time_edit)
        travel_layout.addStretch()
        
        self._travel_dates_container.setVisible(False)

        # 3. MODIFICACIÓN: Helper 'field' optimizado para altura mínima
        def field(title: str, w: QWidget) -> QWidget:
            cont = QWidget()
            v = QVBoxLayout(cont)
            v.setContentsMargins(0, 0, 0, 0)
            v.setSpacing(0) # Pegar etiqueta al input
            
            lbl = QLabel(title)
            # Fuente pequeña para ahorrar espacio vertical
            lbl.setStyleSheet("font-size: 10px; color: #4B5563; font-weight: 700; margin-bottom: 1px;")
            
            v.addWidget(lbl)
            v.addWidget(w)
            return cont

        # Array de campos (CRÍTICO: Mantiene la lógica de redimensionado)
        self._fields = [
            field("Employee", self.user_selector_combo),
            field("Role / Dept", self.role_display),
            field("Badge ID", self.badge_display),
            field("Status / Shift", self.status_selector),
            field("Start Date", self.start_date_edit),
            field("End Date", self.end_date_edit),
            field("Pick Up", self.pickup_combo),
            field("Drop Off", self.dropoff_combo),
            field("Remarks", self.remarks_input),
        ]

        # Grid layout base
        self._register_grid = QGridLayout()
        self._register_grid.setContentsMargins(0, 0, 0, 0)
        self._register_grid.setHorizontalSpacing(8)
        self._register_grid.setVerticalSpacing(2) # Espacio vertical mínimo entre filas
        
        self.main_form_layout.addLayout(self._register_grid)

        # 4. MODIFICACIÓN: Fila inferior unificada (Checkbox + Botón en la misma línea)
        bottom_row = QHBoxLayout()
        bottom_row.setContentsMargins(0, 4, 0, 0)
        bottom_row.setSpacing(10)
        
        bottom_row.addWidget(self.travel_dates_check)
        bottom_row.addStretch()
        bottom_row.addWidget(self.save_button) # Botón a la derecha, misma fila
        
        # El contenedor de viaje va ANTES de la fila inferior para no romper la alineación
        self.main_form_layout.addWidget(self._travel_dates_container)
        self.main_form_layout.addLayout(bottom_row)

        self.load_shift_type_options()

        return container

    def _toggle_travel_dates_visibility(self, checked: bool):
        """Shows or hides the entry/exit date fields based on the checkbox."""
        self._travel_dates_container.setVisible(checked)
        if checked:
            # For convenience, sync the dates from the main period when shown
            self.entry_date_edit.setDate(self.start_date_edit.date())
            self.exit_date_edit.setDate(self.end_date_edit.date())
            # MODIFIED: Set default times as well
            self.entry_time_edit.setTime(QTime(6, 0))
            self.exit_time_edit.setTime(QTime(18, 0))

    def _rebuild_registration_grid(self, columns: int):
        if columns < 1:
            columns = 1
        num_fields = len(self._fields)

        if self._register_grid is None:
            return
        if self._current_form_cols == columns and self._register_grid.count() > 0:
            return
        self._current_form_cols = columns

        # Clear grid
        while self._register_grid.count():
            item = self._register_grid.takeAt(0)
            if item.widget():
                item.widget().setParent(None)

        # Re-add fields
        rows = (num_fields + columns - 1) // columns
        idx = 0
        for r in range(rows):
            for c in range(columns):
                if idx < num_fields:
                    self._register_grid.addWidget(self._fields[idx], r, c)
                    self._register_grid.setColumnStretch(c, 1)
                    idx += 1

    def resizeEvent(self, event):
        super().resizeEvent(event)
        # Simple responsive thresholds for 9 fields (3x3 is ideal)
        w = max(0, self.width())
        cols = 3 if w >= 930 else (2 if w >= 620 else 1)
        if cols != self._current_form_cols:
            self._rebuild_registration_grid(cols)

    # ---------- data loaders ----------

    def load_shift_type_options(self):
        """Load base statuses + custom shift types (from DB) into the combo."""
        self.status_selector.blockSignals(True)
        self.status_selector.clear()

        # Blank option
        self.status_selector.addItem(QIcon(), "— Do Not Mark Days —", {"kind": "none"})

        # Base statuses with color chips
        self.status_selector.addItem(
            _create_color_icon("#FFC7CE"),  # OFF color
            "OFF",
            {
                "kind": "base",
                "status": "OFF",
                "shift_type": None,
                "in_time": None,
                "out_time": None,
                "is_off": True,
            },
        )
        self.status_selector.addItem(
            _create_color_icon("#C6EFCE"),  # ON color
            "ON (Day Shift)",
            {
                "kind": "base",
                "status": "ON",
                "shift_type": "Day Shift",
                "in_time": None,
                "out_time": None,
            },
        )
        self.status_selector.addItem(
            _create_color_icon("#FFFF99"),  # ON NS color
            "ON NS (Night Shift)",
            {
                "kind": "base",
                "status": "ON NS",
                "shift_type": "Night Shift",
                "in_time": None,
                "out_time": None,
            },
        )

        # Custom types
        types = db.get_shift_types(self.source)
        self._custom_shift_map = {t["code"]: t for t in types}  # refresh map
        if types:
            self.status_selector.addItem(
                QIcon(), "—— Custom Shift Types ——", {"kind": "separator"}
            )
            for t in types:
                label = f"{t['name']} [{t['code']}]  {t['in_time']}-{t['out_time']}"
                self.status_selector.addItem(
                    _create_color_icon(t["color_hex"]),
                    label,
                    {
                        "kind": "custom",
                        "code": t["code"],
                        "name": t["name"],
                        "in_time": t["in_time"],
                        "out_time": t["out_time"],
                        "is_off": t.get("is_off", 0),
                    },
                )
        self.status_selector.setCurrentIndex(0)
        self.status_selector.blockSignals(False)

    def _status_options_for_dialog(self):
        """Devuelve la lista de opciones (icono, texto, payload) para los diálogos/combos."""
        options = []

        # Opción neutra (no marcar)
        options.append((QIcon(), "— Do Not Mark Days —", {"kind": "none"}))

        # Base: OFF / ON / ON NS
        options.append(
            (QIcon(), "OFF", {"kind": "base", "status": "OFF", "shift_type": None})
        )
        options.append(
            (QIcon(), "ON (Day)", {"kind": "base", "status": "ON", "shift_type": None})
        )
        options.append(
            (
                QIcon(),
                "ON NS (Night)",
                {"kind": "base", "status": "ON NS", "shift_type": None},
            )
        )

        # Tipos de turno personalizados (self._custom_shift_map ya existe)
        for code, info in (self._custom_shift_map or {}).items():
            display_name = info.get("name") or code
            options.append(
                (
                    QIcon(),
                    display_name,
                    {"kind": "custom", "status": "ON", "shift_type": code},
                )
            )

        return options

    def _status_options_for_dialog(self):
        """
        Devuelve la misma lista de opciones que el combo Status/Shift
        para usarla en el editor de celda y en el diálogo DayScheduleEditor.
        """
        options = []
        for i in range(self.status_selector.count()):
            data = self.status_selector.itemData(i)
            # Saltamos separadores u opciones sin payload
            if not data or data.get("kind") == "separator":
                continue
            icon = self.status_selector.itemIcon(i)
            text = self.status_selector.itemText(i)
            options.append((icon, text, data))
        return options

    # NEW: load available locations into dropdowns
    def load_location_options(self):
        self.pickup_combo.blockSignals(True)
        self.dropoff_combo.blockSignals(True)
        self.pickup_combo.clear()
        self.dropoff_combo.clear()
        self.pickup_combo.addItem("— Select location —", None)
        self.dropoff_combo.addItem("— Select location —", None)
        for loc in db.get_locations(self.source):
            self.pickup_combo.addItem(loc["pickup_location"], loc["pickup_location"])
            self.dropoff_combo.addItem(loc["pickup_location"], loc["pickup_location"])
        self.pickup_combo.setCurrentIndex(0)
        self.dropoff_combo.setCurrentIndex(0)
        self.pickup_combo.blockSignals(False)
        self.dropoff_combo.blockSignals(False)

    def refresh_ui_data(self, use_preloaded=False):
        self.load_shift_type_options()
        self.load_schedule_data(use_preloaded=use_preloaded) # Pasar la bandera
        self.load_users_to_selector()
        self.load_location_options()  # keep combos in sync with Location admin
        self.remarks_input.clear()  # Clear remarks on refresh

    # [CAMBIO 5] Lógica crítica para usar los datos en memoria
    def load_schedule_data(self, use_preloaded=False):
        # -------------------------------------------------------------
        # LÓGICA DE PRE-CARGA (WORKER THREAD)
        # -------------------------------------------------------------
        df = None
        if use_preloaded and hasattr(self, '_initial_preloaded_df') and self._initial_preloaded_df is not None:
            print("DEBUG: Using preloaded dataframe (High Performance Mode).")
            df = self._initial_preloaded_df
            self._initial_preloaded_df = None 
        else:
            df = excel.get_schedule_preview(self.excel_file)
        
        try:
            if df is not None and not df.empty:
                # Importamos users desde la BD para asegurar datos frescos
                users_db = db.get_all_users(self.source)
                
                # Creamos mapas de Badge -> Role y Badge -> Name
                role_map = {str(u["badge"]).strip(): str(u.get("role") or "").strip() for u in users_db}
                name_map = {str(u["badge"]).strip(): str(u.get("name") or "").strip() for u in users_db}

                # Determinamos nombre de columna Badge (RGM usa BADGE, Newmont usa Company ID)
                col_badge = "BADGE" if "BADGE" in df.columns else "Company ID"
                
                if col_badge in df.columns:
                    # Normalizamos columna Badge para cruce exacto
                    df[col_badge] = df[col_badge].astype(str).str.strip()
                    
                    # Sobreescribir ROL (Discipline o ROLE)
                    col_role = "ROLE" if "ROLE" in df.columns else "Discipline"
                    if col_role in df.columns:
                        # .map busca el badge en role_map; .fillna mantiene el valor original si no lo encuentra
                        df[col_role] = df[col_badge].map(role_map).fillna(df[col_role])

                    # Sobreescribir NOMBRE (Solo si existe columna simple NAME, Newmont usa First/Last separados y es más complejo)
                    if "NAME" in df.columns:
                        df["NAME"] = df[col_badge].map(name_map).fillna(df["NAME"])
                        
                    print(f" UI Overlay applied: User metadata synced with DB for display.")
        except Exception as e:
            print(f" Warning: Could not apply DB overlay to preview: {e}")
        
        self._loading_preview = True
        self._cell_original_values.clear()
        self._row_identities.clear()
        self._date_col_dates.clear()

        if df is None or df.empty:
            self.frozen_table.clear()
            self.schedule_table.clear()
            self.frozen_table.setRowCount(0)
            self.schedule_table.setRowCount(0)
            self._loading_preview = False
            return

        if not use_preloaded:
            try:
                ok_horizon, msg_horizon = excel.ensure_rolling_horizon_columns(
                    self.excel_file
                )
                if ok_horizon:
                    print(f"[AUTO-HORIZON] {msg_horizon}")
            except Exception as e:
                print(f"[AUTO-HORIZON CRITICAL] Failed: {e}")

        custom_map = db.get_shift_type_map(self.source)

        # -----------------------------------------------------------
        # ORDENAMIENTO Y ESTRUCTURA
        # -----------------------------------------------------------
        all_cols = list(df.columns)
        actual_frozen_count = min(len(all_cols), FROZEN_COLUMN_COUNT)

        frozen_part = all_cols[:actual_frozen_count]
        date_part = all_cols[actual_frozen_count:]

        date_mapping = []
        for c in date_part:
            d_obj = None
            if hasattr(c, "to_pydatetime"):
                d_obj = c.to_pydatetime().date()
            elif isinstance(c, datetime):
                d_obj = c.date()
            elif isinstance(c, pydate):
                d_obj = c

            if d_obj:
                date_mapping.append((d_obj, c))
            else:
                date_mapping.append((pydate.max, c))

        date_mapping.sort(key=lambda x: x[0])
        sorted_date_cols = [x[1] for x in date_mapping]
        new_column_order = frozen_part + sorted_date_cols
        df = df[new_column_order]

        cols = list(df.columns)
        date_cols = []
        for c in cols:
            if hasattr(c, "to_pydatetime"):
                date_cols.append(c.to_pydatetime().date())
            elif isinstance(c, datetime):
                date_cols.append(c.date())
            else:
                pass

        # Frozen Headers
        actual_frozen_count = min(df.shape[1], FROZEN_COLUMN_COUNT)
        frozen_headers = [str(c) for c in cols[:actual_frozen_count]]

        # Schedule Headers (Compacto: YYYY-MM-DD \n Dia)
        schedule_headers = []
        for d in date_cols:
            # Usamos %a (Mon, Tue) para que sea corto
            schedule_headers.append(f"{d.isoformat()}\n{d.strftime('%a')}")
        self._date_col_dates = list(date_cols)
        self.weekend_header.set_dates(self._date_col_dates)

        # Configurar Tablas
        self.frozen_table.setRowCount(df.shape[0])
        self.frozen_table.setColumnCount(actual_frozen_count)
        self.frozen_table.setHorizontalHeaderLabels(frozen_headers)

        self.schedule_table.setRowCount(df.shape[0])
        self.schedule_table.setColumnCount(len(schedule_headers))
        for idx, header_text in enumerate(schedule_headers):
            self.schedule_table.setHorizontalHeaderItem(
                idx, QTableWidgetItem(header_text)
            )

        # Cargar Filas
        for i, row in df.iterrows():
            badge_val = row.get("BADGE") if hasattr(row, "get") else (row["BADGE"] if "BADGE" in df.columns else "")
            name_val = row.get("NAME") if hasattr(row, "get") else (row["NAME"] if "NAME" in df.columns else "")
            role_val = row.get("ROLE") if hasattr(row, "get") else (row["ROLE"] if "ROLE" in df.columns else "")

            self._row_identities.append(
                {
                    "badge": str(badge_val) if badge_val is not None else "",
                    "name": str(name_val) if name_val is not None else "",
                    "role": str(role_val) if role_val is not None else "",
                }
            )

            for j, val in enumerate(row):
                text = _clean(val)
                item = QTableWidgetItem(text)

                if j < actual_frozen_count:
                    # Frozen table
                    item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    self.frozen_table.setItem(i, j, item)
                else:
                    # Schedule table
                    col_index = j - actual_frozen_count
                    val_str = text.upper().strip()
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

                    if "ON NS" in val_str or "NIGHT" in val_str:
                        item.setBackground(QColor("#FFFF99"))
                    elif val_str == "ON" or "DAY" in val_str or val_str.isdigit():
                        item.setBackground(QColor("#C6EFCE"))
                    elif val_str in ("OFF", "BREAK", "KO", "LEAVE"):
                        item.setBackground(QColor("#FFC7CE"))
                    else:
                        if val_str in custom_map and custom_map[val_str].get("color_hex"):
                            item.setBackground(QColor(custom_map[val_str]["color_hex"]))

                    self.schedule_table.setItem(i, col_index, item)
                    self._cell_original_values[(i, col_index)] = val_str

                    key = self._warn_key_for(i, col_index)
                    if key in self._warn_highlight_keys:
                        item.setBackground(QColor(WARN_BG_HEX))

        # -------------------------------------------------------------
        # CORRECCIÓN DE ESTILO Y VISIBILIDAD (ELIMINAR PADDING)
        # -------------------------------------------------------------
        
        # 1. Ajustar columnas al contenido
        self.frozen_table.resizeColumnsToContents()
        self.schedule_table.resizeColumnsToContents()

        # 2. Definir altura compacta (35px es suficiente si quitamos el padding)
        compact_height = 35
        
        # 3. Fuente pequeña y negrita
        compact_font = QFont()
        compact_font.setPointSize(7)
        compact_font.setBold(True)

        # 4. TRUCO DE INGENIERÍA: Hoja de estilos para eliminar el padding interno
        # Esto permite que el texto use todo el espacio vertical disponible.
        # Ajustamos RGM/Newmont colors si quisieras, pero aquí usamos blanco/básico para legibilidad.
        # NOTA: Ajusta 'background-color' si usas un tema oscuro o corporativo específico.
        header_stylesheet = """
            QHeaderView::section {
                padding-top: 0px;
                padding-bottom: 0px;
                padding-left: 2px;
                padding-right: 2px;
                margin: 0px;
                border-bottom: 1px solid #ccc;
                border-right: 1px solid #ccc;
            }
        """

        # Aplicar a Tabla DERECHA (Fechas)
        h_sched = self.schedule_table.horizontalHeader()
        h_sched.setFont(compact_font)
        h_sched.setFixedHeight(compact_height)
        h_sched.setStyleSheet(header_stylesheet)
        h_sched.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)

        # Aplicar a Tabla IZQUIERDA (Nombres) - Exactamente igual para alineación perfecta
        h_frozen = self.frozen_table.horizontalHeader()
        h_frozen.setFont(compact_font)
        h_frozen.setFixedHeight(compact_height)
        h_frozen.setStyleSheet(header_stylesheet)
        h_frozen.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)

        self._update_frozen_width()
        self._loading_preview = False
        self._center_today_column()

    def _center_today_column(self):
        """Scroll horizontally so today's date is visible and centered (REQ-002)."""
        try:
            if not self._date_col_dates:
                return
            today = QDate.currentDate().toPyDate()
            if today in self._date_col_dates:
                col = self._date_col_dates.index(today)
                if self.schedule_table.rowCount() > 0:
                    self.schedule_table.scrollToItem(
                        self.schedule_table.item(0, col),
                        QAbstractItemView.ScrollHint.PositionAtCenter,
                    )
        except Exception:
            pass

    def _update_frozen_width(self):
        """
        Compute and lock the exact width needed by the left (frozen) panel so
        ROLE, NAME, and BADGE are fully visible without horizontal scrolling.
        """
        try:
            # Make sure columns have been measured
            self.frozen_table.resizeColumnsToContents()

            vheader_w = self.frozen_table.verticalHeader().width()
            frame_w = self.frozen_table.frameWidth() * 2
            columns_w = sum(
                self.frozen_table.columnWidth(c)
                for c in range(self.frozen_table.columnCount())
            )
            padding = 6
            total = vheader_w + frame_w + columns_w + padding
            if total < 240:
                total = 240

            self.frozen_table.setMinimumWidth(total)
            self.frozen_table.setMaximumWidth(total)
        except Exception:
            pass

    def showEvent(self, event):
        super().showEvent(event)
        # Also center on show (e.g. when user navigates to the tab)
        self._center_today_column()
        self._update_frozen_width()

    def _warn_key_for(self, row: int, col: int) -> str:
        """Build a stable session key for a schedule cell using badge + date."""
        badge = ""
        if 0 <= row < len(self._row_identities):
            badge = self._row_identities[row].get("badge", "") or self._row_identities[
                row
            ].get("name", "")
        d = ""
        if 0 <= col < len(self._date_col_dates):
            d = self._date_col_dates[col].isoformat()
        return f"{badge}|{d}"

    def _apply_base_background(self, item: QTableWidgetItem, value_upper: str):
        """Apply default background based on the cell value."""
        if value_upper == "ON":
            item.setBackground(QColor("#C6EFCE"))
        elif value_upper in ("ON NS", "NIGHT"):
            item.setBackground(QColor("#FFFF99"))
        elif value_upper in ("OFF", "BREAK", "KO", "LEAVE"):
            item.setBackground(QColor("#FFC7CE"))
        elif value_upper == "":
            item.setBackground(QColor(255, 255, 255, 0))  # transparent/no fill
        else:
            # leave as-is (could be a custom code already colored on load)
            pass

    def _apply_base_background(self, item: QTableWidgetItem, value_upper: str):
        """Apply default background based on the cell value."""
        if value_upper == "ON":
            item.setBackground(QColor("#C6EFCE"))
        elif value_upper in ("ON NS", "NIGHT"):
            item.setBackground(QColor("#FFFF99"))
        elif value_upper in ("OFF", "BREAK", "KO", "LEAVE"):
            item.setBackground(QColor("#FFC7CE"))
        elif value_upper == "":
            item.setBackground(QColor(255, 255, 255, 0))  # transparent/no fill
        else:
            # leave as-is (could be a custom code already colored on load)
            pass
        
        
    # --------------------------------------------------------------------------
    # MÉTODOS AUXILIARES PARA SMART DRAG & CONSOLIDATION
    # --------------------------------------------------------------------------

    def _calculate_time_logic(self, status, kind, raw_time=None):
        """
        CORREGIDO: Ahora busca primero en los Custom Shift Types cargados en memoria
        (self._custom_shift_map), recuperando la funcionalidad original.
        """
        from datetime import time as dtime
        
        # Validación de seguridad
        if not status:
            return dtime(0, 0)
        
        status_key = status.strip().upper()

        # ---------------------------------------------------------
        # 1. PRIORIDAD: Reglas de Negocio Hardcoded (Newmont / RGM)
        # ---------------------------------------------------------
        if self.source == "Newmont":
            if status_key == "ON": 
                return dtime(6, 0) if kind == "IN" else dtime(12, 0)
            elif status_key == "ON NS": 
                return dtime(12, 0) if kind == "IN" else dtime(6, 0)
        
        elif self.source == "RGM":
            if status_key in ("ON", "ON NS"): 
                return dtime(7, 0)

        # ---------------------------------------------------------
        # 2. PRIORIDAD: Turnos Personalizados (LA LÓGICA RESTAURADA)
        # ---------------------------------------------------------
        # Aquí consultamos el mapa que ya cargaste al inicio en load_shift_types
        if status_key in self._custom_shift_map:
            shift_info = self._custom_shift_map[status_key]
            # Extraer hora según sea Entrada o Salida
            time_str = shift_info.get("in_time") if kind == "IN" else shift_info.get("out_time")
            
            # Usar tu helper existente para convertir string a objeto time
            return _parse_hhmm_to_time(time_str, dtime(0, 0))

        # ---------------------------------------------------------
        # 3. FALLBACK: Datos Crudos del Arrastre
        # ---------------------------------------------------------
        # Si no está en el mapa, intentamos usar lo que venía de la celda origen
        return _parse_hhmm_to_time(raw_time, dtime(0, 0))
    
    def _consolidate_and_record_logistics(self, badge, role, username, op_start, op_end, status, cursor=None):
        """
        EL MOTOR DE CONSOLIDACIÓN:
        1. Verifica si la operación nueva se toca con operaciones existentes (ayer/mañana).
        2. Si se tocan y son trabajo continuo, las fusiona en un solo bloque.
        3. Recalcula las horas de Entrada (Entry) y Salida (Exit) basadas en el PRIMER y ÚLTIMO día.
        4. Escribe la Operación Maestra en la BD.
        """
        # A. Si es un día libre (OFF), NO consolidamos operaciones, solo limpiamos.
        if not db.is_working_status(status, self.source):
            db.delete_operations_in_range(badge, op_start, op_end, cursor=cursor) ### <--- CAMBIO AQUÍ
            return

        final_start = op_start
        final_end = op_end

        # B. Fusión Izquierda (Looking Back - Ayer)
        prev_day = final_start - timedelta(days=1)
        prev_op = db.get_operation_overlapping(badge, prev_day, cursor=cursor)
        
        if prev_op:
            # Doble check: Asegurar que el día anterior en el calendario (schedule) es trabajo
            prev_map = db.get_schedule_map_for_range(badge, prev_day, prev_day, self.source, cursor=cursor)
            prev_st = prev_map.get(prev_day.isoformat(), {}).get("status")
            
            if db.is_working_status(prev_st, self.source):
                # ¡Fusión! Extendemos el inicio al inicio de la operación anterior
                prev_op_start = datetime.strptime(prev_op['start_date'], "%Y-%m-%d").date()
                if prev_op_start < final_start:
                    final_start = prev_op_start

        # C. Fusión Derecha (Looking Forward - Mañana)
        next_day = final_end + timedelta(days=1)
        next_op = db.get_operation_overlapping(badge, next_day, cursor=cursor)
        
        if next_op:
            # Doble check: Asegurar que el día siguiente en el calendario es trabajo
            next_map = db.get_schedule_map_for_range(badge, next_day, next_day, self.source,cursor=cursor)
            next_st = next_map.get(next_day.isoformat(), {}).get("status")
            
            if db.is_working_status(next_st, self.source):
                # ¡Fusión! Extendemos el final al final de la operación siguiente
                next_op_end = datetime.strptime(next_op['end_date'], "%Y-%m-%d").date()
                if next_op_end > final_end:
                    final_end = next_op_end

        # D. Limpieza: Borrar cualquier operación fragmentada en el nuevo rango maestro
        db.delete_operations_in_range(badge, final_start, final_end, cursor=cursor)

        # E. Recálculo Inteligente de Horarios (Entry/Exit)
        # Usamos el status del PRIMER día para la Entry Date
        map_start = db.get_schedule_map_for_range(badge, final_start, final_start, self.source, cursor=cursor)
        st_start = map_start.get(final_start.isoformat(), {}).get("status") or status
        t_in = self._calculate_time_logic(st_start, "IN")

        # Usamos el status del ÚLTIMO día para la Exit Date
        map_end = db.get_schedule_map_for_range(badge, final_end, final_end, self.source, cursor=cursor)
        st_end = map_end.get(final_end.isoformat(), {}).get("status") or status
        t_out = self._calculate_time_logic(st_end, "OUT")

        # F. Inserción de la Operación Unificada
        db.add_operation(
            username=username, 
            role=role, 
            badge=badge,
            start_date=final_start, 
            end_date=final_end,
            created_by=self.logged_username,
            entry_date=datetime.combine(final_start, t_in),
            exit_date=datetime.combine(final_end, t_out),
            cursor=cursor  # <--- CRÍTICO: Pasar el cursor
        )
        # print(f"DEBUG: Consolidated Op: {final_start} -> {final_end}")

    
    def _apply_fill_from_anchor(self):
        """
        Excel-style Fill Right: Copia la celda izquierda (ancla) hacia la derecha.
        
        VERSIÓN MEJORADA: SMART DRAG & CONSOLIDATION
        1. Mantiene tu lógica de selección y corrección de fecha de inicio (Off-By-One).
        2. Usa _calculate_time_logic para estandarizar reglas Newmont/RGM.
        3. Invoca _consolidate_and_record_logistics para fusionar bloques adyacentes.
        """
        # 1. Obtener rangos seleccionados
        selected_ranges = self.schedule_table.selectedRanges()
        if not selected_ranges:
            return

        # Bloquear señales para optimizar rendimiento visual
        self._bulk_editing = True
        
        try:
            # Iterar por cada bloque de selección
            for r_range in selected_ranges:
                top_row = r_range.topRow()
                bottom_row = r_range.bottomRow()
                left_col = r_range.leftColumn()
                right_col = r_range.rightColumn()

                # Si es una sola columna, no hay relleno horizontal
                if left_col == right_col:
                    continue

                # Procesar fila por fila (Empleado por Empleado)
                for r in range(top_row, bottom_row + 1):
                    # --- A. Identificar la FUENTE (Ancla - Celda Izquierda) ---
                    source_item = self.schedule_table.item(r, left_col)
                    base_text = (source_item.text() or "").strip().upper() if source_item else ""
                    
                    # Validar identidad del empleado
                    if not (0 <= r < len(self._row_identities)):
                        continue
                    
                    identity = self._row_identities[r]
                    badge = identity.get("badge")
                    username = identity.get("name")
                    role = identity.get("role")
                    
                    if not badge:
                        continue

                    # --- B. Recuperar Metadatos de la Fuente (DB) ---
                    anchor_date = self._date_col_dates[left_col]
                    schedule_map = db.get_schedule_map_for_range(
                        badge, anchor_date, anchor_date, self.source
                    )
                    source_info = schedule_map.get(anchor_date.isoformat()) or {}

                    # Datos base a propagar
                    new_status = (source_info.get("status") or base_text).upper()
                    new_shift_type = source_info.get("shift_type")
                    new_remark = source_info.get("remark")
                    
                    # Recuperar horas crudas de la fuente
                    raw_in_time = source_info.get("in_time")
                    raw_out_time = source_info.get("out_time")

                    # [TUS CORRECCIONES DE FECHAS SE MANTIENEN]
                    # Rango visual (donde pintamos): desde la siguiente columna (left_col + 1)
                    start_fill_date = self._date_col_dates[left_col + 1]
                    end_fill_date = self._date_col_dates[right_col]
                    
                    # Rango Lógico (Operación): INCLUYE el día ancla (left_col)
                    # Esto asegura que la operación logística arranque el día que seleccionaste, no el siguiente.
                    operation_start_date = self._date_col_dates[left_col]

                    # --- C. REGLA DE NEGOCIO: Determinación de Horarios (INTEGRADA) ---
                    # Usamos el nuevo helper para garantizar consistencia con la consolidación
                    final_in_obj = self._calculate_time_logic(new_status, "IN", raw_in_time)
                    final_out_obj = self._calculate_time_logic(new_status, "OUT", raw_out_time)

                    # Convertir a string para DB (HH:MM)
                    new_in_time_str = final_in_obj.strftime("%H:%M")
                    new_out_time_str = final_out_obj.strftime("%H:%M")

                    # --- D. Actualizar UI Visualmente (Solo las celdas nuevas) ---
                    for c in range(left_col + 1, right_col + 1):
                        item = self.schedule_table.item(r, c)
                        if not item:
                            item = QTableWidgetItem()
                            self.schedule_table.setItem(r, c, item)
                        
                        item.setText(new_status)
                        self._apply_base_background(item, new_status)
                        self._cell_original_values[(r, c)] = new_status

                    # --- E. Persistencia ---
                    # --- E. Persistencia ---
                    try:
                        # [PATCH COLISIÓN] ----------------------------------------
                        # Verificar si el inicio del relleno choca con el día anterior
                        force_flag = self._resolve_force_new_entry_start(badge, start_fill_date, new_status)
                        
                        if force_flag is None:
                            # Usuario canceló en el popup -> Abortar este empleado
                            continue 
                        # ---------------------------------------------------------

                        # 1. Guardar Schedule (SSoT) - Día a día
                        # Se guardan las celdas rellenadas (del 22 en adelante)
                        db.upsert_schedule_range(
                            badge, 
                            start_fill_date, 
                            end_fill_date, 
                            new_status, 
                            new_shift_type, 
                            self.source, 
                            new_in_time_str, 
                            new_out_time_str, 
                            new_remark,
                            force_new_entry_start=force_flag  # <--- NUEVO ARGUMENTO
                        )

                        # 2. Guardar Excel - Fila Física
                        self._is_internal_update = True 
                        excel.update_plan_staff_excel(
                            self.excel_file,
                            username,
                            role,
                            badge,
                            new_status,
                            new_shift_type,
                            start_fill_date,
                            end_fill_date,
                            self.source,
                            new_in_time_str,
                            new_out_time_str,
                        )

                        # 3. [MODIFICACIÓN CLAVE] CONSOLIDACIÓN LOGÍSTICA (Operations)
                        # En lugar de guardar ciegamente, llamamos al consolidador.
                        # Le pasamos operation_start_date (el día 21) y end_fill_date (el día 25).
                        # Él se encargará de ver si hay que fusionar con el 20 o el 26.
                        self._consolidate_and_record_logistics(
                            badge, 
                            role, 
                            username,
                            operation_start_date, # INCLUYE EL ANCLA
                            end_fill_date,
                            new_status
                        )

                    except Exception as e:
                        print(f"Error saving drag-fill for {badge}: {e}")

        finally:
            self._is_internal_update = False
            self._bulk_editing = False
            self.rotation_changed.emit()
            self.schedule_table.viewport().update()
    
    
    # ---------- REQ-001: inline OFF→ON/ON NS guard ----------
    def _on_schedule_cell_changed(self, item: QTableWidgetItem):
        # 1. AGREGAR: Chequeo de la bandera _is_handling_change
        if (
            self._loading_preview
            or self._bulk_editing
            or getattr(self, "_is_handling_change", False)
        ):
            return

        # 2. ACTIVAR BANDERA
        self._is_handling_change = True

        try:
            r = item.row()
            c = item.column()
            new_text = (item.text() or "").strip().upper()
            old_text = (
                (self._cell_original_values.get((r, c), "") or "").strip().upper()
            )

            # ---------------- REQ-001: OFF/blank -> ON / ON NS ----------------
            if old_text in ("OFF", "") and new_text in ("ON", "ON NS"):
                box = QMessageBox(self)
                box.setIcon(QMessageBox.Icon.Warning)
                box.setWindowTitle("Confirm Change")
                box.setText(
                    "The employee is on a day off. Do you want to set it to ON?"
                )
                accept_btn = box.addButton("Accept", QMessageBox.ButtonRole.AcceptRole)
                box.addButton("Cancel", QMessageBox.ButtonRole.RejectRole)
                box.exec()
                if box.clickedButton() != accept_btn:
                    # Revertir a valor original y color base
                    with QSignalBlocker(self.schedule_table):
                        item.setText(old_text)
                    self._apply_base_background(item, old_text)
                    return
                else:
                    # Mantener el nuevo valor, marcar warn suave
                    with QSignalBlocker(self.schedule_table):
                        item.setText(new_text)  # normalizar mayúsculas
                    item.setBackground(QColor(WARN_BG_HEX))
                    self._warn_highlight_keys.add(self._warn_key_for(r, c))
            else:
                # Sin guard especial; color base y limpiar warn si aplica
                self._apply_base_background(item, new_text)
                key = self._warn_key_for(r, c)
                if new_text not in ("ON", "ON NS") and key in self._warn_highlight_keys:
                    self._warn_highlight_keys.discard(key)

            # ---------------- Identidad de fila / columna ----------------
            if r < 0 or r >= len(self._row_identities):
                return
            if c < 0 or c >= len(self._date_col_dates):
                return

            identity = self._row_identities[r]
            username = identity.get("name") or ""
            badge = identity.get("badge") or ""
            role = identity.get("role") or ""

            if not badge:
                # Sin badge no podemos guardar nada consistente
                return

            base_date = self._date_col_dates[c]

            # ---------------- Rango horizontal (para autofill) ----------------
            selected_ranges = self.schedule_table.selectedRanges()
            col_range = [c]
            if selected_ranges:
                sel = selected_ranges[0]
                if sel.topRow() <= r <= sel.bottomRow():
                    left = max(c, sel.leftColumn())
                    right = sel.rightColumn()
                    col_range = list(range(left, right + 1))

            # ---------------- Valores iniciales para el diálogo ----------------
            # Leemos info actual de BD (status, remark, pickup/dropoff) para el día base
            schedule_map = db.get_schedule_map_for_range(
                badge, base_date, base_date, self.source
            )
            day_info = schedule_map.get(base_date.isoformat(), {}) or {}

            # --- Lógica de resolución de código (MANTENER TU ARREGLO PREVIO) ---
            raw_cell_text = new_text if new_text else ""
            resolved_status_code = raw_cell_text

            possible_options = self._status_options_for_dialog()

            for icon, label, data in possible_options:
                if not isinstance(data, dict):
                    continue

                # 1. ¿Coincide con la etiqueta visual?
                if label.strip().upper() == raw_cell_text:
                    if data.get("kind") == "base":
                        resolved_status_code = data.get("status")
                    elif data.get("kind") == "custom":
                        resolved_status_code = data.get("code")
                    break

                # 2. ¿Coincide con el código interno?
                internal_code = (
                    data.get("status")
                    if data.get("kind") == "base"
                    else data.get("code")
                )
                if internal_code and str(internal_code).upper() == raw_cell_text:
                    resolved_status_code = internal_code
                    break

            current_status_text = resolved_status_code
            # -----------------------------------------------------------------

            pickup_init, dropoff_init = db.get_user_location_for_date(badge, base_date)
            initial = {
                "status_text": current_status_text,
                "pickup": pickup_init,
                "dropoff": dropoff_init,
                "remark": day_info.get("remark") or "",
            }

            locations = [
                loc["pickup_location"] for loc in db.get_locations(self.source)
            ]

            editor = DayScheduleEditor(
                self,
                status_options=self._status_options_for_dialog(),
                locations=locations,
                initial=initial,
            )

            # --- AQUÍ OCURRÍA EL DOBLE TRIGGER ---
            # Al ejecutarse editor.exec(), se pierde foco, se dispara itemChanged de nuevo.
            # Pero como _is_handling_change es True, la segunda llamada entra al 'if' inicial y retorna.
            if editor.exec() != QDialog.DialogCode.Accepted:
                # Usuario canceló: revertimos el cambio visual y salimos
                with QSignalBlocker(self.schedule_table):
                    item.setText(old_text)
                self._apply_base_background(item, old_text)
                return

            payload = editor.result_payload()
            sel = payload["selection"]
            pickup = payload["pickup"]
            dropoff = payload["dropoff"]
            remark = payload["remark"]
            apply_to_range = payload["apply_to_range"]

            entry_datetime = payload.get("entry_datetime")
            exit_datetime = payload.get("exit_datetime")

            # Validación simple
            if entry_datetime and exit_datetime and entry_datetime > exit_datetime:
                QMessageBox.warning(
                    self,
                    "Date Error",
                    "Entry date/time cannot be after Exit date/time.",
                )
                with QSignalBlocker(self.schedule_table):
                    item.setText(old_text)
                self._apply_base_background(item, old_text)
                return

            # Interpretar selección
            if not sel or sel.get("kind") in ("none", "separator"):
                schedule_status = None
                shift_type = None
                in_time = out_time = None
            elif sel.get("kind") == "base":
                schedule_status, shift_type = sel["status"], sel["shift_type"]
                in_time, out_time = sel.get("in_time"), sel.get("out_time")
            else:  # custom
                schedule_status, shift_type = sel["code"], sel["name"]
                in_time, out_time = sel.get("in_time"), sel.get("out_time")

            # Si "Do Not Mark Days"
            if schedule_status is None:
                with QSignalBlocker(self.schedule_table):
                    for cc in [c] if not apply_to_range else col_range:
                        it = self.schedule_table.item(r, cc)
                        if it is None:
                            it = QTableWidgetItem("")
                            self.schedule_table.setItem(r, cc, it)
                        it.setText("")
                        self._apply_base_background(it, "")
                return

            # Aplicar cambios (Fechas y Horas)
            from datetime import datetime, time as dtime

            cols_sorted = sorted(col_range if apply_to_range else [c])
            first_col_idx = cols_sorted[0]
            last_col_idx = cols_sorted[-1]

            start_date_block = self._date_col_dates[first_col_idx]
            end_date_block = self._date_col_dates[last_col_idx]

            if entry_datetime and exit_datetime:
                entry_dt_for_save = entry_datetime
                exit_dt_for_save = exit_datetime
            else:
                from datetime import time as dtime

                def_in = dtime(7, 0)
                def_out = dtime(7, 0)

                if self.source == "Newmont":
                    if schedule_status == "ON NS":
                        def_in = dtime(12, 0)
                        def_out = dtime(6, 0)
                    elif schedule_status == "ON":
                        def_in = dtime(6, 0)
                        def_out = dtime(12, 0)
                else:  # RGM
                    if schedule_status == "ON NS":
                        def_in = dtime(7, 0)
                        def_out = dtime(7, 0)

                entry_time_obj = _parse_hhmm_to_time(in_time, def_in)
                exit_time_obj = _parse_hhmm_to_time(out_time, def_out)

                entry_dt_for_save = datetime.combine(start_date_block, entry_time_obj)
                exit_dt_for_save = datetime.combine(end_date_block, exit_time_obj)

            # Actualizar UI visualmente
            with QSignalBlocker(self.schedule_table):
                for col_idx in cols_sorted:
                    item_ui = self.schedule_table.item(r, col_idx)
                    if not item_ui:
                        item_ui = QTableWidgetItem()
                        self.schedule_table.setItem(r, col_idx, item_ui)

                    text_to_show = schedule_status if schedule_status else ""
                    item_ui.setText(text_to_show)
                    self._apply_base_background(item_ui, text_to_show)
                    self._cell_original_values[(r, col_idx)] = text_to_show

            # Actualizar Base de Datos
            try:
                # [PATCH COLISIÓN] ----------------------------------------
                # start_date_block es la fecha donde inicia el cambio
                force_flag = self._resolve_force_new_entry_start(badge, start_date_block, schedule_status)
                
                if force_flag is None:
                    # Usuario canceló -> Revertir visualmente y salir
                    with QSignalBlocker(self.schedule_table):
                        item.setText(old_text)
                    self._apply_base_background(item, old_text)
                    return
                # ---------------------------------------------------------
                db.add_operation(
                    username=username,
                    role=role,
                    badge=badge,
                    start_date=start_date_block,
                    end_date=end_date_block,
                    created_by=self.logged_username,
                    entry_date=entry_dt_for_save,
                    exit_date=exit_dt_for_save,
                )

                db.upsert_schedule_range(
                    badge,
                    start_date_block,
                    end_date_block,
                    schedule_status,
                    shift_type,
                    self.source,
                    in_time,
                    out_time,
                    remark,
                    force_new_entry_start=force_flag
                )

                if pickup or dropoff:
                    db.assign_user_location_range(
                        badge, start_date_block, end_date_block, pickup, dropoff
                    )
                    db.log_event(
                        self.logged_username,
                        self.source,
                        "LOCATION_ASSIGN_INLINE",
                        f"{badge} {start_date_block}..{end_date_block} PU={pickup} DO={dropoff}",
                    )

                # Actualizar Excel
                self._is_internal_update = True
                success, message = excel.update_plan_staff_excel(
                    self.excel_file,
                    username,
                    role,
                    badge,
                    schedule_status,
                    shift_type,
                    start_date_block,
                    end_date_block,
                    self.source,
                    in_time,
                    out_time,
                )
                QTimer.singleShot(
                    2000, lambda: setattr(self, "_is_internal_update", False)
                )

                if not success:
                    self._is_internal_update = False
                    raise Exception(message)

                db.log_event(
                    self.logged_username,
                    self.source,
                    "SHIFT_MODIFICATION_INLINE",
                    f"Updated range {start_date_block} to {end_date_block} for {badge}. Remark: {remark}",
                )

            except Exception as e:
                self._is_internal_update = False
                QMessageBox.critical(self, "Save Error", f"Error saving data: {str(e)}")
                self.refresh_ui_data()
                return

            # Finalización
            self.check_excel_health()
            self.rotation_changed.emit()

        finally:
            # 3. LIBERAR BANDERA (CRÍTICO)
            self._is_handling_change = False

    # ---------- MODIFIED: hover card logic ----------
   # ---------- MODIFIED: hover card logic ----------
    def _show_shift_tooltip(self, row: int, col: int):
        """
        Displays the shift info card.
        FIX: Hides logistics (Entry/Exit/Pickup/Dropoff) for custom types marked as 'is_off',
        but keeps their specific Name and Color identity.
        """
        # 1. Validation and Setup
        if not (
            0 <= row < len(self._row_identities)
            and 0 <= col < len(self._date_col_dates)
        ):
            self._shift_info_card.hide()
            return

        item = self.schedule_table.item(row, col)
        if not item or not item.text().strip():
            self._shift_info_card.hide()
            return

        # Geometry for anchoring
        cell_rect_viewport = self.schedule_table.visualItemRect(item)
        global_top_left = self.schedule_table.viewport().mapToGlobal(
            cell_rect_viewport.topLeft()
        )
        global_cell_rect = QRect(global_top_left, cell_rect_viewport.size())

        # 2. Get Data Identity
        identity = self._row_identities[row]
        badge = identity.get("badge")
        hover_date = self._date_col_dates[col]

        if not badge or not hover_date:
            self._shift_info_card.hide()
            return

        # 3. Fetch SSoT Data
        schedule_data = db.get_schedule_map_for_range(
            badge, hover_date, hover_date, self.source
        )
        day_info = schedule_data.get(hover_date.isoformat())
        pickup, dropoff = db.get_user_location_for_date(badge, hover_date)

        operations = db.get_operations_filtered(
            text=badge, d_from=hover_date, d_to=hover_date
        )
        # Filter and sort operations (newest first)
        valid_ops = [op for op in operations if op.get("badge") == badge]
        valid_ops.sort(key=lambda x: x.get("id", 0), reverse=True)
        operation_info = valid_ops[0] if valid_ops else None

        if not day_info:
            final_html = "<p style='margin:0;'>Information not available</p>"
            self._shift_info_card.show_info(global_cell_rect, final_html)
            return

        # 4. Prepare Logic
        status_code = (day_info.get("status") or "N/A").upper()
        remark = day_info.get("remark")
        
        shift_title = status_code
        in_time = day_info.get("in_time")
        out_time = day_info.get("out_time")
        
        # --- DETECCIÓN DE "TREAT AS OFF" ---
        is_custom_off = False
        custom_color_hex = "#FFFFFF" # Default title color

        if status_code in self._custom_shift_map:
            info = self._custom_shift_map[status_code]
            is_custom_off = info.get("is_off", False)
            # Opcional: Si quisieras usar el color en el título del tooltip
            # custom_color_hex = info.get("color_hex", "#FFFFFF")

        # --- LOGICA DE TÍTULOS Y HORARIOS ---
        if status_code == "ON":
            shift_title = "ON (Day Shift)"
            # Default times logic...
            if not in_time: in_time = "06:00" if self.source == "Newmont" else "07:00"
            if not out_time: out_time = "12:00" if self.source == "Newmont" else "07:00"

        elif status_code == "ON NS":
            shift_title = "ON NS (Night Shift)"
            if not in_time: in_time = "12:00" if self.source == "Newmont" else "07:00"
            if not out_time: out_time = "06:00" if self.source == "Newmont" else "07:00"

        elif status_code in self._custom_shift_map:
            custom_info = self._custom_shift_map[status_code]
            shift_title = custom_info.get("name", status_code)
            
            # CRÍTICO: Si es 'Treat as OFF', matamos los horarios para que no se muestren
            if is_custom_off:
                in_time = None
                out_time = None
            else:
                # Si es un custom normal (Working), llenamos defaults si faltan
                if not in_time: in_time = custom_info.get("in_time")
                if not out_time: out_time = custom_info.get("out_time")

        elif status_code == "OFF":
            shift_title = "OFF"
            in_time = None
            out_time = None

        # 5. Build HTML Content
        # Usamos un color blanco fuerte para el título para asegurar contraste en el tooltip oscuro
        title_style = f"style='margin: 0 0 2px 0; font-size: 14px; color: #FFFFFF; font-weight: 600;'"
        schedule_style = "style='margin: 0; font-size: 13px; color: #FFFFFF;'"

        content_lines = [f"<p {title_style}>{shift_title}</p>"]

        # Mostrar reloj solo si hay horas Y NO es un día OFF
        if in_time and out_time and not is_custom_off:
            content_lines.append(f"<p {schedule_style}>⏰ {in_time} – {out_time}</p>")

        # --- LOGICA DE FILTRADO (ENTRY / EXIT) ---
        # Solo mostrar vuelos si NO es OFF y NO es Custom OFF
        if operation_info and status_code != "OFF" and not is_custom_off:
            entry_dt_str = operation_info.get("entry_date", "")
            exit_dt_str = operation_info.get("exit_date", "")

            if entry_dt_str and exit_dt_str and " " in entry_dt_str:
                try:
                    entry_dt = datetime.strptime(entry_dt_str, "%Y-%m-%d %H:%M")
                    exit_dt = datetime.strptime(exit_dt_str, "%Y-%m-%d %H:%M")
                    
                    travel_html = "<div style='border-top: 1px solid #4B5563; padding-top: 6px; margin-top: 8px;'>"
                    travel_html += f"<p {schedule_style}>✈️ <b>Entry:</b> {entry_dt.strftime('%Y-%m-%d %H:%M')}</p>"
                    travel_html += f"<p {schedule_style}>✈️ <b>Exit:</b> {exit_dt.strftime('%Y-%m-%d %H:%M')}</p>"
                    travel_html += "</div>"
                    content_lines.append(travel_html)
                except ValueError:
                    pass

        # --- LOGICA DE FILTRADO (PICKUP / DROP OFF) ---
        # Solo mostrar logística si NO es OFF y NO es Custom OFF
        pickup_clean = _clean(pickup)
        dropoff_clean = _clean(dropoff)

        if (pickup_clean or dropoff_clean) and status_code != "OFF" and not is_custom_off:
            logistics_html = "<div style='border-top: 1px solid #4B5563; padding-top: 6px; margin-top: 8px;'>"
            logistics_html += f"<p {schedule_style}>📍 <b>Pick Up:</b> {pickup_clean or 'Not assigned'}</p>"
            logistics_html += f"<p {schedule_style}>📍 <b>Drop Off:</b> {dropoff_clean or 'Not assigned'}</p>"
            logistics_html += "</div>"
            content_lines.append(logistics_html)

        # Remarks (Siempre mostrar si existen, incluso en OFF)
        if remark:
            remark_style = "style='margin: 0; font-size: 13px; color: #FFFFFF;'"
            remark_block_style = "style='border-top: 1px solid #1565C0; padding-top: 8px; margin-top: 8px;'"
            remark_html = f"<div {remark_block_style}>"
            remark_html += f"<p {remark_style}><b>Remark:</b> {remark}</p>"
            remark_html += "</div>"
            content_lines.append(remark_html)

        html_body = "".join(content_lines)
        final_html = f"<div style='line-height: 1.3;'>{html_body}</div>"

        self._shift_info_card.show_info(global_cell_rect, final_html)

    def load_users_to_selector(self):
        """
        Carga usuarios Y sus ubicaciones por defecto en el selector.
        """
        self.user_selector_combo.blockSignals(True)
        self.user_selector_combo.clear()
        
        # --- CAMBIO AQUÍ: Usamos la nueva función que trae los defaults ---
        # Antes era: db.get_all_users(self.source)
        self.users_for_selector = db.get_users_with_defaults(self.source)
        # ----------------------------------------------------------------
        
        self.user_selector_combo.addItem("-- Select a user --")
        for user in self.users_for_selector:
            self.user_selector_combo.addItem(user["name"])
            
        self.user_selector_combo.setCurrentIndex(0)
        self.user_selector_combo.blockSignals(False)
        
        # Limpiamos los campos dependientes
        self.role_display.clear()
        self.badge_display.clear()

    def refresh_users_only(self):
        """Repopulate ONLY the users combo (for instant sync)."""
        self.load_users_to_selector()

    def autofill_user_data(self, index):
        """
        Prellena rol/badge/defaults del usuario, pero respeta el Status actual.
        Si el Status es tipo OFF, NO llenamos la logística.
        """
        # 1. Obtener estado actual del selector de estatus
        current_sel = self.status_selector.currentData()
        is_current_status_off = _is_off_like_payload(current_sel) if isinstance(current_sel, dict) else False

        if index > 0:
            # La lista 'users_for_selector' coincide con el índice del combo (offset -1)
            user = self.users_for_selector[index - 1]
            
            # Llenar campos de solo lectura
            self.role_display.setText(user.get("role", ""))
            self.badge_display.setText(str(user.get("badge", "")) or "")

            # Obtener defaults del usuario
            def_pu = user.get("pickup_location")
            def_do = user.get("dropoff_location")

            # GUARDAR ESTADO: Actualizamos la memoria del "último working" 
            # para poder restaurar si el usuario cambia de OFF -> ON.
            self._last_working_pickup = def_pu or None
            self._last_working_dropoff = def_do or None

            # 2. Lógica: Llenar combos O Limpiar combos basado en el Status
            with QSignalBlocker(self.pickup_combo), QSignalBlocker(self.dropoff_combo):
                if is_current_status_off:
                    # El status dice NO -> Limpiar
                    self.pickup_combo.setCurrentIndex(0)
                    self.dropoff_combo.setCurrentIndex(0)
                else:
                    # El status dice SI -> Llenar con Defaults
                    self._set_combo_text(self.pickup_combo, def_pu)
                    self._set_combo_text(self.dropoff_combo, def_do)
        else:
            # Resetear UI si no hay usuario seleccionado
            self.role_display.clear()
            self.badge_display.clear()
            self._last_working_pickup = None
            self._last_working_dropoff = None
            
            with QSignalBlocker(self.pickup_combo), QSignalBlocker(self.dropoff_combo):
                self.pickup_combo.setCurrentIndex(0)
                self.dropoff_combo.setCurrentIndex(0)

    def _set_combo_text(self, combo, text):
        """Helper para setear el combo por texto de forma segura"""
        if text:
            idx = combo.findData(text)
            combo.setCurrentIndex(idx if idx >= 0 else 0)
        else:
            combo.setCurrentIndex(0)

    def _apply_schedule_period(
        self,
        username: str,
        badge: str,
        role: str,
        start_date,
        end_date,
        status: str | None,
        shift_type: str | None,
        pickup: str | None,
        dropoff: str | None,
        remark: str | None,
        in_time=None,
        out_time=None,
    ):
        """
        Aplica un período [start_date, end_date] a DB y Excel.
        Se basa en la misma lógica que save_plan_changes, pero
        sin mostrar diálogos.
        """
        # -------------------------------------------------------
        # 1) Decidir qué horas vamos a usar
        #    - Si vienen in_time/out_time como parámetros (copiados
        #      del día ancla), usamos esas tal cual.
        #    - Si no, usamos las horas definidas en el shift_type.
        # -------------------------------------------------------
        if in_time or out_time:
            # Caso 1: venimos de _apply_fill_from_anchor y ya tenemos horas
            final_in_time = in_time
            final_out_time = out_time
        else:
            # Caso 2: comportamiento normal → mirar el shift_type
            final_in_time = None
            final_out_time = None
            if shift_type and self._custom_shift_map:
                st = self._custom_shift_map.get(shift_type)
                if st:
                    final_in_time = st.get("in_time")
                    final_out_time = st.get("out_time")

        # -------------------------------------------------------
        # 2) Convertir esas horas HH:MM (o None) a datetime.time
        # -------------------------------------------------------
        entry_datetime = None
        exit_datetime = None

        if final_in_time or final_out_time:
            default_in = datetime.min.time()
            default_out = datetime.max.time()

            in_time_obj = _parse_hhmm_to_time(final_in_time, default_in)
            out_time_obj = _parse_hhmm_to_time(final_out_time, default_out)

           # 1. Calculamos si aplica "1+D" (Si es ON/ON NS -> end_date + 1)
            # Usamos el helper que creamos. Nota: 'shift_type' ya viene como argumento.
            real_exit_date = self._calculate_rgm_exit_date(shift_type, end_date)

            # 2. Usamos real_exit_date en lugar de end_date
            entry_datetime = datetime.combine(start_date, in_time_obj)
            exit_datetime = datetime.combine(real_exit_date, out_time_obj)

        # DB: operación
        db.add_operation(
            username=username,
            role=role,
            badge=badge,
            start_date=start_date,
            end_date=end_date,
            created_by=self.logged_username,
            entry_date=entry_datetime,
            exit_date=exit_datetime,
        )

        # DB: schedule
        if status is not None:
            db.upsert_schedule_range(
                badge,
                start_date,
                end_date,
                status,
                shift_type,
                self.source,
                in_time,
                out_time,
                remark,
            )
        else:
            db.clear_schedule_range(badge, start_date, end_date, self.source)

        # DB: locations
        if pickup or dropoff:
            db.assign_user_location_range(badge, start_date, end_date, pickup, dropoff)
            db.log_event(
                self.logged_username,
                self.source,
                "LOCATION_ASSIGN",
                f"{username} ({badge}) {start_date}..{end_date} PU={pickup} DO={dropoff}",
            )

        # Excel: actualizar PlanStaff
        success, message = excel.update_plan_staff_excel(
            self.excel_file,
            username,
            role,
            badge,
            status,
            shift_type,
            start_date,
            end_date,
            self.source,
            in_time,
            out_time,
        )
        if not success:
            # No muestro QMessageBox aquí para no interrumpir al usuario
            # pero podrías loguearlo si quieres.
            print("Excel update failed from cell edit:", message)
    # --------------------------------------------------------------------------
    # HELPER: Detección de Colisiones (Working -> Working)
    # --------------------------------------------------------------------------
    def _resolve_force_new_entry_start(self, badge, start_date, new_status):
        """
        Verifica si se está creando una continuidad Working->Working.
        Retorna:
          1 -> El usuario eligió SEPARAR (force_new_entry_start=1)
          0 -> El usuario eligió CONTINUAR o no hubo colisión (force_new_entry_start=0)
          None -> El usuario CANCELÓ la operación.
        """
        # 1. Si el nuevo estado es OFF o vacío, no hay colisión (no se parte viaje)
        if not new_status or not db.is_working_status(new_status, self.source):
            return 0 

        # 2. Consultar el día anterior en la BD
        prev_day = start_date - timedelta(days=1)
        # Usamos get_schedule_map_for_range para obtener info precisa del día previo
        prev_map = db.get_schedule_map_for_range(badge, prev_day, prev_day, self.source)
        
        # Extraer status del diccionario (key suele ser fecha string ISO)
        prev_day_str = prev_day.strftime("%Y-%m-%d")
        
        # Defensive coding: prev_map puede venir vacío o con otra key
        prev_info = prev_map.get(prev_day_str, {})
        prev_status = prev_info.get("status")

        # 3. Si el día anterior era OFF, no hay colisión (Off -> Working es normal)
        if not prev_status or not db.is_working_status(prev_status, self.source):
            return 0

        # 4. COLISIÓN DETECTADA (Working -> Working) -> Preguntar al usuario
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Icon.Warning)
        box.setWindowTitle("⚠️ Colisión de Turnos Detectada")
        box.setText(
            f"Se detectó continuidad de turnos para {badge}:\n\n"
            f"Día Previo: {prev_day_str} ({prev_status})\n"
            f"Nuevo Día : {start_date.strftime('%Y-%m-%d')} ({new_status})\n\n"
            "El sistema uniría esto en un solo viaje. ¿Qué desea hacer?"
        )

        cont_btn = box.addButton("Mantener Unido (1 Viaje)", QMessageBox.ButtonRole.AcceptRole)
        sep_btn  = box.addButton("Partir Turno (Nueva Salida)", QMessageBox.ButtonRole.DestructiveRole)
        cancel_btn = box.addButton("Cancelar", QMessageBox.ButtonRole.RejectRole)

        box.setDefaultButton(cont_btn)
        box.exec()
        
        clicked = box.clickedButton()

        if clicked == cancel_btn or clicked is None:
            return None # Señal de abortar

        if clicked == sep_btn:
            return 1 # Force split
            
        return 0 # Default: unir
    # ---------- actions ----------
    def save_plan_changes(self):
        # 1. Limpiar estados de error visuales
        mark_error(self.user_selector_combo, False)
        mark_error(self.role_display, False)
        mark_error(self.start_date_edit, False)
        mark_error(self.end_date_edit, False)
        mark_error(self.entry_date_edit, False)
        mark_error(self.exit_date_edit, False)

        # 2. Leer datos bÃ¡sicos del formulario
        username = self.user_selector_combo.currentText()
        badge = self.badge_display.text()
        role = self.role_display.text()
        start_date = self.start_date_edit.date().toPyDate()
        end_date = self.end_date_edit.date().toPyDate()

        pickup = self.pickup_combo.currentData() or None
        dropoff = self.dropoff_combo.currentData() or None
        remark = self.remarks_input.text().strip() or None

        # 3. Validaciones bÃ¡sicas
        if not username or username == "-- Select a user --":
            mark_error(self.user_selector_combo, True)
            QMessageBox.warning(self, "Incomplete Data", "Please select an employee.")
            return

        if not role:
            mark_error(self.role_display, True)
            QMessageBox.warning(
                self, "Incomplete Data", "Please select a role/department."
            )
            return

        if start_date > end_date:
            mark_error(self.start_date_edit, True)
            mark_error(self.end_date_edit, True)
            QMessageBox.warning(
                self, "Date Error", "Start date cannot be after end date."
            )
            return

        # 4. Interpretar selecciÃ³n del turno (Status/Shift)
        sel = self.status_selector.currentData()
        if not sel or sel.get("kind") in ("none", "separator"):
            schedule_status = None
            shift_type = None
            in_time_raw, out_time_raw = None, None
        elif sel.get("kind") == "base":
            schedule_status, shift_type = sel["status"], sel["shift_type"]
            in_time_raw, out_time_raw = sel.get("in_time"), sel.get("out_time")
        else:  # custom
            schedule_status, shift_type = sel["code"], sel["name"]
            in_time_raw, out_time_raw = sel.get("in_time"), sel.get("out_time")

        # ---------------------------------------------------------------------
        # 5. CÃLCULO DE FECHAS Y HORAS DE VIAJE (CORRECCIÃ“N PUNTUAL)
        # ---------------------------------------------------------------------
        entry_datetime = None
        exit_datetime = None

        if self.travel_dates_check.isChecked():
            # OPCIÃ“N A: El usuario define fechas especÃ­ficas manualmente
            entry_date = self.entry_date_edit.date().toPyDate()
            entry_time = self.entry_time_edit.time().toPyTime()
            entry_datetime = datetime.combine(entry_date, entry_time)

            exit_date = self.exit_date_edit.date().toPyDate()
            exit_time = self.exit_time_edit.time().toPyTime()
            exit_datetime = datetime.combine(exit_date, exit_time)

            # ValidaciÃ³n lÃ³gica de viaje
            if entry_datetime > exit_datetime:
                mark_error(self.entry_date_edit, True)
                mark_error(self.exit_date_edit, True)
                QMessageBox.warning(
                    self,
                    "Date Error",
                    "Entry date/time cannot be after Exit date/time.",
                )
                return

            if entry_datetime.date() > start_date or exit_datetime.date() < end_date:
                reply = QMessageBox.question(
                    self,
                    "Confirm Dates",
                    "The travel period does not fully encompass the work period. Continue?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No,
                )
                if reply == QMessageBox.StandardButton.No:
                    return

        elif schedule_status is not None:
            # OPCIÃ“N B: AutomÃ¡tico (Check desmarcado) -> Calcular Defaults
            # Se usa la fecha de inicio del periodo para Entry y fin para Exit.
            # Se inyectan las horas segÃºn el tipo de turno o reglas de negocio.

            # Hora por defecto base (07:00 / 07:00) Para RGM
            t_in = datetime.strptime("07:00", "%H:%M").time()
            t_out = datetime.strptime("07:00", "%H:%M").time()

            # LÃ³gica de Horas
            if in_time_raw and out_time_raw:
                # Si el turno (custom) ya trae horas definidas en DB
                try: 
                    t_in = datetime.strptime(str(in_time_raw)[:5], "%H:%M").time()
                    t_out = datetime.strptime(str(out_time_raw)[:5], "%H:%M").time()
                except: pass
            else:
                # Reglas Hardcoded (SSoT fallbacks)
                is_newmont = (self.source == "Newmont")
                
                if is_newmont:
                    if schedule_status == "ON": # DÃ­a Newmont
                        t_in = datetime.strptime("06:00", "%H:%M").time()
                        t_out = datetime.strptime("12:00", "%H:%M").time()
                    elif schedule_status == "ON NS": # Noche Newmont
                        t_in = datetime.strptime("12:00", "%H:%M").time()
                        t_out = datetime.strptime("06:00", "%H:%M").time()
                else:
                    # [MODIFICADO] Reglas RGM: Siempre 07:00 - 07:00
                    if schedule_status in ("ON", "ON NS"):
                        t_in = datetime.strptime("07:00", "%H:%M").time()
                        t_out = datetime.strptime("07:00", "%H:%M").time()

            # Crear los datetimes finales para guardar en BD
            # --- INICIO DEL CAMBIO 1+D ---
            
            # 1. Calculamos la fecha de salida real
            # Si es RGM y es ON/ON NS -> Salida es MaÃ±ana (end_date + 1)
            # Si es RGM y es Otro (Capacitacion) -> Salida es Hoy (end_date)
            # El mÃ©todo _calculate_rgm_exit_date encapsula esta lÃ³gica para no repetir if/else aquÃ­.
            real_exit_date = self._calculate_rgm_exit_date(schedule_status, end_date)

            # 2. Combinamos con las horas (t_in / t_out) que ya calculaste arriba
            entry_datetime = datetime.combine(start_date, t_in)
            exit_datetime = datetime.combine(real_exit_date, t_out) # Usamos real_exit_date
            
            # --- FIN DEL CAMBIO 1+D ---

        # ---------------------------------------------------------------------
        # FIN DE LA CORRECCIÃ“N
        # ---------------------------------------------------------------------

        # 6. DetecciÃ³n de Conflictos (Overwrite check)
        conflicts_excel = excel.find_conflicts(
            self.excel_file, username, badge, start_date, end_date
        )
        conflicts_db_map = db.get_schedule_map_for_range(
            badge, start_date, end_date, self.source
        )
        if conflicts_excel or conflicts_db_map:
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Warning)
            box.setWindowTitle("Overwrite Shift Confirmation")
            box.setText("Are you sure you want to modify the existing shift?")
            accept_btn = box.addButton("Accept", QMessageBox.ButtonRole.AcceptRole)
            box.addButton("Cancel", QMessageBox.ButtonRole.RejectRole)
            box.exec()
            if box.clickedButton() != accept_btn:
                return

        # Para auditorÃ­a
        prev_map = db.get_schedule_map_for_range(
            badge, start_date, end_date, self.source
        )

        # ----------------------------------------------------------
        # 6.5) Shift Collision Detector (Working-Working)
        # ----------------------------------------------------------
        # Regla:
        #   - Si Ayer fue OFF (o Non-Working) y Hoy es Working: guardar silencioso (force_new_entry=0).
        #   - Si Ayer fue Working y Hoy es Working: preguntar si se une al viaje actual o se separa.
        force_new_entry_flag = 0

        try:
            print(f"[SCD] schedule_status={schedule_status!r}, source={self.source!r}, badge={badge!r}, start_date={start_date}")
            if schedule_status and db.is_working_status(schedule_status, self.source):
                prev_day = start_date - timedelta(days=1)
                prev_day_map = db.get_schedule_map_for_range(badge, prev_day, prev_day, self.source)
                prev_status = prev_day_map.get(prev_day.isoformat(), {}).get("status")
                print(f"[SCD] prev_day={prev_day}, prev_day_map={prev_day_map}, prev_status={prev_status!r}")

                if prev_status and db.is_working_status(prev_status, self.source):
                    print("[SCD] >>> WORKING->WORKING detected! Showing popup...")
                    box = QMessageBox(self)
                    box.setIcon(QMessageBox.Icon.Question)
                    box.setWindowTitle("Shift Collision Detector")
                    box.setText(
                        f"El día anterior ({prev_day.isoformat()}) tiene turno activo: {prev_status}.\n\n"
                        "¿Desea UNIR esto al viaje actual o registrar una NUEVA entrada?"
                    )

                    join_btn = box.addButton("Unir (Continuar)", QMessageBox.ButtonRole.AcceptRole)
                    sep_btn  = box.addButton("Separar (Nueva Entrada)", QMessageBox.ButtonRole.DestructiveRole)
                    cancel_btn = box.addButton("Cancelar", QMessageBox.ButtonRole.RejectRole)

                    box.setDefaultButton(join_btn)
                    box.exec()

                    clicked = box.clickedButton()
                    if clicked == sep_btn:
                        force_new_entry_flag = 1
                        print("[SCD] User chose: SEPARAR (force_new_entry=1)")
                    elif clicked == join_btn:
                        force_new_entry_flag = 0
                        print("[SCD] User chose: UNIR (force_new_entry=0)")
                    else:
                        print("[SCD] User chose: CANCELAR — aborting save")
                        return  # Cancelar — aborta el guardado
                else:
                    print(f"[SCD] No collision: prev_day was OFF/empty (force_new_entry=0)")
            else:
                print(f"[SCD] Skipped: schedule_status={schedule_status!r} is not working or is None")
        except Exception as e:
            import traceback as _tb
            print(f"[SCD] *** EXCEPTION in Shift Collision Detector: {e}")
            _tb.print_exc()
            force_new_entry_flag = 0

        # 7. Guardar en BD (SSoT)
        if schedule_status is not None:
            # AquÃ­ es donde se guardan los datetimes calculados (entry_datetime/exit_datetime)
            db.add_operation(
                username=username,
                role=role,
                badge=badge,
                start_date=start_date,
                end_date=end_date,
                created_by=self.logged_username,
                entry_date=entry_datetime,  # Ahora siempre tendrÃ¡ valor si hay turno
                exit_date=exit_datetime,  # Ahora siempre tendrÃ¡ valor si hay turno
            )
            db.upsert_schedule_range(
                badge,
                start_date,
                end_date,
                schedule_status,
                shift_type,
                self.source,
                in_time_raw,
                out_time_raw,
                remark,
                force_new_entry_start=force_new_entry_flag,
            )
        else:  # Limpiar rango ("Do Not Mark Days")
            db.clear_schedule_range(badge, start_date, end_date, self.source)

        is_off_day_guard = False
        if sel:
            # 1. Es "Do Not Mark Days" (kind='none')
            if sel.get("kind") == "none":
                is_off_day_guard = True
            # 2. Es Custom con flag de OFF (ej. SICK, VACATION) -> sel.get("is_off")
            elif sel.get("is_off"):
                is_off_day_guard = True
            # 3. Es Base 'OFF'
            elif sel.get("status") == "OFF":
                is_off_day_guard = True
        else:
            # Si no hay selecciÃ³n (sel is None), asumimos que no se marca (OFF-like)
            is_off_day_guard = True

        if is_off_day_guard:
            pickup = None
            dropoff = None


        # Guardar ubicaciÃ³n si aplica
        if pickup or dropoff:
            db.assign_user_location_range(badge, start_date, end_date, pickup, dropoff)
            db.log_event(
                self.logged_username,
                self.source,
                "LOCATION_ASSIGN",
                f"{username} ({badge}) {start_date}..{end_date} PU={pickup} DO={dropoff}",
            )

        # 8. Actualizar Excel
        self._is_internal_update = True  # Flag para evitar recarga innecesaria
        success, message = excel.update_plan_staff_excel(
            self.excel_file,
            username,
            role,
            badge,
            schedule_status,
            shift_type,
            start_date,
            end_date,
            self.source,
            in_time_raw,
            out_time_raw,
        )
        QTimer.singleShot(2000, lambda: setattr(self, "_is_internal_update", False))

        # 9. AuditorÃ­a y FinalizaciÃ³n
        new_map = db.get_schedule_map_for_range(
            badge, start_date, end_date, self.source
        )
        db.log_event(
            self.logged_username,
            self.source,
            "SHIFT_MODIFICATION",
            f"{username} ({badge}) {start_date}..{end_date} prev={prev_map} new={new_map} remark={remark}; Excel={'OK' if success else 'ERR'}",
        )

        box = QMessageBox(self)
        box.setIcon(
            QMessageBox.Icon.Information if success else QMessageBox.Icon.Warning
        )
        box.setWindowTitle("Success" if success else "Warning")
        box.setText(message if success else ("Saved to DB. " + message))
        box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
        box.exec()

        self.refresh_ui_data()
        self.check_excel_health()
        self.rotation_changed.emit()

    def generate_report(self):
        settings = db.get_report_settings(self.logged_username, self.source)
        s = self.report_start_date.date().toPyDate()
        e = self.report_end_date.date().toPyDate()

        excel_data, message = excel.generate_transport_report(
            self.excel_file, s, e, settings
        )

        if not excel_data:
            QMessageBox.critical(self, "Report Error", message)
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Transport Report",
            f"Transport_Report_{self.source}_{datetime.now().strftime('%Y%m%d')}.xlsx",
            "Excel Files (*.xlsx)",
        )

        if not file_path:
            return
        try:
            with open(file_path, "wb") as f:
                f.write(excel_data)
            QMessageBox.information(
                self, "Success", f"{message}\n\nReport saved to:\n{file_path}"
            )
            db.log_event(
                self.logged_username,
                self.source,
                "DATA_EXPORT",
                f"TRANSPORT -> {file_path}",
            )
        except Exception as e:
            QMessageBox.critical(
                self, "Save Error", f"Could not save the file.\nError: {e}"
            )

    # =======================================================
    # NUEVA FUNCIONALIDAD: Lógica para el reporte de estadías
    # =======================================================
    def _generate_stay_report(self):
        """Generates and saves the Onsite Stay Period report."""
        start_date = self.report_start_date.date().toPyDate()
        end_date = self.report_end_date.date().toPyDate()

        if start_date > end_date:
            QMessageBox.warning(
                self, "Date Range Error", "Start date cannot be after end date."
            )
            return

        try:
            # Default filename suggestion
            default_filename = f"Onsite_Stay_Report_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx"

            # Open file dialog to choose where to save
            filePath, _ = QFileDialog.getSaveFileName(
                self,
                "Save Onsite Stay Report",
                default_filename,
                "Excel Files (*.xlsx);;All Files (*)",
            )

            if not filePath:
                return  # User cancelled

            # Generate the report bytes
            report_bytes, msg = excel.generate_stay_period_report(
                self.excel_file, start_date, end_date
            )

            if report_bytes:
                with open(filePath, "wb") as f:
                    f.write(report_bytes)

                QMessageBox.information(
                    self,
                    "Report Generated",
                    f"Successfully generated and saved report:\n{os.path.basename(filePath)}",
                )
                db.log_event(
                    self.logged_username,
                    self.source,
                    "REPORT_STAY_GENERATED",
                    f"Generated for range {start_date} to {end_date}.",
                )
            else:
                QMessageBox.critical(self, "Report Generation Failed", msg)

        except Exception as e:
            QMessageBox.critical(
                self,
                "Error",
                f"An unexpected error occurred while generating the report: {e}",
            )
            db.log_event(
                self.logged_username,
                self.source,
                "ERROR_REPORT_STAY",
                f"Failed for range {start_date} to {end_date}: {e}",
            )

    def export_plan_from_db(self):
        """FR-03: Export plan (from DB state; includes custom shift types)."""
        users = db.get_all_users(self.source)
        schedules = db.get_schedules_for_source(self.source)

        if not users:
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Warning)
            box.setWindowTitle("No Data")
            box.setText("There are no users to export.")
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
            box.exec()
            return

        default_name = (
            f"PlanStaff_{self.source}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )
        dest_path, _ = QFileDialog.getSaveFileName(
            self, "Save Plan Staff", default_name, "Excel Files (*.xlsx)"
        )
        if not dest_path:
            return

        ok, msg = excel.export_plan_from_db(
            self.excel_file, users, schedules, dest_path, self.source
        )
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Icon.Information if ok else QMessageBox.Icon.Critical)
        box.setWindowTitle("Export" if ok else "Error")
        box.setText(msg)
        box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
        box.exec()

        if ok:
            db.log_event(
                self.logged_username,
                self.source,
                "DATA_EXPORT",
                f"PLAN_EXPORT -> {dest_path}",
            )

    def refresh_ui_data(self, use_preloaded=False):
        self.load_shift_type_options()
        self.load_schedule_data(use_preloaded=use_preloaded) # Pasar la bandera
        self.load_users_to_selector()
        self.load_location_options()  # keep combos in sync with Location admin
        self.remarks_input.clear()  # Clear remarks on refresh

    # ---------- Excel Health / Monitoring ----------
    def check_excel_health(self):
        # This entire block is wrapped in a try/except to prevent a crash
        # if the timer fires after the widget has been destroyed (e.g., on close).
        try:
            exists = os.path.exists(self.excel_file)
            if not exists:
                self.excel_health_label.setText(
                    "Excel status: ❌ Not found (it may have been moved, deleted, or renamed)."
                )
                self.excel_health_label.setStyleSheet(
                    "color: #B00020; font-weight: bold;"
                )
                if not self._missing_prompt_shown:
                    self._missing_prompt_shown = True
                    # Use a single shot timer to call the prompt after the current event loop finishes,
                    # which is safer than opening a dialog directly from this handler.
                    QTimer.singleShot(0, self.prompt_regenerate_or_locate)
                return

            # Exists -> validate structure and detect changes
            mtime = os.path.getmtime(self.excel_file)
            structure_ok, errors, meta = excel.validate_excel_structure(self.excel_file)
            if structure_ok:
                # ✅ Show the signed-in site (RGM/Newmont), not the structural variant
                self.excel_health_label.setText(
                    f"Excel status: ✅ OK ({self.source}) — {os.path.basename(self.excel_file)}"
                )
                self.excel_health_label.setStyleSheet(
                    "color: #1B5E20; font-weight: bold;"
                )
            else:
                self.excel_health_label.setText(
                    "Excel status: ⚠️ Invalid structure. Use 'Regenerate' or fix the file."
                )
                self.excel_health_label.setStyleSheet(
                    "color: #E65100; font-weight: bold;"
                )

            # If file changed (mtime) -> refresh preview
            if self._last_excel_mtime is None or mtime != self._last_excel_mtime:
                # --- AGREGAR ESTE BLOQUE DE SEGURIDAD ---
                if self._is_internal_update:
                    # Actualizamos el mtime para que la próxima vez no crea que es nuevo
                    self._last_excel_mtime = mtime
                    print("DEBUG: Ignorando recarga por guardado interno.")
                    return
                # ----------------------------------------
                self._last_excel_mtime = mtime
                self.load_schedule_data()
        except RuntimeError:
            # This error occurs if the QLabel widget has been deleted by the time
            # this timer callback runs. We can safely ignore it.
            pass
        except Exception as e:
            # For any other unexpected error, we can stop the timer and log it.
            # It's better to check if the label still exists before trying to set its text.
            if self.excel_health_label:
                self.excel_health_label.setText(
                    f"Excel status: ⚠️ Error validating file: {e}"
                )
                self.excel_health_label.setStyleSheet(
                    "color: #E65100; font-weight: bold;"
                )
            self.file_watch_timer.stop()

    def prompt_regenerate_or_locate(self):
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Icon.Warning)
        msg.setWindowTitle("Plan Staff file not available")
        msg.setText(
            f"The file cannot be found:\n{self.excel_file}\n\n"
            f"The system can regenerate it from the DB (SSoT) or you can locate it manually."
        )
        regen_btn = msg.addButton("🛠️ Regenerate now", QMessageBox.ButtonRole.AcceptRole)
        locate_btn = msg.addButton("📂 Locate file…", QMessageBox.ButtonRole.ActionRole)
        msg.addButton("Cancel", QMessageBox.ButtonRole.RejectRole)
        msg.exec()

        if msg.clickedButton() == regen_btn:
            self.regenerate_excel_from_db()
        elif msg.clickedButton() == locate_btn:
            new_path, _ = QFileDialog.getOpenFileName(
                self, "Select PlanStaff", "", "Excel Files (*.xlsx)"
            )
            if new_path:
                # Normalizar ruta
                new_path = os.path.abspath(new_path)
                old_path = self.excel_file
                self.excel_file = new_path

                # --- NUEVO: Persistir y notificar ---
                # 1. Guardar en BD
                db.set_file_path(self.source, "plan_staff", new_path, updated_by=self.logged_username)
                db.log_event(self.logged_username, self.source, "EXCEL_PATH_UPDATE",
                             f"plan_staff: {old_path} -> {new_path}")

                # 2. Emitir señal para avisar a otros widgets
                self.excel_path_changed.emit(self.source, new_path)
                # ------------------------------------

                self._missing_prompt_shown = False
                self.check_excel_health()
                self.refresh_ui_data()

    def regenerate_excel_from_db(self):
        ok, msg = excel.regenerate_plan_from_db(self.excel_file, self.source)
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Icon.Information if ok else QMessageBox.Icon.Critical)
        box.setWindowTitle("Regenerate" if ok else "Error")
        box.setText(msg)
        box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
        box.exec()
        if ok:
            db.log_event(
                self.logged_username,
                self.source,
                "DATA_EXPORT",
                f"PLAN_REGENERATE -> {self.excel_file}",
            )
            self._missing_prompt_shown = False
            self.check_excel_health()
            self.refresh_ui_data()

    def refresh_excel_from_db_ui(self):
        ok, msg = excel.refresh_excel_from_db(self.excel_file, self.source)
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Icon.Information if ok else QMessageBox.Icon.Critical)
        box.setWindowTitle("Refresh" if ok else "Error")
        box.setText(msg)
        box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
        box.exec()
        if ok:
            db.log_event(
                self.logged_username,
                self.source,
                "DATA_SYNC",
                f"PLAN_REFRESH -> {self.excel_file}",
            )
            self.check_excel_health()
            self.refresh_ui_data()

    def validate_excel_structure_ui(self):
        ok, errors, meta = excel.validate_excel_structure(self.excel_file)
        box = QMessageBox(self)
        if ok:
            box.setIcon(QMessageBox.Icon.Information)
            box.setWindowTitle("Valid Structure")
            box.setText(
                f"Template: {meta.get('variant','?')} | Date columns: {meta.get('date_columns',0)}"
            )
        else:
            box.setIcon(QMessageBox.Icon.Warning)
            box.setWindowTitle("Invalid Structure")
            box.setText("Issues detected:\n\n" + "\n".join(f"- {e}" for e in errors))
        box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
        box.exec()

    def compare_excel_db_ui(self):
        report = excel.check_db_sync_with_excel(self.excel_file, self.source)
        mismatches = report.get("schedule_mismatches", [])
        text = []
        text.append(f"Users in Excel: {report.get('users_in_excel',0)}")
        text.append(f"Users in DB:    {report.get('users_in_db',0)}")
        if report.get("missing_badges_in_db"):
            text.append(
                f"\nMissing in DB (badges): {', '.join(report['missing_badges_in_db'])}"
            )
        if report.get("extra_badges_in_db"):
            text.append(
                f"Extra in DB (badges not in Excel): {', '.join(report['extra_badges_in_db'])}"
            )
        text.append(f"\nSchedule mismatches: {len(mismatches)}")
        preview = "\n".join(text[:1000])

        box = QMessageBox(self)
        box.setIcon(QMessageBox.Icon.Information)
        box.setWindowTitle("Excel vs DB Comparison")
        box.setText(preview if len(preview) < 1500 else (preview[:1500] + "\n..."))
        box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
        box.exec()
        
    def keyPressEvent(self, event):
        # COPY (Use QKeySequence, not Qt.KeySequence)
        if event.matches(QKeySequence.StandardKey.Copy):
            ScheduleClipboardService.copy_to_clipboard(
                self.schedule_table, 
                self._row_identities, 
                self._date_col_dates, 
                self._custom_shift_map
            )
            return

        # PASTE (Use QKeySequence, not Qt.KeySequence)
        if event.matches(QKeySequence.StandardKey.Paste):
            self._handle_paste_operation()
            return

        super().keyPressEvent(event)

    def _handle_paste_operation(self):
        """
        Maneja la operación de PEGAR (Paste) replicando la lógica "Smart Drag".
        
        Flujo por celda:
        1. Parsing del Clipboard.
        2. Cálculo de horas (Calculate Time Logic).
        3. Upsert en Schedule (SSoT).
        4. Consolidación Logística (Operations) inmediata para reparar vecinos.
        5. Actualización Excel.
        """
        # 1. Parsing del Clipboard
        payload, fmt = ScheduleClipboardService.parse_clipboard()
        if not payload or not payload.get("grid"):
            return

        selected_ranges = self.schedule_table.selectedRanges()
        if not selected_ranges:
            return

        # 2. Geometría
        anchor_row = selected_ranges[0].topRow()
        anchor_col = selected_ranges[0].leftColumn()
        
        sel_rows = selected_ranges[0].rowCount()
        sel_cols = selected_ranges[0].columnCount()
        
        src_rows = payload['rows']
        src_cols = payload['cols']
        src_grid = payload['grid']

        # Ajuste de tamaño (Repetir patrón o 1:1)
        if sel_rows == 1 and sel_cols == 1:
            target_rows = src_rows
            target_cols = src_cols
        else:
            target_rows = sel_rows
            target_cols = sel_cols

        # UI State
        self._is_internal_update = True 
        self._bulk_editing = True
        
        # Para optimización de Excel (escritura masiva al final)
        excel_updates_queue = [] 
        
        try:
            conn = db.sqlite3.connect(db.DB_FILE)
            conn.row_factory = db.sqlite3.Row
            cursor = conn.cursor()
            cursor.execute("BEGIN TRANSACTION")

            try:
                for r in range(target_rows):
                    for c in range(target_cols):
                        # --- Coordenadas ---
                        abs_row = anchor_row + r
                        abs_col = anchor_col + c

                        # Validaciones de límites
                        if abs_row >= self.schedule_table.rowCount(): continue
                        if abs_col >= self.schedule_table.columnCount(): continue
                        
                        # Obtener Identidad
                        identity = self._row_identities[abs_row] if abs_row < len(self._row_identities) else None
                        if not identity: continue
                        
                        badge = identity.get("badge")
                        username = identity.get("name")
                        role = identity.get("role")
                        if not badge: continue
                        
                        # Fecha Objetivo
                        if abs_col >= len(self._date_col_dates): continue
                        target_date = self._date_col_dates[abs_col]

                        # --- Extracción de datos del portapapeles ---
                        src_r_idx = r % src_rows
                        src_c_idx = c % src_cols
                        cell_data = src_grid[src_r_idx][src_c_idx]
                        raw_text = cell_data.get("clean_code", "")
                        
                        # --- [LOGICA SMART DRAG] 1. Calculo de Horas ---
                        # Usamos raw_text como status base. 
                        # NOTA: Al pegar texto plano, no tenemos horas "raw" previas, 
                        # así que confiamos en la lógica de negocio para rellenarlas.
                        final_in = self._calculate_time_logic(raw_text, "IN", None)
                        final_out = self._calculate_time_logic(raw_text, "OUT", None)
                        
                        in_str = final_in.strftime("%H:%M") if final_in else None
                        out_str = final_out.strftime("%H:%M") if final_out else None
                        # [PATCH COLISIÓN] --------------------------------
                        # Validamos colisión celda por celda al pegar
                        force_flag = self._resolve_force_new_entry_start(badge, target_date, raw_text)
                        
                        if force_flag is None:
                            # Si cancela, saltamos esta celda (o podrías hacer 'break' para cancelar todo)
                            continue
                        # -------------------------------------------------
                        
                        # --- [LOGICA SMART DRAG] 2. DB Schedule Upsert ---
                        db.upsert_schedule_range(
                            badge, target_date, target_date, 
                            raw_text, None, self.source, 
                            in_str, out_str, "",
                            force_new_entry_start=force_flag,
                            cursor=cursor
                        )
                        
                        # --- [LOGICA SMART DRAG] 3. Consolidación Logística ---
                        # Esto es lo que repara los "vecinos". Al consolidar el día X, 
                        # el sistema revisa X-1 y X+1. Si rompe un bloque antiguo, 
                        # la función _consolidate_and_record_logistics se encarga de 
                        # re-crear la operación correcta para este día.
                        self._consolidate_and_record_logistics(
                            badge, role, username, 
                            target_date, target_date, raw_text,
                            cursor=cursor
                        )

                        # --- Actualización Visual ---
                        item = self.schedule_table.item(abs_row, abs_col)
                        if not item:
                            item = QTableWidgetItem()
                            self.schedule_table.setItem(abs_row, abs_col, item)
                        item.setText(raw_text)
                        self._apply_base_background(item, raw_text)
                        self._cell_original_values[(abs_row, abs_col)] = raw_text

                        # --- Cola para Excel ---
                        bg_brush = item.background()
                        color_hex = bg_brush.color().name() if bg_brush.style() != Qt.BrushStyle.NoBrush else None
                        
                        excel_updates_queue.append({
                            "badge": badge,
                            "date": target_date,
                            "val": raw_text,
                            "in": in_str,
                            "out": out_str,
                            "color": color_hex
                        })

                cursor.execute("COMMIT")

            except Exception as e:
                cursor.execute("ROLLBACK")
                raise e
            finally:
                conn.close()

            # =================================================================
            # FASE EXCEL: Escritura Masiva (Igual que antes, optimizado)
            # =================================================================
            if excel_updates_queue and os.path.exists(self.excel_file):
                try:
                    wb = openpyxl.load_workbook(self.excel_file)
                    ws = wb.active
                    
                    header_map = {cell.value: cell.column for cell in ws[1] if isinstance(cell.value, str)}
                    date_map = {cell.value.date(): cell.column for cell in ws[1] if isinstance(cell.value, datetime)}
                    badge_col_idx = header_map.get("BADGE") or header_map.get("Company ID")
                    
                    if badge_col_idx:
                        row_map = {}
                        for r_idx in range(2, ws.max_row + 1):
                            cell_val = ws.cell(row=r_idx, column=badge_col_idx).value
                            if cell_val: row_map[str(cell_val).strip()] = r_idx
                        
                        for up in excel_updates_queue:
                            r_idx = row_map.get(str(up["badge"]).strip())
                            c_idx = date_map.get(up["date"])
                            if r_idx and c_idx:
                                cell = ws.cell(row=r_idx, column=c_idx)
                                val_str = up["val"]
                                cell.value = val_str
                                
                                if up["color"] and val_str:
                                    clean_hex = up["color"].replace("#", "").upper()
                                    if len(clean_hex) == 6:
                                        cell.fill = PatternFill(start_color=clean_hex, end_color=clean_hex, fill_type="solid")
                                else:
                                    cell.fill = PatternFill(fill_type=None)
                                    
                                if val_str not in ("ON", "ON NS", "OFF", "") and up["in"]:
                                    cell.comment = Comment(f"{up['in']}-{up['out']}", "ShiftType")
                                else:
                                    cell.comment = None

                        wb.save(self.excel_file)
                except Exception as ex_excel:
                    print(f"Error writing to Excel during paste: {ex_excel}")

            # Sincronización final
            self.check_excel_health()

        except Exception as e:
            QMessageBox.critical(self, "Paste Error", f"Failed to paste data: {e}")
            import traceback
            traceback.print_exc()
        
        finally:
            self._is_internal_update = False
            self._bulk_editing = False
            self.rotation_changed.emit()
            self.schedule_table.viewport().update()
            
    def _on_register_status_changed(self, index):
        data = self.status_selector.itemData(index)
        if not isinstance(data, dict):
            return

        is_off = _is_off_like_payload(data)

        with QSignalBlocker(self.pickup_combo), QSignalBlocker(self.dropoff_combo):
            if is_off:
                # 1. Lógica OFF: 
                # Guardamos la selección actual antes de borrar (por si fue manual)
                cur_pu = self.pickup_combo.currentData()
                cur_do = self.dropoff_combo.currentData()
                if cur_pu: self._last_working_pickup = cur_pu
                if cur_do: self._last_working_dropoff = cur_do

                # Limpiar
                self.pickup_combo.setCurrentIndex(0)
                self.dropoff_combo.setCurrentIndex(0)
            
            else:
                # 2. Lógica ON: Restaurar estado previo válido
                # Chequear si está vacío actualmente, si es así, restaurar memoria
                if self.pickup_combo.currentIndex() <= 0:
                    self._set_combo_text(self.pickup_combo, getattr(self, '_last_working_pickup', None))
                
                if self.dropoff_combo.currentIndex() <= 0:
                    self._set_combo_text(self.dropoff_combo, getattr(self, '_last_working_dropoff', None))
                
                # Fallback: Si la memoria está vacía, intentar recargar del Objeto Usuario actual
                if self.pickup_combo.currentIndex() <= 0 or self.dropoff_combo.currentIndex() <= 0:
                     uidx = self.user_selector_combo.currentIndex()
                     if uidx > 0:
                         user = self.users_for_selector[uidx - 1]
                         if self.pickup_combo.currentIndex() <= 0:
                             self._set_combo_text(self.pickup_combo, user.get('pickup_location'))
                         if self.dropoff_combo.currentIndex() <= 0:
                             self._set_combo_text(self.dropoff_combo, user.get('dropoff_location'))

    def _calculate_rgm_exit_date(self, shift_code, end_date):
        """
        Si es RGM y el turno es ON u ON NS, la salida es al día siguiente.
        Para cualquier otro turno (manual), la salida es el mismo día.
        """
        if self.source == "RGM" and shift_code in ["ON", "ON NS"]:
            from datetime import timedelta
            return end_date + timedelta(days=1)
        return end_date

# -------------------------------------------------------------
# Widget: Rotation History (own tab, without ID column)
# -------------------------------------------------------------
class RotationHistoryWidget(QWidget):
    def __init__(self, created_by: str | None = None):
        """
        Si created_by es None => modo 'admin' (ve todo).
        En caso contrario, solo muestra operaciones creadas por ese usuario.
        """
        super().__init__()
        self._created_by = (created_by or "").strip() or None
        self._filter_state = {}
        self._debounce_timer = QTimer(self)
        self._debounce_timer.setSingleShot(True)
        self._debounce_timer.timeout.connect(self.refresh_data)

        layout = QVBoxLayout(self)

        # -- Filter Panel --
        filter_layout = QHBoxLayout()
        filter_layout.setContentsMargins(0, 0, 0, 8)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search Name/Badge/Role...")
        self.search_input.textChanged.connect(self._request_refresh)

        self.role_combo = QComboBox()
        self.role_combo.currentIndexChanged.connect(self._request_refresh)

        self.date_from = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_from.setDisplayFormat("yyyy-MM-dd")
        self.date_from.dateChanged.connect(self._request_refresh)

        self.date_to = QDateEdit()
        self.date_to.setCalendarPopup(True)
        self.date_to.setDisplayFormat("yyyy-MM-dd")
        self.date_to.dateChanged.connect(self._request_refresh)

        self.active_today_check = QCheckBox("Active today")
        self.active_today_check.toggled.connect(self._toggle_active_today)

        self.sort_combo = QComboBox()
        self.sort_combo.addItems(["Start Date (desc)", "Name (asc)"])
        self.sort_combo.currentIndexChanged.connect(self._request_refresh)

        reset_button = QPushButton("Reset")
        reset_button.clicked.connect(self.reset_filters)

        filter_layout.addWidget(self.search_input, 2)
        filter_layout.addWidget(QLabel("Role:"))
        filter_layout.addWidget(self.role_combo, 1)
        filter_layout.addWidget(QLabel("Date Range:"))
        filter_layout.addWidget(self.date_from)
        filter_layout.addWidget(QLabel("–"))
        filter_layout.addWidget(self.date_to)
        filter_layout.addWidget(self.active_today_check)
        filter_layout.addStretch()
        filter_layout.addWidget(QLabel("Sort by:"))
        filter_layout.addWidget(self.sort_combo)
        filter_layout.addWidget(reset_button)

        layout.addLayout(filter_layout)

        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        layout.addWidget(self.table)

        self.reset_filters()  # To initialize and load data

    def _request_refresh(self):
        self._debounce_timer.start(DEBOUNCE_MS)

    def _toggle_active_today(self, checked):
        self.date_from.setEnabled(not checked)
        self.date_to.setEnabled(not checked)
        if checked:
            today = QDate.currentDate()
            self.date_from.setDate(today)
            self.date_to.setDate(today)
        self._request_refresh()

    def _populate_role_filter(self):
        self.role_combo.blockSignals(True)
        current_role = self.role_combo.currentText()
        self.role_combo.clear()
        self.role_combo.addItem("All Roles", None)
        # roles solo de operaciones creadas por el usuario (si aplica)
        all_records = db.get_operations_filtered(created_by=self._created_by)
        roles = sorted(list(set(r["role"] for r in all_records if r.get("role"))))
        self.role_combo.addItems(roles)

        idx = self.role_combo.findText(current_role)
        if idx != -1:
            self.role_combo.setCurrentIndex(idx)
        self.role_combo.blockSignals(False)

    def reset_filters(self):
        with QSignalBlocker(self.search_input), QSignalBlocker(
            self.role_combo
        ), QSignalBlocker(self.date_from), QSignalBlocker(self.date_to), QSignalBlocker(
            self.active_today_check
        ), QSignalBlocker(
            self.sort_combo
        ):
            self.search_input.clear()
            self._populate_role_filter()
            self.role_combo.setCurrentIndex(0)
            self.date_from.setDate(QDate(2000, 1, 1))
            self.date_to.setDate(QDate.currentDate().addYears(5))
            self.active_today_check.setChecked(False)
            self.sort_combo.setCurrentIndex(0)
        self.refresh_data()

    def refresh_data(self):
        # Persist filter state
        self._filter_state["text"] = self.search_input.text()
        self._filter_state["role"] = (
            self.role_combo.currentText()
            if self.role_combo.currentIndex() > 0
            else None
        )
        self._filter_state["sort"] = (
            "name_asc" if self.sort_combo.currentIndex() == 1 else "start_date_desc"
        )

        d_from = self.date_from.date().toPyDate()
        d_to = self.date_to.date().toPyDate()

        records = db.get_operations_filtered(
            text=self._filter_state["text"],
            role=self._filter_state["role"],
            d_from=d_from,
            d_to=d_to,
            sort_by=self._filter_state["sort"],
            created_by=self._created_by,  # NUEVO: solo mis rotaciones
        )

        headers = ["Name", "Role", "Badge", "Start Date", "End Date", "Created By"]
        self.table.setRowCount(len(records))
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)

        for row_idx, record in enumerate(records):
            self.table.setItem(row_idx, 0, QTableWidgetItem(record["username"]))
            self.table.setItem(row_idx, 1, QTableWidgetItem(record["role"]))
            self.table.setItem(row_idx, 2, QTableWidgetItem(record["badge"]))
            self.table.setItem(row_idx, 3, QTableWidgetItem(record["start_date"]))
            self.table.setItem(row_idx, 4, QTableWidgetItem(record["end_date"]))
            self.table.setItem(
                row_idx, 5, QTableWidgetItem(record.get("created_by", ""))
            )

        self.table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch
        )


# -------------------------------------------------------------
# Widget: Users CRUD (with Import from Excel)
# -------------------------------------------------------------
# Dentro de main_window.py

class CrudWidget(QWidget):
    # Signals for immediate UI sync
    import_done = pyqtSignal(str)  # emits 'source' when import finishes
    users_changed = pyqtSignal(str)  # emits 'source' when user list changes

    def __init__(self, source: str, excel_file: str, logged_username: str):
        super().__init__()
        self.source = source
        self.excel_file = excel_file
        self.logged_username = logged_username or "Unknown"
        self.current_user_id = None

        self._filter_state = {}
        self._debounce_timer = QTimer(self)
        self._debounce_timer.setSingleShot(True)
        self._debounce_timer.timeout.connect(self.load_users_table)

        layout = QHBoxLayout(self)

        # ---------------------------------------------------------
        # Left panel: User Form
        # ---------------------------------------------------------
        form_layout = QGridLayout()
        form_layout.setContentsMargins(8, 8, 8, 8)

        self.crud_name_input = QLineEdit()
        self.crud_role_input = QComboBox()
        self.crud_role_input.setEditable(False)
        self.crud_badge_input = QLineEdit()

        # --- NEW: Default Location Combos ---
        self.def_pickup_combo = QComboBox()
        self.def_dropoff_combo = QComboBox()
        
        # Buttons
        self.crud_save_button = QPushButton("💾 Save User")
        self.crud_save_button.clicked.connect(self.save_crud_user)
        self.crud_new_button = QPushButton("✨ New User")
        self.crud_new_button.clicked.connect(self.clear_crud_form)
        self.crud_delete_button = QPushButton("❌ Delete User")
        self.crud_delete_button.clicked.connect(self.delete_crud_user)

        self.import_button = QPushButton("📥 Import from Excel → DB (validated)")
        self.import_button.clicked.connect(self.import_users_from_excel)

        # Styling
        self.crud_save_button.setProperty("variant", "primary")
        self.crud_new_button.setProperty("variant", "secondary")
        self.crud_delete_button.setProperty("danger", True)
        self.import_button.setProperty("variant", "secondary")

        # Layout Setup
        row = 0
        form_layout.addWidget(QLabel("Full Name:"), row, 0)
        form_layout.addWidget(self.crud_name_input, row, 1)
        row += 1
        form_layout.addWidget(QLabel("Role/Department:"), row, 0)
        form_layout.addWidget(self.crud_role_input, row, 1) # Adding the combo
        row += 1
        form_layout.addWidget(QLabel("Badge (ID):"), row, 0)
        form_layout.addWidget(self.crud_badge_input, row, 1)
        row += 1
        
        # Separator for Logistics
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setFrameShadow(QFrame.Shadow.Sunken)
        form_layout.addWidget(sep, row, 0, 1, 2)
        row += 1
        
        lbl_logistics = QLabel("Default Logistics (Pre-defined):")
        font_l = lbl_logistics.font(); font_l.setBold(True)
        lbl_logistics.setFont(font_l)
        form_layout.addWidget(lbl_logistics, row, 0, 1, 2)
        row += 1

        form_layout.addWidget(QLabel("Default Pick Up:"), row, 0)
        form_layout.addWidget(self.def_pickup_combo, row, 1)
        row += 1
        form_layout.addWidget(QLabel("Default Drop Off:"), row, 0)
        form_layout.addWidget(self.def_dropoff_combo, row, 1)
        row += 1

        # Action Buttons
        button_layout = QHBoxLayout()
        #button_layout.addWidget(self.crud_new_button)
        button_layout.addWidget(self.crud_save_button)
        form_layout.addLayout(button_layout, row, 0, 1, 2)
        row += 1
        form_layout.addWidget(self.crud_delete_button, row, 0, 1, 2)
        row += 1
        #form_layout.addWidget(self.import_button, row, 0, 1, 2)

        form_group = create_group_box("Manage User", form_layout)
        form_group.setFixedWidth(400)

        # ---------------------------------------------------------
        # Right panel: Users Table
        # ---------------------------------------------------------
        table_panel = QWidget()
        table_layout = QVBoxLayout(table_panel)

        # Filter Panel
        filter_bar = QHBoxLayout()
        self.user_search_input = QLineEdit()
        self.user_search_input.setPlaceholderText("Search Name/Badge...")
        self.user_search_input.textChanged.connect(self._request_refresh)

        self.user_role_combo = QComboBox()
        self.user_role_combo.currentIndexChanged.connect(self._request_refresh)

        self.user_badge_prefix_input = QLineEdit()
        self.user_badge_prefix_input.setPlaceholderText("Badge prefix...")
        self.user_badge_prefix_input.textChanged.connect(self._request_refresh)

        user_reset_btn = QPushButton("Reset")
        user_reset_btn.clicked.connect(self.reset_filters)

        filter_bar.addWidget(self.user_search_input, 2)
        filter_bar.addWidget(QLabel("Role:"))
        filter_bar.addWidget(self.user_role_combo, 1)
        filter_bar.addWidget(QLabel("Badge Prefix:"))
        filter_bar.addWidget(self.user_badge_prefix_input, 1)
        filter_bar.addWidget(user_reset_btn)
        table_layout.addLayout(filter_bar)

        self.users_table = QTableWidget()
        self.users_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.users_table.setAlternatingRowColors(True)
        self.users_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.users_table.itemClicked.connect(self.load_user_to_crud_form)
        table_layout.addWidget(self.users_table)

        table_group = create_group_box("Registered Users List", table_layout)
        layout.addWidget(form_group)
        layout.addWidget(table_group)

        # Initial Load
        self.refresh_ui_data()
    def populate_role_combo(self):
        """Refreshes the Role dropdown from the DB Master List."""
        current_text = self.crud_role_input.currentText()
        self.crud_role_input.blockSignals(True)
        self.crud_role_input.clear()
        
        # Add a default blank
        self.crud_role_input.addItem("", None)
        
        # Fetch from roles table
        roles = db.get_roles(self.source)
        print(f"DEBUG: Cargando roles para {self.source}. Encontrados: {len(roles)}")
        for r in roles:
            self.crud_role_input.addItem(r["name"], r["id"])
            
        # Try to restore selection if text matches
        idx = self.crud_role_input.findText(current_text)
        if idx >= 0:
            self.crud_role_input.setCurrentIndex(idx)
            
        self.crud_role_input.blockSignals(False)
        
    def _request_refresh(self):
        self._debounce_timer.start(DEBOUNCE_MS)

    def _populate_role_filter(self):
        # Populate Role Filter
        self.user_role_combo.blockSignals(True)
        current_role = self.user_role_combo.currentText()
        self.user_role_combo.clear()
        self.user_role_combo.addItem("All Roles")
        all_users = db.get_all_users(self.source)
        roles = sorted(list(set(u["role"] for u in all_users if u.get("role"))))
        self.user_role_combo.addItems(roles)
        idx = self.user_role_combo.findText(current_role)
        if idx > 0:
            self.user_role_combo.setCurrentIndex(idx)
        self.user_role_combo.blockSignals(False)
        
        # Populate Location Combos (Default Pick/Drop)
        self._populate_location_combos(self.def_pickup_combo)
        self._populate_location_combos(self.def_dropoff_combo)

    def _populate_location_combos(self, combo: QComboBox):
        """Helper to fill location combos keeping current selection if possible."""
        current_data = combo.currentData()
        combo.blockSignals(True)
        combo.clear()
        combo.addItem("— None —", None)
        locations = db.get_locations(self.source)
        for loc in locations:
            combo.addItem(loc["pickup_location"], loc["pickup_location"])
        
        if current_data:
            idx = combo.findData(current_data)
            if idx >= 0:
                combo.setCurrentIndex(idx)
        combo.blockSignals(False)

    def reset_filters(self):
        with QSignalBlocker(self.user_search_input), QSignalBlocker(
            self.user_role_combo
        ), QSignalBlocker(self.user_badge_prefix_input):
            self.user_search_input.clear()
            self.user_role_combo.setCurrentIndex(0)
            self.user_badge_prefix_input.clear()
        self.load_users_table()

    def load_users_table(self):
        self._filter_state["text"] = self.user_search_input.text()
        self._filter_state["role"] = (
            self.user_role_combo.currentText()
            if self.user_role_combo.currentIndex() > 0
            else None
        )
        self._filter_state["badge_prefix"] = self.user_badge_prefix_input.text()

        # --- MODIFIED: Use the JOINED query to get defaults ---
        # Note: Filtering by defaults isn't requested, but we need to display them.
        # If filtering is active, we might need a more complex query or filter in Python.
        # Since get_users_filtered doesn't support defaults, we use get_users_with_defaults
        # and filter in Python for simplicity (assuming < 2000 users). 
        # For production with 50k+ users, we'd add defaults to get_users_filtered SQL.
        
        all_users_with_defaults = db.get_users_with_defaults(self.source)
        
        # Apply filters in Python
        filtered_users = []
        txt = (self._filter_state["text"] or "").lower()
        role_filter = self._filter_state["role"]
        pfx = (self._filter_state["badge_prefix"] or "").lower()

        for u in all_users_with_defaults:
            if txt and (txt not in u["name"].lower() and txt not in u["badge"].lower()):
                continue
            if role_filter and u["role"] != role_filter:
                continue
            if pfx and not u["badge"].lower().startswith(pfx):
                continue
            filtered_users.append(u)

        # Columns: ID, Name, Role, Badge, Default Pick, Default Drop
        headers = ["ID", "Name", "Role", "Badge", "Def. Pick Up", "Def. Drop Off"]
        self.users_table.setRowCount(len(filtered_users))
        self.users_table.setColumnCount(len(headers))
        self.users_table.setHorizontalHeaderLabels(headers)

        for row, user in enumerate(filtered_users):
            self.users_table.setItem(row, 0, QTableWidgetItem(str(user["id"])))
            self.users_table.setItem(row, 1, QTableWidgetItem(user["name"]))
            self.users_table.setItem(row, 2, QTableWidgetItem(user["role"]))
            self.users_table.setItem(row, 3, QTableWidgetItem(user["badge"]))
            
            # New Columns
            pu = user.get("pickup_location") or ""
            do = user.get("dropoff_location") or ""
            
            item_pu = QTableWidgetItem(pu)
            item_pu.setForeground(QColor("#1565C0") if pu else QColor("#9E9E9E")) # Blue if set
            self.users_table.setItem(row, 4, item_pu)
            
            item_do = QTableWidgetItem(do)
            item_do.setForeground(QColor("#1565C0") if do else QColor("#9E9E9E"))
            self.users_table.setItem(row, 5, item_do)

        self.users_table.setColumnHidden(0, True)  # hide ID column
        self.users_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        # Give more space to Logistics columns if needed
        self.users_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        self.users_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)

    def load_user_to_crud_form(self, item):
        row = item.row()
        self.current_user_id = int(self.users_table.item(row, 0).text())
        self.crud_name_input.setText(self.users_table.item(row, 1).text())
        
        # Load Role into Combo
        role_text = self.users_table.item(row, 2).text()
        idx = self.crud_role_input.findText(role_text)
        if idx >= 0:
            self.crud_role_input.setCurrentIndex(idx)
        else:
            # If the role isn't in the list (legacy data), we can't select it easily 
            # if editable is False. 
            # Ideally, migration handled this. Defaults to blank if not found.
            self.crud_role_input.setCurrentIndex(0)

        self.crud_badge_input.setText(self.users_table.item(row, 3).text())
        
        # Load Defaults from table (columns 4 and 5)
        # Note: We rely on the text in the table matching the data in the combo
        def_pickup = self.users_table.item(row, 4).text()
        def_dropoff = self.users_table.item(row, 5).text()
        
        idx_pu = self.def_pickup_combo.findData(def_pickup if def_pickup else None)
        self.def_pickup_combo.setCurrentIndex(idx_pu if idx_pu >= 0 else 0)
        
        idx_do = self.def_dropoff_combo.findData(def_dropoff if def_dropoff else None)
        self.def_dropoff_combo.setCurrentIndex(idx_do if idx_do >= 0 else 0)

    def clear_crud_form(self):
        self.current_user_id = None
        self.crud_name_input.clear()
        self.crud_role_input.setCurrentIndex(0)
        self.crud_badge_input.clear()
        self.def_pickup_combo.setCurrentIndex(0)
        self.def_dropoff_combo.setCurrentIndex(0)
        self.users_table.clearSelection()

    def save_crud_user(self):
        name = self.crud_name_input.text().strip()
        # Get Role from Combo Text (NOT ID)
        # This ensures we save "Driver" to the user table, keeping PlanStaffWidget compatible.
        role = self.crud_role_input.currentText().strip() 
        
        badge = self.crud_badge_input.text().strip()
        
        # Get defaults
        def_pickup = self.def_pickup_combo.currentData()
        def_dropoff = self.def_dropoff_combo.currentData()

        if not name or not role or not badge:
            QMessageBox.warning(self, "Incomplete Data", "Name, Role, and Badge are required.")
            return

        # 1. Save User Core Data
        if self.current_user_id:
            success, message = db.update_user(self.current_user_id, name, role, badge, self.source)
        else:
            success, message = db.add_user(name, role, badge, self.source)

        if success:
            # 2. Save Default Logistics (SSoT: user_locations where is_default=1)
            try:
                db.set_user_default_locations(badge, def_pickup, def_dropoff)
                message += "\nDefault locations updated."
            except Exception as e:
                message += f"\nWarning: Could not save locations ({e})"

            # 3. Sync Excel
            try:
                excel.refresh_excel_from_db(self.excel_file, self.source)
            except Exception as e:
                print(f"Auto-refresh failed: {e}")

            # 4. Refresh UI
            self._populate_role_filter() # Refresh roles in case user added a new one
            self.refresh_ui_data()       # Refresh table to show new columns
            self.populate_role_combo()
            self.users_changed.emit(self.source) # Notify Plan Staff tab

        box = QMessageBox(self)
        box.setIcon(QMessageBox.Icon.Information if success else QMessageBox.Icon.Warning)
        box.setWindowTitle("Success" if success else "Error")
        box.setText(message)
        box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
        box.exec()

    def delete_crud_user(self):
        if not self.current_user_id:
            QMessageBox.warning(self, "No Selection", "Please select a user to delete.")
            return

        badge_to_remove = self.crud_badge_input.text().strip()

        confirm = QMessageBox(self)
        confirm.setIcon(QMessageBox.Icon.Question)
        confirm.setWindowTitle("Confirm User Deletion")
        confirm.setText(f"Are you sure you want to delete user '{self.crud_name_input.text()}'?")
        yes_btn = confirm.addButton("Yes", QMessageBox.ButtonRole.YesRole)
        confirm.addButton("No", QMessageBox.ButtonRole.NoRole)
        confirm.exec()

        if confirm.clickedButton() == yes_btn:
            success_db, message_db = db.delete_user(self.current_user_id)
            final_msg = message_db

            if success_db:
                # Also clean defaults from DB (Optional but clean)
                # Note: db.delete_user only deletes from 'users'. 
                # Ideally, foreign keys would handle this, or we run a manual delete on user_locations.
                # For now, leaving orphan defaults is harmless, but we could add:
                # db.set_user_default_locations(badge_to_remove, None, None) 
                
                success_excel, msg_excel = excel.remove_user_from_excel(self.excel_file, badge_to_remove)
                if success_excel:
                    final_msg += f"\n\nRemoved from Excel: {badge_to_remove}"
                else:
                    final_msg += f"\n\nWarning: Excel sync failed ({msg_excel})"

                db.log_event(self.logged_username, self.source, "USER_DELETE", f"Deleted {badge_to_remove}")
                
                self.refresh_ui_data()
                self.users_changed.emit(self.source)

            QMessageBox.information(self, "Deletion Result", final_msg)

    def import_users_from_excel(self):
        # (Sin cambios en esta función)
        try:
            inserted, skipped, upserts = excel.import_excel_to_db(self.excel_file, self.source)
            db.log_event(self.logged_username, self.source, "DATA_IMPORT", f"users_inserted={inserted}; skipped={skipped}")
            QMessageBox.information(self, "Import Complete", f"Imported {inserted} new users.\nSkipped {skipped} existing.\nUpserted {upserts} schedules.")
            self.refresh_ui_data()
            self.users_changed.emit(self.source)
            self.import_done.emit(self.source)
        except ValueError as ve:
            QMessageBox.critical(self, "Invalid Excel", str(ve))

    def refresh_ui_data(self):
        self.populate_role_combo()
        self._populate_role_filter()
        self.load_users_table()
        self.clear_crud_form()

# -------------------------------------------------------------
# Widget: Shift Types Admin (Admin and Site Managers)
# -------------------------------------------------------------
class ShiftTypeAdminWidget(QWidget):
    types_changed = pyqtSignal(str)  # emits source

    def __init__(self, source: str, excel_file: str, logged_username: str):
        super().__init__()
        self.source = source
        self.excel_file = excel_file
        self.logged_username = logged_username or "Unknown"
        self.current_type_id = None
        self.current_old_code = None

        self._filter_state = {}
        self._debounce_timer = QTimer(self)
        self._debounce_timer.setSingleShot(True)
        self._debounce_timer.timeout.connect(self.refresh_table)

        layout = QHBoxLayout(self)

        # Left: form
        form_layout = QGridLayout()
        self.name_input = QLineEdit()
        self.code_input = QLineEdit()
        self.color_display = QLineEdit()
        self.color_display.setReadOnly(True)
        self.pick_color_btn = QPushButton("🎨 Pick Color")
        self.pick_color_btn.clicked.connect(self.pick_color)

        self.in_time_edit = QTimeEdit()
        self.in_time_edit.setDisplayFormat("HH:mm")
        self.in_time_edit.setTime(QTime(8, 0))
        self.out_time_edit = QTimeEdit()
        self.out_time_edit.setDisplayFormat("HH:mm")
        self.out_time_edit.setTime(QTime(17, 0))

        self.new_btn = QPushButton("✨ New Shift Type")
        self.new_btn.clicked.connect(self.clear_form)
        self.save_btn = QPushButton("💾 Save")
        self.save_btn.clicked.connect(self.save_type)
        self.delete_btn = QPushButton("❌ Delete")
        self.delete_btn.clicked.connect(self.delete_type)

        # Button variants
        self.new_btn.setProperty("variant", "secondary")
        self.save_btn.setProperty("variant", "primary")
        self.delete_btn.setProperty("danger", True)

        form_layout.addWidget(QLabel("Name:"), 0, 0)
        form_layout.addWidget(self.name_input, 0, 1)
        form_layout.addWidget(QLabel("Code (short):"), 1, 0)
        form_layout.addWidget(self.code_input, 1, 1)
        form_layout.addWidget(QLabel("Color:"), 2, 0)
        h_color = QHBoxLayout()
        h_color.addWidget(self.color_display)
        h_color.addWidget(self.pick_color_btn)
        form_layout.addLayout(h_color, 2, 1)
       # --- MODIFICACIÓN: Guardamos referencias a los Labels ---
        self.lbl_in = QLabel("IN time (HH:MM):")
        self.lbl_out = QLabel("OUT time (HH:MM):")

        form_layout.addWidget(self.lbl_in, 3, 0)
        form_layout.addWidget(self.in_time_edit, 3, 1)
        form_layout.addWidget(self.lbl_out, 4, 0)
        form_layout.addWidget(self.out_time_edit, 4, 1)

        # --- NUEVO: Checkbox 'Treat as OFF' ---
        self.is_off_check = QCheckBox("Treat as 'OFF' (Non-working day)")
        self.is_off_check.setToolTip(
            "Check this if the shift (e.g., Vacation, Sick Leave) should be treated\n"
            "as a day OFF for transport and onsite-stay reports."
        )
        form_layout.addWidget(QLabel("Behavior:"), 5, 0)
        form_layout.addWidget(self.is_off_check, 5, 1)
        # --------------------------------------

        actions = QHBoxLayout()
        #actions.addWidget(self.new_btn)
        actions.addWidget(self.save_btn)
        actions.addWidget(self.delete_btn)
        
        # NOTA: Cambiamos el row de 5 a 6 para hacer espacio
        form_layout.addLayout(actions, 6, 0, 1, 2)

        form_group = create_group_box("Shift Type", form_layout)
        form_group.setFixedWidth(420)

        # Right: table
        table_panel = QWidget()
        table_layout = QVBoxLayout(table_panel)

        # --- Filter Panel for Shift Types ---
        st_filter_bar = QHBoxLayout()
        self.st_search_input = QLineEdit()
        self.st_search_input.setPlaceholderText("Search Name/Code...")
        self.st_search_input.textChanged.connect(self._request_refresh)

        self.st_time_from = QTimeEdit(QTime(0, 0))
        self.st_time_from.setDisplayFormat("HH:mm")
        self.st_time_from.timeChanged.connect(self._request_refresh)

        self.st_time_to = QTimeEdit(QTime(23, 59))
        self.st_time_to.setDisplayFormat("HH:mm")
        self.st_time_to.timeChanged.connect(self._request_refresh)

        self.st_usage_combo = QComboBox()
        self.st_usage_combo.addItems(["All", "In use", "Not in use"])
        self.st_usage_combo.currentIndexChanged.connect(self._request_refresh)

        st_reset_btn = QPushButton("Reset")
        st_reset_btn.clicked.connect(self.reset_filters)

        st_filter_bar.addWidget(self.st_search_input, 2)
        st_filter_bar.addWidget(QLabel("IN time from:"))
        st_filter_bar.addWidget(self.st_time_from)
        st_filter_bar.addWidget(QLabel("to:"))
        st_filter_bar.addWidget(self.st_time_to)
        st_filter_bar.addWidget(QLabel("Usage:"))
        st_filter_bar.addWidget(self.st_usage_combo, 1)
        st_filter_bar.addWidget(st_reset_btn)
        table_layout.addLayout(st_filter_bar)

        self.types_table = QTableWidget()
        self.types_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.types_table.setAlternatingRowColors(True)
        self.types_table.itemClicked.connect(self.load_to_form)
        table_layout.addWidget(self.types_table)

        table_group = create_group_box(f"{self.source} Shift Types", table_layout)
        layout.addWidget(form_group)
        layout.addWidget(table_group)
        # --- MODIFICACIÓN: Conectar Checkbox a la visibilidad ---
        self.is_off_check.toggled.connect(self.toggle_time_inputs)
        self.reset_filters()

    def _request_refresh(self):
        self._debounce_timer.start(DEBOUNCE_MS)

    def reset_filters(self):
        with QSignalBlocker(self.st_search_input), QSignalBlocker(
            self.st_time_from
        ), QSignalBlocker(self.st_time_to), QSignalBlocker(self.st_usage_combo):
            self.st_search_input.clear()
            self.st_time_from.setTime(QTime(0, 0))
            self.st_time_to.setTime(QTime(23, 59))
            self.st_usage_combo.setCurrentIndex(0)
        self.refresh_table()

    def toggle_time_inputs(self, checked):
        """
        Oculta o muestra los inputs de tiempo según si es 'Treat as OFF'.
        Si es OFF (checked=True), ocultamos los tiempos (Visible=False).
        """
        is_working = not checked
        
        # Mostrar u ocultar etiquetas y campos
        self.lbl_in.setVisible(is_working)
        self.in_time_edit.setVisible(is_working)
        self.lbl_out.setVisible(is_working)
        self.out_time_edit.setVisible(is_working)
        
        # Opcional: Si se marca como OFF, limpiamos visualmente a 00:00
        if checked:
            self.in_time_edit.setTime(QTime(0, 0))
            self.out_time_edit.setTime(QTime(0, 0))
        else:
            # Si se desmarca (vuelve a ser laboral) y está en 00:00, restaurar defaults
            if self.in_time_edit.time().toString("HH:mm") == "00:00":
                self.in_time_edit.setTime(QTime(8, 0))
                self.out_time_edit.setTime(QTime(17, 0))
    
    def pick_color(self):
        color = QColorDialog.getColor(
            QColor(self.color_display.text() or "#FFC000"), self, "Pick a Color"
        )
        if color.isValid():
            self.color_display.setText(color.name())

    def load_to_form(self, item):
        row = item.row()
        self.current_type_id = int(self.types_table.item(row, 0).text())
        self.name_input.setText(self.types_table.item(row, 1).text())
        
        code = self.types_table.item(row, 2).text()
        self.code_input.setText(code)
        
        color_hex = self.types_table.item(row, 3).text()
        self.color_display.setText(color_hex)
        
        # Nota: La columna 4 es el previo de color, saltamos a la 5 y 6
        in_time = self.types_table.item(row, 5).text()
        out_time = self.types_table.item(row, 6).text()
        self.in_time_edit.setTime(QTime.fromString(in_time, "HH:mm"))
        self.out_time_edit.setTime(QTime.fromString(out_time, "HH:mm"))
        self.current_old_code = code

        # --- NUEVO: Cargar estado 'is_off' desde la BD ---
        # Obtenemos todos los tipos para buscar el atributo 'is_off' del actual
        all_types = db.get_shift_types(self.source)
        record = next((t for t in all_types if t['id'] == self.current_type_id), None)
        
        if record:
            # Convertimos 1/0 a True/False
            val = bool(record.get('is_off', 0))
            self.is_off_check.setChecked(val)
        else:
            self.is_off_check.setChecked(False)
            
        # --- MODIFICACIÓN: Forzar actualización visual ---
        self.toggle_time_inputs(self.is_off_check.isChecked())

    def clear_form(self):
        self.current_type_id = None
        self.current_old_code = None
        self.name_input.clear()
        self.code_input.clear()
        self.color_display.setText("#FFC000")
        self.in_time_edit.setTime(QTime(8, 0))
        self.out_time_edit.setTime(QTime(17, 0))
        
       # --- NUEVO ---
        self.is_off_check.setChecked(False) 
        
        # --- MODIFICACIÓN: Asegurar que los tiempos sean visibles al limpiar ---
        self.toggle_time_inputs(False) 
        
        self.types_table.clearSelection()

    def save_type(self):
        name = self.name_input.text().strip()
        code = self.code_input.text().strip().upper()
        color_hex = self.color_display.text().strip() or "#FFC000"
        in_time = self.in_time_edit.time().toString("HH:mm")
        out_time = self.out_time_edit.time().toString("HH:mm")
        
        # --- NUEVO: Leer valor del checkbox ---
        is_off_val = self.is_off_check.isChecked()

        if not name or not code:
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Warning)
            box.setWindowTitle("Incomplete Data")
            box.setText("Name and Code are required.")
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
            box.exec()
            return

        if self.current_type_id:
            # ACTUALIZAR
            ok, msg, old_code, new_code = db.update_shift_type(
                self.current_type_id,
                self.source,
                name,
                code,
                color_hex,
                in_time,
                out_time,
                is_off=is_off_val  # <--- Pasamos el nuevo parámetro
            )
            if ok:
                if old_code and new_code and old_code != new_code:
                    excel.apply_shift_type_update_to_excel(
                        self.excel_file, self.source, old_code, new_code, color_hex
                    )
                db.log_event(
                    self.logged_username,
                    self.source,
                    "SHIFT_TYPE_UPDATE",
                    f"{old_code} -> {new_code} | Off={is_off_val}"
                )
                self.types_changed.emit(self.source)
            
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Information if ok else QMessageBox.Icon.Warning)
            box.setWindowTitle("Save" if ok else "Error")
            box.setText(msg)
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
            box.exec()
        else:
            # CREAR
            ok, msg = db.create_shift_type(
                self.source, name, code, color_hex, in_time, out_time, 
                is_off=is_off_val # <--- Pasamos el nuevo parámetro
            )
            if ok:
                db.log_event(
                    self.logged_username,
                    self.source,
                    "SHIFT_TYPE_CREATE",
                    f"{code} | Off={is_off_val}"
                )
                self.types_changed.emit(self.source)
            
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Information if ok else QMessageBox.Icon.Warning)
            box.setWindowTitle("Create" if ok else "Error")
            box.setText(msg)
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
            box.exec()

        self.refresh_table()
        self.clear_form()

    def delete_type(self):
        if not self.current_type_id:
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Warning)
            box.setWindowTitle("No Selection")
            box.setText("Please select a shift type in the table to delete.")
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
            box.exec()
            return

        confirm = QMessageBox(self)
        confirm.setIcon(QMessageBox.Icon.Question)
        confirm.setWindowTitle("Confirm Deletion")
        confirm.setText(f"Are you sure you want to delete {self.name_input.text()}?")
        yes_btn = confirm.addButton("Yes", QMessageBox.ButtonRole.YesRole)
        confirm.addButton("No", QMessageBox.ButtonRole.NoRole)
        confirm.exec()

        if confirm.clickedButton() == yes_btn:
            ok, msg, source, code = db.delete_shift_type(self.current_type_id)
            if ok:
                db.log_event(
                    self.logged_username, self.source, "SHIFT_TYPE_DELETE", f"{code}"
                )
                self.types_changed.emit(self.source)
            box = QMessageBox(self)
            box.setIcon(
                QMessageBox.Icon.Information if ok else QMessageBox.Icon.Warning
            )
            box.setWindowTitle("Delete" if ok else "Cannot delete")
            box.setText(msg)
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
            box.exec()

            self.refresh_table()
            self.clear_form()

    def refresh_table(self):
        self._filter_state["text"] = self.st_search_input.text()
        self._filter_state["in_from"] = self.st_time_from.time().toString("HH:mm")
        self._filter_state["in_to"] = self.st_time_to.time().toString("HH:mm")
        self._filter_state["usage"] = (
            self.st_usage_combo.currentText()
            if self.st_usage_combo.currentIndex() > 0
            else None
        )

        types = db.get_shift_types_filtered(
            source=self.source,
            text=self._filter_state["text"],
            in_from=self._filter_state["in_from"],
            in_to=self._filter_state["in_to"],
            usage=self._filter_state["usage"],
        )
        headers = ["ID", "Name", "Code", "Color", "Color Preview", "IN", "OUT"]
        self.types_table.setRowCount(len(types))
        self.types_table.setColumnCount(len(headers))
        self.types_table.setHorizontalHeaderLabels(headers)
        for r, t in enumerate(types):
            self.types_table.setItem(r, 0, QTableWidgetItem(str(t["id"])))
            self.types_table.setItem(r, 1, QTableWidgetItem(t["name"]))
            self.types_table.setItem(r, 2, QTableWidgetItem(t["code"]))
            self.types_table.setItem(r, 3, QTableWidgetItem(t["color_hex"]))

            # Add color preview item
            preview_item = QTableWidgetItem()
            preview_item.setIcon(_create_color_icon(t["color_hex"]))
            self.types_table.setItem(r, 4, preview_item)

            self.types_table.setItem(r, 5, QTableWidgetItem(t["in_time"]))
            self.types_table.setItem(r, 6, QTableWidgetItem(t["out_time"]))

        self.types_table.setColumnHidden(0, True)
        self.types_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch
        )
        # Adjust preview column width
        preview_col_index = headers.index("Color Preview")
        self.types_table.horizontalHeader().setSectionResizeMode(
            preview_col_index, QHeaderView.ResizeMode.ResizeToContents
        )
        self.types_table.setColumnWidth(preview_col_index, 40)


# -------------------------------------------------------------
# NEW Widget: Location Admin
# -------------------------------------------------------------
class LocationAdminWidget(QWidget):
    """
    Si scope_source es 'RGM' o 'Newmont' -> modo estándar (solo su empresa).
    Si scope_source es None -> modo Admin (todas, con filtro y selector de dueño).
    """

    locations_changed = pyqtSignal()

    def __init__(self, scope_source: str | None = None):
        super().__init__()
        self.scope_source = scope_source  # None = admin; "RGM"/"Newmont" = normal
        self.loc_id = None
        self._filter_state = {}
        self._debounce_timer = QTimer(self)
        self._debounce_timer.setSingleShot(True)
        self._debounce_timer.timeout.connect(self._reload_table)

        layout = QHBoxLayout(self)

        # --- Formulario ---
        form = QGridLayout()
        row = 0

        self.loc_input = QLineEdit()
        form.addWidget(QLabel("Location name:"), row, 0)
        form.addWidget(self.loc_input, row, 1)
        row += 1

        # En modo admin, permitir elegir dueño (source) del registro
        self.owner_combo = None
        if self.scope_source is None:
            self.owner_combo = QComboBox()
            self.owner_combo.addItems(["RGM", "Newmont"])
            form.addWidget(QLabel("Owner (Source):"), row, 0)
            form.addWidget(self.owner_combo, row, 1)
            row += 1

        btn_new = QPushButton("✨ New")
        btn_save = QPushButton("💾 Save")
        btn_del = QPushButton("❌ Delete")
        btn_save.setProperty("variant", "primary")
        btn_del.setProperty("danger", True)
        h = QHBoxLayout()
        #h.addWidget(btn_new)
        h.addWidget(btn_save)
        h.addWidget(btn_del)
        form.addLayout(h, row, 0, 1, 2)

        form_group = create_group_box("Location", form)
        form_group.setFixedWidth(420)

        # --- Tabla ---
        table_panel = QWidget()
        table_box = QVBoxLayout(table_panel)

        controls_row = QHBoxLayout()

        self.loc_search_input = QLineEdit()
        self.loc_search_input.setPlaceholderText("Search location...")
        self.loc_search_input.textChanged.connect(self._request_refresh)
        controls_row.addWidget(self.loc_search_input, 1)

        # Filtro por empresa (solo en Admin)
        self.filter_combo = None
        if self.scope_source is None:
            self.filter_combo = QComboBox()
            self.filter_combo.addItem("All Sources", None)
            self.filter_combo.addItem("RGM", "RGM")
            self.filter_combo.addItem("Newmont", "Newmont")
            self.filter_combo.currentIndexChanged.connect(self._request_refresh)
            controls_row.addWidget(QLabel("Source:"))
            controls_row.addWidget(self.filter_combo)

        self.sort_combo = QComboBox()
        self.sort_combo.addItems(["Sort by Name", "Sort by Source"])
        self.sort_combo.currentIndexChanged.connect(self._request_refresh)
        controls_row.addWidget(self.sort_combo)

        reset_btn = QPushButton("Reset")
        reset_btn.clicked.connect(self._reset_filters)
        controls_row.addWidget(reset_btn)

        table_box.addLayout(controls_row)

        self.loc_table = QTableWidget()
        self.loc_table.setAlternatingRowColors(True)
        table_box.addWidget(self.loc_table)

        table_group = create_group_box("Locations", table_box)

        layout.addWidget(form_group)
        layout.addWidget(table_group)

        # Eventos
        btn_new.clicked.connect(self._new_loc)
        btn_save.clicked.connect(self._save_loc)
        btn_del.clicked.connect(self._delete_loc)
        self.loc_table.itemClicked.connect(self._load_to_form)

        self._reset_filters()

    # --- helpers ---
    def _request_refresh(self):
        self._debounce_timer.start(DEBOUNCE_MS)

    def _reset_filters(self):
        with QSignalBlocker(self.loc_search_input), QSignalBlocker(self.sort_combo):
            self.loc_search_input.clear()
            self.sort_combo.setCurrentIndex(0)
            if self.filter_combo:
                with QSignalBlocker(self.filter_combo):
                    self.filter_combo.setCurrentIndex(0)
        self._reload_table()

    def _effective_filter_source(self) -> str | None:
        if self.scope_source is not None:
            return self.scope_source  # perfil normal
        # admin
        if self.filter_combo is None:
            return None
        return self.filter_combo.currentData()

    def _reload_table(self):
        src = self._effective_filter_source()
        text = self.loc_search_input.text()
        sort = "source" if self.sort_combo.currentIndex() == 1 else "name"

        self._filter_state["source"] = src
        self._filter_state["text"] = text
        self._filter_state["sort"] = sort

        rows = db.get_locations_filtered(source=src, text=text, sort_by=sort)

        if self.scope_source is None:
            headers = ["ID", "Source", "Location"]
        else:
            headers = ["ID", "Location"]
        self.loc_table.setRowCount(len(rows))
        self.loc_table.setColumnCount(len(headers))
        self.loc_table.setHorizontalHeaderLabels(headers)
        for r, row in enumerate(rows):
            self.loc_table.setItem(r, 0, QTableWidgetItem(str(row["id"])))
            if self.scope_source is None:
                self.loc_table.setItem(r, 1, QTableWidgetItem(row["source"]))
                self.loc_table.setItem(r, 2, QTableWidgetItem(row["pickup_location"]))
            else:
                self.loc_table.setItem(r, 1, QTableWidgetItem(row["pickup_location"]))
        self.loc_table.setColumnHidden(0, True)
        self.loc_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch
        )

    def _new_loc(self):
        self.loc_id = None
        self.loc_input.clear()
        if self.owner_combo is not None:
            self.owner_combo.setCurrentIndex(0)
        self.loc_table.clearSelection()

    def _save_loc(self):
        name = (self.loc_input.text() or "").strip()
        if not name:
            QMessageBox.warning(self, "Input Error", "Location name cannot be empty.")
            return

        # Determinar 'source' destino del registro
        if self.scope_source is None:
            dest_source = self.owner_combo.currentText() if self.owner_combo else "RGM"
        else:
            dest_source = self.scope_source

        if self.loc_id:
            if self.scope_source is None:
                ok, msg = db.update_location_admin(self.loc_id, name, dest_source)
            else:
                ok, msg = db.update_location(self.loc_id, name, dest_source)
        else:
            ok, msg = db.create_location(name, dest_source)

        QMessageBox.information(self, "Location", msg)
        self._reload_table()
        self.locations_changed.emit()
        self._new_loc()

    def _delete_loc(self):
        if not self.loc_id:
            QMessageBox.warning(self, "Location", "Please select a row.")
            return
        if self.scope_source is None:
            ok, msg = db.delete_location_admin(self.loc_id)
        else:
            ok, msg = db.delete_location(self.loc_id, self.scope_source)
        QMessageBox.information(self, "Location", msg)
        self._reload_table()
        self.locations_changed.emit()
        self._new_loc()

    def _load_to_form(self, item):
        row = item.row()
        self.loc_id = int(self.loc_table.item(row, 0).text())
        if self.scope_source is None:
            # Admin: columnas = ID | Source | Location
            self.loc_input.setText(self.loc_table.item(row, 2).text())
            if self.owner_combo is not None:
                src = self.loc_table.item(row, 1).text()
                idx = self.owner_combo.findText(src)
                if idx >= 0:
                    self.owner_combo.setCurrentIndex(idx)
        else:
            # Normal: columnas = ID | Location
            self.loc_input.setText(self.loc_table.item(row, 1).text())


# -------------------------------------------------------------
# Widget: Audit Log (visible for Admin; reusable otherwise)
# -------------------------------------------------------------
class AuditLogWidget(QWidget):
    def __init__(self, source: str | None):
        super().__init__()
        self.source = source
        layout = QVBoxLayout(self)
        self.audit_table = QTableWidget()
        layout.addWidget(self.audit_table)

        refresh_btn = QPushButton("🔄 Refresh")
        refresh_btn.setProperty("variant", "secondary")
        refresh_btn.clicked.connect(self.load_audit_log_data)
        layout.addWidget(refresh_btn, alignment=Qt.AlignmentFlag.AlignRight)

        self.load_audit_log_data()

    def load_audit_log_data(self):
        events = db.get_audit_log(source=self.source)
        headers = ["Timestamp", "User", "Source", "Action", "Detail"]
        self.audit_table.setRowCount(len(events))
        self.audit_table.setColumnCount(len(headers))
        self.audit_table.setHorizontalHeaderLabels(headers)

        for r, ev in enumerate(events):
            self.audit_table.setItem(r, 0, QTableWidgetItem(ev.get("ts", "")))
            self.audit_table.setItem(r, 1, QTableWidgetItem(ev.get("username", "")))
            self.audit_table.setItem(r, 2, QTableWidgetItem(ev.get("source", "")))
            self.audit_table.setItem(r, 3, QTableWidgetItem(ev.get("action_type", "")))
            self.audit_table.setItem(r, 4, QTableWidgetItem(ev.get("detail", "")))

        self.audit_table.setAlternatingRowColors(True)
        self.audit_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch
        )


# -------------------------------------------------------------
# Main window (normal profile: RGM or Newmont)
# -------------------------------------------------------------
class MainWindow(QMainWindow):
    logout_signal = pyqtSignal()

    def __init__(
        self,
        user_role,
        excel_file,
        logged_username=None,
        can_manage_shift_types: bool = False,
        preloaded_data=None # [CAMBIO 6] Nuevo argumento opcional
    ):
        super().__init__()
        self.user_role = user_role  # RGM or Newmont. Used as 'source'
        self.excel_file = excel_file
        self.logged_username = logged_username or "Unknown"
        self.can_manage_shift_types = bool(can_manage_shift_types)

        db.setup_database()
        db.log_event(
            self.logged_username,
            self.user_role,
            "USER_LOGIN",
            f"Excel={self.excel_file}",
        )

        self.setWindowTitle(
            f"👨‍✈️ Operations Manager - Profile: {self.user_role} | User: {self.logged_username}"
        )
        self.setGeometry(100, 100, 1200, 800)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Top bar
        top_layout = QHBoxLayout()
        title_label = QLabel(f"Transport & Operations Manager ({self.user_role})")
        font = title_label.font()
        font.setPointSize(20)
        font.setBold(True)
        title_label.setFont(font)

        self.logged_user_label = QLabel(f"👤 {self.logged_username}")
        lu_font = self.logged_user_label.font()
        lu_font.setBold(True)
        self.logged_user_label.setFont(lu_font)
        self.logged_user_label.setStyleSheet("padding: 0 12px;")

        logout_button = QPushButton("🔒 Log Out")
        logout_button.setFixedWidth(150)
        logout_button.setProperty("variant", "text")
        logout_button.clicked.connect(self.handle_logout)

        top_layout.addWidget(title_label)
        top_layout.addStretch()
        top_layout.addWidget(self.logged_user_label)
        top_layout.addWidget(logout_button)
        main_layout.addLayout(top_layout)

        # Tabs
        tabs = QTabWidget()
        main_layout.addWidget(tabs)

        # 1) Plan Staff (preview/register/reports)
        # [CAMBIO 7] Pasamos el preloaded_data al widget
        self.plan_widget = PlanStaffWidget(
            self.user_role, 
            self.excel_file, 
            self.logged_username,
            preloaded_data=preloaded_data 
        )
        
        self.plan_widget.excel_path_changed.connect(self._on_excel_path_changed)
        tabs.addTab(self.plan_widget, "📅 Plan Staff & Reports")

        # 2) Rotation History (new tab, no ID column)
        self.rotation_widget = RotationHistoryWidget(
            created_by=self.logged_username
        )  # ← NUEVO: solo mis registros
        tabs.addTab(self.rotation_widget, "🔁 Rotation History")
        # Refresh rotation history whenever plan saves a rotation
        self.plan_widget.rotation_changed.connect(self.rotation_widget.refresh_data)

        # 3) Users CRUD
        self.crud_widget = CrudWidget(
            self.user_role, self.excel_file, self.logged_username
        )
        tabs.addTab(self.crud_widget, "👥 Users (CRUD)")

        # 4) Shift Types (only if user can manage them)
        if self.can_manage_shift_types:
            self.shift_types_widget = ShiftTypeAdminWidget(
                self.user_role, self.excel_file, self.logged_username
            )
            tabs.addTab(self.shift_types_widget, f"⚙️ {self.user_role} Shift Types")
            # Refresh combos/preview when shift types change
            self.shift_types_widget.types_changed.connect(
                lambda src: self.plan_widget.refresh_ui_data()
            )

       # 5) Locations admin
        self.location_widget = LocationAdminWidget(scope_source=self.user_role)
        tabs.addTab(self.location_widget, "📍 Location")
        self.location_widget.locations_changed.connect(
            self.plan_widget.load_location_options
        )
        self.location_widget.locations_changed.connect(
            lambda: self.crud_widget._populate_location_combos(self.crud_widget.def_pickup_combo)
        )
        self.location_widget.locations_changed.connect(
            lambda: self.crud_widget._populate_location_combos(self.crud_widget.def_dropoff_combo)
        )

        # --- CHANGE: ADD ROLES TAB ---
        self.role_widget = RoleAdminWidget(scope_source=self.user_role)
        tabs.addTab(self.role_widget, "👔 Roles / Dept")
        
        # Update CrudWidget when roles change
        self.role_widget.roles_changed.connect(
            self.crud_widget.populate_role_combo
        )
        # Update Filter combo in CrudWidget too
        self.role_widget.roles_changed.connect(
            self.crud_widget._populate_role_filter
        )

        # 6) Settings Tab
        self.settings_widget = ReportSettingsWidget(
            self.logged_username, self.user_role
        )
        tabs.addTab(self.settings_widget, "⚙️ Settings")

        # Hot sync
        self.crud_widget.users_changed.connect(
            lambda src: (
                self.plan_widget.refresh_ui_data() if src == self.user_role else None
            )
        )
        self.crud_widget.import_done.connect(
            lambda src: (
                self.plan_widget.refresh_ui_data() if src == self.user_role else None
            )
        )

    def _sync_after_users_changed(self, src: str):
        if src == self.user_role:
            self.plan_widget.refresh_users_only()
            QApplication.processEvents()  # ensure UI repaints

    def handle_logout(self):
        self.logout_signal.emit()
        self.close()
        
        
        
  
    
    def _on_excel_path_changed(self, source: str, new_path: str) -> None:
        """Actualiza la ruta del Excel en todos los widgets cuando cambia."""
        if source != self.user_role:
            return

        self.excel_file = new_path

        # Actualizar CRUD Widget si existe
        if hasattr(self, "crud_widget"):
            self.crud_widget.excel_file = new_path

        # Actualizar Shift Types Widget si existe
        if getattr(self, "can_manage_shift_types", False) and hasattr(self, "shift_types_widget"):
            self.shift_types_widget.excel_file = new_path


# -------------------------------------------------------------
# Administrator window (unified access)
# -------------------------------------------------------------
class AdminMainWindow(QMainWindow):
    logout_signal = pyqtSignal()

    def __init__(self, logged_username: str, rgm_excel: str, newmont_excel: str):
        super().__init__()
        self.logged_username = logged_username or "admin"

        db.setup_database()
        db.log_event(
            self.logged_username,
            "Administrator",
            "USER_LOGIN",
            f"Access to admin console | RGM={rgm_excel} | Newmont={newmont_excel}",
        )

        self.setWindowTitle(f"🛡️ Administrator Console | User: {self.logged_username}")
        self.setGeometry(100, 100, 1400, 900)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Top bar
        top_layout = QHBoxLayout()
        title_label = QLabel("Unified Access — RGM & Newmont")
        font = title_label.font()
        font.setPointSize(20)
        font.setBold(True)
        title_label.setFont(font)

        self.logged_user_label = QLabel(f"👤 {self.logged_username} (Administrator)")
        lu_font = self.logged_user_label.font()
        lu_font.setBold(True)
        self.logged_user_label.setFont(lu_font)
        self.logged_user_label.setStyleSheet("padding: 0 12px;")

        logout_button = QPushButton("🔒 Log Out")
        logout_button.setFixedWidth(150)
        logout_button.setProperty("variant", "text")
        logout_button.clicked.connect(self.handle_logout)

        top_layout.addWidget(title_label)
        top_layout.addStretch()
        top_layout.addWidget(self.logged_user_label)
        top_layout.addWidget(logout_button)
        main_layout.addLayout(top_layout)

        # Tabs
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        # 1) RGM CRUD
        self.rgm_crud = CrudWidget("RGM", rgm_excel, self.logged_username)
        self.tabs.addTab(self.rgm_crud, "👥 RGM CRUD")

        # 2) RGM Plan Staff
        self.rgm_plan = PlanStaffWidget("RGM", rgm_excel, self.logged_username)
        self.tabs.addTab(self.rgm_plan, "📅 RGM Plan Staff")


        # 3) Newmont CRUD
        self.nm_crud = CrudWidget("Newmont", newmont_excel, self.logged_username)
        self.tabs.addTab(self.nm_crud, "👥 Newmont CRUD")

        # 4) Newmont Plan Staff
        self.nm_plan = PlanStaffWidget("Newmont", newmont_excel, self.logged_username)
        self.tabs.addTab(self.nm_plan, "📅 Newmont Plan Staff")

        self.rgm_plan.excel_path_changed.connect(self._on_excel_path_changed)
        self.nm_plan.excel_path_changed.connect(self._on_excel_path_changed)

        # 5) Rotation History (global; no ID column)
        self.rotation_history = RotationHistoryWidget(
            created_by=None
        )  # AHORA (modo admin = None → sin filtro)
        self.tabs.addTab(self.rotation_history, "🔁 Rotation History")
        # Refresh when either plan tab writes a rotation
        self.rgm_plan.rotation_changed.connect(self.rotation_history.refresh_data)
        self.nm_plan.rotation_changed.connect(self.rotation_history.refresh_data)

        # 6) Audit Log (global)
        audit_all = AuditLogWidget(source=None)
        self.tabs.addTab(audit_all, "📝 Audit Log")

        # 7) Shift Types (Admin for both sites)
        self.rgm_types = ShiftTypeAdminWidget("RGM", rgm_excel, self.logged_username)
        self.tabs.addTab(self.rgm_types, "⚙️ RGM Shift Types")

        self.nm_types = ShiftTypeAdminWidget(
            "Newmont", newmont_excel, self.logged_username
        )
        self.tabs.addTab(self.nm_types, "⚙️ Newmont Shift Types")

        # 8) Locations (global admin)
        self.location_admin = LocationAdminWidget(scope_source=None)
        self.tabs.addTab(self.location_admin, "📍 Locations")
        # refresh dropdowns on both plan tabs when the master list changes
        self.location_admin.locations_changed.connect(
            lambda: self.rgm_plan.load_location_options()
        )
        self.location_admin.locations_changed.connect(
            lambda: self.nm_plan.load_location_options()
        )

        # 9) Settings
        settings_container = QWidget()
        settings_layout = QVBoxLayout(settings_container)
        settings_tabs = QTabWidget()
        settings_layout.addWidget(settings_tabs)

        rgm_settings = ReportSettingsWidget(self.logged_username, "RGM")
        settings_tabs.addTab(rgm_settings, "RGM Report Settings")

        nm_settings = ReportSettingsWidget(self.logged_username, "Newmont")
        settings_tabs.addTab(nm_settings, "Newmont Report Settings")

        self.tabs.addTab(settings_container, "⚙️ Settings")

        # Hot sync
        # CAMBIO: Usamos refresh_ui_data() en lugar de refresh_users_only()
        # para forzar la recarga de la grilla (tabla) y que desaparezca la fila borrada.

        self.rgm_crud.users_changed.connect(lambda src: self.rgm_plan.refresh_ui_data())
        self.rgm_crud.import_done.connect(lambda src: self.rgm_plan.refresh_ui_data())

        self.nm_crud.users_changed.connect(lambda src: self.nm_plan.refresh_ui_data())
        self.nm_crud.import_done.connect(lambda src: self.nm_plan.refresh_ui_data())

        # Estos de abajo ya estaban bien, los dejas igual:
        self.rgm_types.types_changed.connect(
            lambda src: self.rgm_plan.refresh_ui_data()
        )
        self.nm_types.types_changed.connect(lambda src: self.nm_plan.refresh_ui_data())

    def handle_logout(self):
        self.logout_signal.emit()
        self.close()
        
    def _on_excel_path_changed(self, source: str, new_path: str) -> None:
        if source == "RGM":
            self.rgm_excel = new_path
            self.rgm_crud.excel_file = new_path
            self.rgm_types.excel_file = new_path
        elif source == "Newmont":
            self.newmont_excel = new_path
            self.nm_crud.excel_file = new_path
            self.nm_types.excel_file = new_path
