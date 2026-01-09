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
from PyQt6.QtGui import QColor, QFont, QCursor, QIcon, QPixmap, QPainter, QPen
from datetime import datetime, date as pydate, timedelta
from PyQt6.QtWidgets import QStyledItemDelegate


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
        pen.setColor(QColor("#A9A9A9"))  # Darker gray for visibility on light backgrounds
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
    names = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
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

        # Autofill a la derecha
        self.apply_to_range_chk = QCheckBox(
            "Apply to all selected cells to the right"
        )
        
        from PyQt6.QtCore import QDate, QTime

        # --- Travel dates opcionales (igual concepto que en el formulario grande) ---
        self.travel_diff_chk = QCheckBox("Travel dates are different from work day")

        # Valores por defecto (puedes ajustarlos luego si quieres)
        today = QDate.currentDate()
        self.entry_date_edit = QDateEdit(today)
        self.entry_date_edit.setCalendarPopup(True)
        self.entry_date_edit.setDisplayFormat("dd/MM/yyyy")

        self.entry_time_edit = QTimeEdit(QTime(6, 0))
        self.entry_time_edit.setDisplayFormat("HH:mm")

        self.exit_date_edit = QDateEdit(today)
        self.exit_date_edit.setCalendarPopup(True)
        self.exit_date_edit.setDisplayFormat("dd/MM/yyyy")

        self.exit_time_edit = QTimeEdit(QTime(18, 0))
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

        # Ocultamos el bloque hasta que el usuario marque el check
        self._travel_container.setVisible(False)
        self.travel_diff_chk.toggled.connect(self._travel_container.setVisible)


        # Form layout
        form = QFormLayout()
        form.addRow("Status / Shift:", self.status_combo)
        form.addRow("Pick Up:", self.pickup_combo)
        form.addRow("Drop Off:", self.dropoff_combo)
        form.addRow("Remarks:", self.remark_edit)
        form.addRow("", self.travel_diff_chk)      # 👈 NUEVO
        form.addRow("", self._travel_container)    # 👈 NUEVO


        # Botones
        btn_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
        )
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)

        layout.addLayout(form)
        layout.addWidget(self.apply_to_range_chk)
        layout.addWidget(btn_box)

    def result_payload(self):
        from datetime import datetime

        selection = self.status_combo.currentData() or {}

        entry_dt = None
        exit_dt = None
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
            "apply_to_range": self.apply_to_range_chk.isChecked(),
            "entry_datetime": entry_dt,   # 👈 NUEVO
            "exit_datetime": exit_dt,     # 👈 NUEVO
        }




# -------------------------------------------------------------
# Widget: Plan Staff (Preview, Register, Reports)
# -------------------------------------------------------------
class PlanStaffWidget(QWidget):
    # Emitted after saving a change so the Rotation History tab can refresh
    rotation_changed = pyqtSignal()

    def __init__(self, source: str, excel_file: str, logged_username: str):
        super().__init__()
        self.source = source  # "RGM" | "Newmont"
        self.excel_file = excel_file
        self.logged_username = logged_username or "Unknown"
        self._is_internal_update = False
        self._last_excel_mtime = None
        self._missing_prompt_shown = False
        self._custom_shift_map = db.get_shift_type_map(self.source)
    



        # For REQ-001 tracking
        self._loading_preview = False
        self._cell_original_values = {}  # (row, col) -> original text
        self._row_identities = []  # index -> {"name":..., "badge":...}
        self._date_col_dates = []  # schedule_table column index -> pydate
        self._warn_highlight_keys = set()  # {"<badge>|YYYY-MM-DD", ...}
        self._bulk_editing = False  # para detectar relleno masivo por arrastre


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
            "1. Register Employee Schedule (DB is SSoT)", collapsed=False
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
        self.refresh_ui_data()

        # --- File monitor: detect moved/deleted/renamed file ---
        self.file_watch_timer = QTimer(self)
        self.file_watch_timer.setInterval(2000)  # 2s
        self.file_watch_timer.timeout.connect(self.check_excel_health)
        self.file_watch_timer.start()
        self.check_excel_health()

        # Track current responsive columns for the register grid
        self._current_form_cols = 3
        self._rebuild_registration_grid(self._current_form_cols)

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

        # Controls
        self.user_selector_combo = QComboBox()
        self.user_selector_combo.currentIndexChanged.connect(self.autofill_user_data)

        self.role_display = QLineEdit()
        self.role_display.setReadOnly(True)

        self.badge_display = QLineEdit()
        self.badge_display.setReadOnly(True)

        self.status_selector = QComboBox()

        self.start_date_edit = QDateEdit(QDate.currentDate())
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDisplayFormat("dd/MM/yyyy")

        self.end_date_edit = QDateEdit(QDate.currentDate().addDays(7))
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setDisplayFormat("dd/MM/yyyy")

        # NEW: location dropdowns
        self.pickup_combo = QComboBox()
        self.dropoff_combo = QComboBox()
        self.remark_edit = QLineEdit()
        
        self.apply_to_range_chk = QCheckBox("Apply to all selected cells to the right")

        
        # NEW: remarks input
        self.remarks_input = QLineEdit()
        self.remarks_input.setPlaceholderText("Optional: add a note for this period...")

        # Restored blue Save button (center action bar)
        self.save_button = QPushButton("💾 Save Changes to DB Excel")
        self.save_button.clicked.connect(self.save_plan_changes)
        self.save_button.setProperty("variant", "primary")
        
        # NEW: Separate travel dates controls
        self.travel_dates_check = QCheckBox("Travel dates are different from work period")
        self.travel_dates_check.toggled.connect(self._toggle_travel_dates_visibility)

        self.entry_date_edit = QDateEdit(QDate.currentDate())
        self.entry_date_edit.setCalendarPopup(True)
        self.entry_date_edit.setDisplayFormat("dd/MM/yyyy")
        
        # MODIFIED: Add time input for entry
        self.entry_time_edit = QTimeEdit(QTime(6, 0))
        self.entry_time_edit.setDisplayFormat("HH:mm")

        self.exit_date_edit = QDateEdit(QDate.currentDate().addDays(14))
        self.exit_date_edit.setCalendarPopup(True)
        self.exit_date_edit.setDisplayFormat("dd/MM/yyyy")
        
        # MODIFIED: Add time input for exit
        self.exit_time_edit = QTimeEdit(QTime(18, 0))
        self.exit_time_edit.setDisplayFormat("HH:mm")

        self._travel_dates_container = QWidget()
        travel_layout = QHBoxLayout(self._travel_dates_container)
        travel_layout.setContentsMargins(0, 0, 0, 0)
        travel_layout.addWidget(QLabel("Entry Date:"))
        travel_layout.addWidget(self.entry_date_edit)
        travel_layout.addWidget(QLabel("Time:"))
        travel_layout.addWidget(self.entry_time_edit)
        travel_layout.addSpacing(20)
        travel_layout.addWidget(QLabel("Exit Date:"))
        travel_layout.addWidget(self.exit_date_edit)
        travel_layout.addWidget(QLabel("Time:"))
        travel_layout.addWidget(self.exit_time_edit)
        travel_layout.addStretch()
        self._travel_dates_container.setVisible(False)


        # Field containers (label on top)
        def field(title: str, w: QWidget) -> QWidget:
            cont = QWidget()
            v = QVBoxLayout(cont)
            v.setContentsMargins(0, 0, 0, 0)
            v.setSpacing(4)
            lbl = QLabel(title)
            v.addWidget(lbl)
            v.addWidget(w)
            return cont

        self._fields = [
            field("Select Employee", self.user_selector_combo),
            field("Role / Department", self.role_display),
            field("Badge (ID)", self.badge_display),
            field("Status / Shift", self.status_selector),
            field("Period Start Date", self.start_date_edit),
            field("Period End Date", self.end_date_edit),
            field("Pick Up Location", self.pickup_combo),
            field("Drop Off Location", self.dropoff_combo),
            field("Remarks", self.remarks_input),
        ]

        # Base grid (we will re-pack it responsively in _rebuild_registration_grid)
        self._register_grid = QGridLayout()
        self._register_grid.setContentsMargins(8, 6, 8, 6)
        self._register_grid.setHorizontalSpacing(12)
        self._register_grid.setVerticalSpacing(10)
        self.main_form_layout.addLayout(self._register_grid)

        # Add new travel date controls after the grid
        self.main_form_layout.addWidget(self.travel_dates_check)
        self.main_form_layout.addWidget(self._travel_dates_container)

        # Save bar (full width)
        self._save_bar = QHBoxLayout()
        self._save_bar.addStretch()
        self._save_bar.addWidget(self.save_button)
        self._save_bar.addStretch()
        self.main_form_layout.addLayout(self._save_bar)

        # Load options (base + custom)
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
                    },
                )
        self.status_selector.setCurrentIndex(0)
        self.status_selector.blockSignals(False)
        
    def _status_options_for_dialog(self):
        """Devuelve la lista de opciones (icono, texto, payload) para los diálogos/combos."""
        options = []

        # Opción neutra (no marcar)
        options.append(
            (QIcon(), "— Do Not Mark Days —", {"kind": "none"})
        )

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

    def load_schedule_data(self):
        df = excel.get_schedule_preview(self.excel_file)
        self._loading_preview = True
        self._cell_original_values.clear()
        self._row_identities.clear()
        self._date_col_dates.clear()

        if df.empty:
            self.frozen_table.clear()
            self.schedule_table.clear()
            self.frozen_table.setRowCount(0)
            self.schedule_table.setRowCount(0)
            self._loading_preview = False
            return
        """
        Loads data from Excel + DB into the table widget.
        Uses a helper thread to avoid freezing UI (conceptually).
        """
        # --- NUEVO: HORIZONTE MÓVIL AUTOMÁTICO ---
        # Antes de cargar nada, aseguramos que el Excel tenga las columnas futuras
        # según la fecha de hoy.
        try:
            ok_horizon, msg_horizon = excel.ensure_rolling_horizon_columns(self.excel_file)
            if ok_horizon:
                print(f"[AUTO-HORIZON] {msg_horizon}") # Log de consola
                # Opcional: Mostrar mensaje no intrusivo en barra de estado si existiera
            elif "Error" in msg_horizon:
                print(f"[AUTO-HORIZON ERROR] {msg_horizon}")
        except Exception as e:
            print(f"[AUTO-HORIZON CRITICAL] Failed: {e}")
        # -------------------------------------------

        # color mapping for custom codes
        custom_map = db.get_shift_type_map(self.source)

        # -----------------------------------------------------------
        # MODIFICACIÓN: Ordenar columnas cronológicamente antes de procesar
        # -----------------------------------------------------------
        
        # 1. Identificar columnas fijas y columnas de fecha
        all_cols = list(df.columns)
        actual_frozen_count = min(len(all_cols), FROZEN_COLUMN_COUNT)
        
        frozen_part = all_cols[:actual_frozen_count]
        date_part = all_cols[actual_frozen_count:]

        # 2. Crear una lista de tuplas (ObjetoFecha, NombreColumnaOriginal)
        date_mapping = []
        for c in date_part:
            d_obj = None
            # Intentar extraer objeto fecha real
            if hasattr(c, "to_pydatetime"):
                d_obj = c.to_pydatetime().date()
            elif isinstance(c, datetime):
                d_obj = c.date()
            elif isinstance(c, pydate):
                d_obj = c
            
            # Si logramos obtener una fecha, la guardamos para ordenar
            if d_obj:
                date_mapping.append((d_obj, c))
            else:
                # Si no es fecha (raro), lo mandamos al final
                date_mapping.append((pydate.max, c))

        # 3. Ordenar la lista basándonos en la FECHA real (no texto)
        date_mapping.sort(key=lambda x: x[0])

        # 4. Reconstruir el DataFrame con el nuevo orden
        # Esto asegura que los DATOS de las celdas coincidan con los encabezados
        sorted_date_cols = [x[1] for x in date_mapping]
        new_column_order = frozen_part + sorted_date_cols
        
        df = df[new_column_order]





        # Prepare headers
        cols = list(df.columns)
        # Identify date columns (right side)
        date_cols = []
        for c in cols:
            # pandas may give Timestamp-like objects; keep them as date
            if hasattr(c, "to_pydatetime"):
                date_cols.append(c.to_pydatetime().date())
            elif isinstance(c, datetime):
                date_cols.append(c.date())
            else:
                # not a date header
                pass
       
        # Frozen
        actual_frozen_count = min(df.shape[1], FROZEN_COLUMN_COUNT)
        frozen_headers = [str(c) for c in cols[:actual_frozen_count]]
        
        # Schedule (date) headers -> one line with date + weekday (abbrev)
        schedule_headers = []
        for d in date_cols:
            schedule_headers.append(f"{d.isoformat()}")
        self._date_col_dates = list(date_cols)  # keep exact order
        

        # Build tables
        self.frozen_table.setRowCount(df.shape[0])
        self.frozen_table.setColumnCount(actual_frozen_count)
        self.frozen_table.setHorizontalHeaderLabels(frozen_headers)

        self.schedule_table.setRowCount(df.shape[0])
        self.schedule_table.setColumnCount(len(schedule_headers))
        for idx, header_text in enumerate(schedule_headers):
            self.schedule_table.setHorizontalHeaderItem(
                idx, QTableWidgetItem(header_text)
            )

        # Load rows
        for i, row in df.iterrows():
           # identity
            badge_val = (
                row.get("BADGE")
                if hasattr(row, "get")
                else (row["BADGE"] if "BADGE" in df.columns else "")
            )
            name_val = (
                row.get("NAME")
                if hasattr(row, "get")
                else (row["NAME"] if "NAME" in df.columns else "")
            )
            
            # --- AGREGAR ESTO (INICIO) ---
            role_val = (
                row.get("ROLE")
                if hasattr(row, "get")
                else (row["ROLE"] if "ROLE" in df.columns else "")
            )
            # --- AGREGAR ESTO (FIN) ---

            self._row_identities.append(
                {
                    "badge": str(badge_val) if badge_val is not None else "",
                    "name": str(name_val) if name_val is not None else "",
                    "role": str(role_val) if role_val is not None else "", # <--- AGREGAR ESTA LÍNEA
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

                    # center text in day cells
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

                    # Base colors
                    if "ON NS" in val_str or "NIGHT" in val_str:
                        item.setBackground(QColor("#FFFF99"))
                    elif val_str == "ON" or "DAY" in val_str or val_str.isdigit():
                        item.setBackground(QColor("#C6EFCE"))
                    elif val_str in ("OFF", "BREAK", "KO", "LEAVE"):
                        item.setBackground(QColor("#FFC7CE"))
                    else:
                        # custom code?
                        if val_str in custom_map and custom_map[val_str].get(
                            "color_hex"
                        ):
                            item.setBackground(QColor(custom_map[val_str]["color_hex"]))

                    self.schedule_table.setItem(i, col_index, item)
                    # Track original value for REQ-001
                    self._cell_original_values[(i, col_index)] = val_str

                    # Re-apply warning highlight if previously set
                    key = self._warn_key_for(i, col_index)
                    if key in self._warn_highlight_keys:
                        item.setBackground(QColor(WARN_BG_HEX))

        self.frozen_table.resizeColumnsToContents()
        self.schedule_table.resizeColumnsToContents()
        self._update_frozen_width()
        self._loading_preview = False

        # REQ-002: focus today's date
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

    def _apply_fill_from_anchor(self):
        """
        Excel-style fill: copy the current cell to the other selected cells.
        CORRECCIÓN: Ahora calcula dinámicamente el entry_date y exit_date para el
        nuevo rango completo, asegurando que se registren las horas (times)
        consistentemente como en el Metodo 1 (Formulario).
        """
        current_item = self.schedule_table.currentItem()
        if not current_item:
            return

        base_text = (current_item.text() or "").strip().upper()
        if not base_text:
            return

        selected_indexes = self.schedule_table.selectedIndexes()
        if len(selected_indexes) <= 1:
            return

        anchor_row = current_item.row()
        anchor_col = current_item.column()

        # Solo celdas en la misma fila
        target_columns = sorted(
            {
                idx.column()
                for idx in selected_indexes
                if idx.row() == anchor_row and idx.column() != anchor_col
            }
        )
        if not target_columns:
            return

        # --- Parte 1: Actualizar visualmente la tabla (UI) ---
        self._bulk_editing = True
        try:
            for col in target_columns:
                item = self.schedule_table.item(anchor_row, col)
                if item is None:
                    item = QTableWidgetItem()
                    self.schedule_table.setItem(anchor_row, col, item)
                item.setText(base_text)
                self._apply_base_background(item, base_text)
        finally:
            self._bulk_editing = False

        # --- Parte 2: Lógica de Persistencia (DB + Excel) ---
        
        # Validar identidad
        if not (0 <= anchor_row < len(self._row_identities)):
            return
        identity = self._row_identities[anchor_row]
        badge = (identity.get("badge") or "").strip()
        username = (identity.get("name") or "").strip() or badge
        role = (identity.get("role") or "").strip()

        if not badge:
            return

        # Definir el Rango de Fechas Afectado (Start -> End)
        all_cols = sorted(set(target_columns + [anchor_col]))
        first_col, last_col = all_cols[0], all_cols[-1]
        
        if not (0 <= first_col < len(self._date_col_dates) and 0 <= last_col < len(self._date_col_dates)):
            return

        start_date = self._date_col_dates[first_col]  # Nueva fecha inicio del bloque
        end_date = self._date_col_dates[last_col]     # Nueva fecha fin del bloque (donde soltaste el click)
        anchor_date = self._date_col_dates[anchor_col]

        # --- Obtener info BASE del día ancla ---
        schedule_map = db.get_schedule_map_for_range(badge, anchor_date, anchor_date, self.source)
        day_info = schedule_map.get(anchor_date.isoformat()) or {}

        schedule_status = (day_info.get("status") or base_text).strip().upper()
        shift_type = day_info.get("shift_type")
        in_time = day_info.get("in_time")
        out_time = day_info.get("out_time")
        remark = day_info.get("remark")
        
        # Recuperar ubicación del ancla
        pickup, dropoff = db.get_user_location_for_date(badge, anchor_date)

        # Si faltan datos en BD, intentar inferir del mapa de tipos
        if not shift_type or not in_time or not out_time:
            if schedule_status in self._custom_shift_map:
                info = self._custom_shift_map[schedule_status]
                shift_type = shift_type or info.get("name")
                in_time = in_time or info.get("in_time")
                out_time = out_time or info.get("out_time")

        # --- CRÍTICO: Calcular Entry/Exit Date CORRECTOS con Hora ---
        # 1. Buscar si el ancla ya tenía una operación con horas definidas
        operations = db.get_operations_filtered(text=badge, d_from=anchor_date, d_to=anchor_date)
        operations = [op for op in operations if op.get("badge") == badge]
        op_info = operations[0] if operations else None

       # 2. Definir horas por defecto según Proyecto (Source)
        # Regla Newmont: 06:00 - 12:00 (Día) | 12:00 - 06:00 (Noche)
        # Regla RGM:     06:00 - 18:00 (Día) | 18:00 - 06:00 (Noche)
        
        # Valores base (RGM Dia)
        def_in_str = "06:00"
        def_out_str = "18:00"

        # Ajuste dinámico según Source y Status
        if self.source == "Newmont":
            if schedule_status == "ON NS": # Noche Newmont
                def_in_str = "12:00"
                def_out_str = "06:00"
            elif schedule_status == "ON": # Día Newmont
                def_in_str = "06:00"
                def_out_str = "12:00"
        else: # RGM
            if schedule_status == "ON NS": # Noche RGM
                def_in_str = "18:00"
                def_out_str = "06:00"
            # else: Dia RGM se queda con el base 06:00-18:00

        default_entry_time = datetime.strptime(def_in_str, "%H:%M").time()
        default_exit_time = datetime.strptime(def_out_str, "%H:%M").time()
        
        final_entry_time = default_entry_time
        final_exit_time = default_exit_time

        # 3. Si el ancla tenía horas específicas, las heredamos
        if op_info:
            entry_dt_str = op_info.get("entry_date")
            exit_dt_str = op_info.get("exit_date")
            if entry_dt_str and " " in entry_dt_str:
                try:
                    final_entry_time = datetime.strptime(entry_dt_str, "%Y-%m-%d %H:%M").time()
                except ValueError:
                    pass
            if exit_dt_str and " " in exit_dt_str:
                try:
                    final_exit_time = datetime.strptime(exit_dt_str, "%Y-%m-%d %H:%M").time()
                except ValueError:
                    pass
        
        # 4. Construir los nuevos timestamps usando las FECHAS del nuevo rango + las HORAS detectadas
        # Esto asegura que el último día (donde soltaste el scroll) tenga la fecha correcta y la hora de salida.
        new_entry_datetime = datetime.combine(start_date, final_entry_time)
        new_exit_datetime = datetime.combine(end_date, final_exit_time)

        # --- Guardar en BD (SSoT) ---
        
        # 1. Registrar operación con timestamps completos (SOLUCIÓN DEL PROBLEMA)
        db.add_operation(
            username=username,
            role=role,
            badge=badge,
            start_date=start_date,
            end_date=end_date,
            created_by=self.logged_username,
            entry_date=new_entry_datetime, # Ahora tiene la fecha correcta de inicio + hora
            exit_date=new_exit_datetime    # Ahora tiene la fecha correcta de fin + hora
        )

        # 2. Actualizar schedules individuales
        db.upsert_schedule_range(
            badge, start_date, end_date, schedule_status, 
            shift_type, self.source, in_time, out_time, remark
        )

        # 3. Ubicaciones
        if pickup or dropoff:
            db.assign_user_location_range(badge, start_date, end_date, pickup, dropoff)

        # --- Guardar en Excel ---
        success, message = excel.update_plan_staff_excel(
            self.excel_file, username, role, badge,
            schedule_status, shift_type, start_date, end_date,
            self.source, in_time, out_time
        )

        # --- Auditoría ---
        prev_map = db.get_schedule_map_for_range(badge, start_date, end_date, self.source) # solo referencia
        new_map = db.get_schedule_map_for_range(badge, start_date, end_date, self.source)
        
        db.log_event(
            self.logged_username, self.source, "SHIFT_MODIFICATION",
            f"[DragFill] {username} ({badge}) {start_date}..{end_date} entry={new_entry_datetime} exit={new_exit_datetime}"
        )

        if not success:
            QMessageBox.warning(self, "Excel Update", "Saved to DB, but Excel issue:\n" + message)

        self.rotation_changed.emit()


    # ---------- REQ-001: inline OFF→ON/ON NS guard ----------
    def _on_schedule_cell_changed(self, item: QTableWidgetItem):
        if self._loading_preview or self._bulk_editing:
            return

        r = item.row()
        c = item.column()
        new_text = (item.text() or "").strip().upper()
        old_text = (self._cell_original_values.get((r, c), "") or "").strip().upper()

        # ---------------- REQ-001: OFF/blank -> ON / ON NS ----------------
        if old_text in ("OFF", "") and new_text in ("ON", "ON NS"):
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Warning)
            box.setWindowTitle("Confirm Change")
            box.setText("The employee is on a day off. Do you want to set it to ON?")
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
        current_status_text = (day_info.get("status") or new_text or "").upper()

        pickup_init, dropoff_init = db.get_user_location_for_date(badge, base_date)
        initial = {
            "status_text": current_status_text,
            "pickup": pickup_init,
            "dropoff": dropoff_init,
            "remark": day_info.get("remark") or "",
        }

        locations = [loc["pickup_location"] for loc in db.get_locations(self.source)]

        editor = DayScheduleEditor(
            self,
            status_options=self._status_options_for_dialog(),
            locations=locations,
            initial=initial,
        )
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

        # Validación simple: si los dio y están invertidos, avisar y cancelar
        if entry_datetime and exit_datetime and entry_datetime > exit_datetime:
            QMessageBox.warning(
                self,
                "Date Error",
                "Entry date/time cannot be after Exit date/time.",
            )
            # revertimos visualmente al valor anterior
            with QSignalBlocker(self.schedule_table):
                item.setText(old_text)
            self._apply_base_background(item, old_text)
            return


        # ---------------- Interpretar selección (igual que save_plan_changes) ----------------
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

        # Si eligió "Do Not Mark Days" limpiamos celdas y no tocamos BD/Excel
        if schedule_status is None:
            with QSignalBlocker(self.schedule_table):
                for cc in ([c] if not apply_to_range else col_range):
                    it = self.schedule_table.item(r, cc)
                    if it is None:
                        it = QTableWidgetItem("")
                        self.schedule_table.setItem(r, cc, it)
                    it.setText("")
                    self._apply_base_background(it, "")
            return

       
     
        # ---------------- Aplicar cambios (BLOQUE OPTIMIZADO) ----------------
        from datetime import datetime, time as dtime

        # 1. Definir el rango total de fechas afectado
        cols_sorted = sorted(col_range if apply_to_range else [c])
        first_col_idx = cols_sorted[0]
        last_col_idx = cols_sorted[-1]

        start_date_block = self._date_col_dates[first_col_idx]
        end_date_block = self._date_col_dates[last_col_idx]

        # 2. Definir Horas (Entry/Exit) para el bloque
        if entry_datetime and exit_datetime:
            entry_dt_for_save = entry_datetime
            exit_dt_for_save = exit_datetime
        else:
            # Si no hay horas específicas del diálogo, usar las del tipo de turno o defaults
            # --- CORRECCIÓN INICIO: Lógica dinámica por Source (Newmont/RGM) ---
            from datetime import time as dtime
            
            # Valores por defecto base
            def_in = dtime(6, 0)
            def_out = dtime(18, 0) # Default RGM

            if self.source == "Newmont":
                if schedule_status == "ON NS": 
                    def_in = dtime(12, 0)
                    def_out = dtime(6, 0)
                elif schedule_status == "ON": 
                    def_in = dtime(6, 0)
                    def_out = dtime(12, 0)
            else: # RGM
                if schedule_status == "ON NS": 
                    def_in = dtime(18, 0)
                    def_out = dtime(6, 0)

            entry_time_obj = _parse_hhmm_to_time(in_time, def_in) 
            exit_time_obj = _parse_hhmm_to_time(out_time, def_out) 
            # --- CORRECCIÓN FIN ---

            # Entry es el PRIMER día a la hora de entrada
            entry_dt_for_save = datetime.combine(start_date_block, entry_time_obj)
            # Exit es el ÚLTIMO día a la hora de salida
            exit_dt_for_save = datetime.combine(end_date_block, exit_time_obj)
          

        # 3. Actualizar la UI visualmente (CRÍTICO: Bloquear señales para evitar recursividad)
        # Esto da feedback inmediato al usuario sin esperar al IO del disco
        with QSignalBlocker(self.schedule_table):
            for col_idx in cols_sorted:
                item_ui = self.schedule_table.item(r, col_idx)
                if not item_ui:
                    item_ui = QTableWidgetItem()
                    self.schedule_table.setItem(r, col_idx, item_ui)
                
                # Texto a mostrar (si es None o Do Not Mark, limpiar)
                text_to_show = schedule_status if schedule_status else ""
                item_ui.setText(text_to_show)
                self._apply_base_background(item_ui, text_to_show)
                
                # Actualizar valor original para tracking de REQ-001
                self._cell_original_values[(r, col_idx)] = text_to_show
# 4. Actualizar Base de Datos (SSoT)
        try:
            # A. Registrar la Operación (Historial)
            db.add_operation(
                username=username,
                role=role,
                badge=badge,
                start_date=start_date_block,
                end_date=end_date_block,
                created_by=self.logged_username,
                entry_date=entry_dt_for_save,
                exit_date=exit_dt_for_save
            )

            # B. Guardar el Schedule (CORRECCIÓN: Se agrega 'remark')
            db.upsert_schedule_range(
                badge,
                start_date_block,
                end_date_block,
                schedule_status,
                shift_type,
                self.source,
                in_time,        # hora entrada (str)
                out_time,       # hora salida (str)
                remark          # <--- CORRECCIÓN CRÍTICA: Antes no se enviaba
            )

            # C. Guardar Ubicaciones (CORRECCIÓN: Antes faltaba este bloque en el diálogo)
            if pickup or dropoff:
                db.assign_user_location_range(
                    badge, 
                    start_date_block, 
                    end_date_block, 
                    pickup, 
                    dropoff
                )
                # Log opcional para auditoría
                db.log_event(
                    self.logged_username,
                    self.source,
                    "LOCATION_ASSIGN_INLINE",
                    f"{badge} {start_date_block}..{end_date_block} PU={pickup} DO={dropoff}"
                )

            # 5. Actualizar Excel (SSoT secundario)
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

            # ... resto del código (QTimer, logs, etc) sigue igual ...
            QTimer.singleShot(2000, lambda: setattr(self, '_is_internal_update', False))

            if not success:
                self._is_internal_update = False 
                raise Exception(message)

            db.log_event(
                self.logged_username,
                self.source,
                "SHIFT_MODIFICATION_INLINE",
                f"Updated range {start_date_block} to {end_date_block} for {badge}. Remark: {remark}"
            )

        except Exception as e:
            self._is_internal_update = False 
            QMessageBox.critical(self, "Save Error", f"Error saving data: {str(e)}")
            self.refresh_ui_data()
            return

        # 6. Finalización
        # NO llamamos a refresh_ui_data() aquí para evitar el parpadeo. 
        # Ya confiamos en que la UI, DB y Excel están sincronizados.
        self.check_excel_health() # Solo verificamos que el archivo siga existiendo
        self.rotation_changed.emit()

    # ---------- MODIFIED: hover card logic ----------
    def _show_shift_tooltip(self, row: int, col: int):
        # Ensure row/col are valid
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

        # Get the global position of the cell for anchoring the card
        cell_rect_viewport = self.schedule_table.visualItemRect(item)
        global_top_left = self.schedule_table.viewport().mapToGlobal(
            cell_rect_viewport.topLeft()
        )
        global_cell_rect = QRect(global_top_left, cell_rect_viewport.size())

        # --- Get identifiers from internal state ---
        identity = self._row_identities[row]
        badge = identity.get("badge")
        hover_date = self._date_col_dates[col]

        if not badge or not hover_date:
            self._shift_info_card.hide()
            return

        # --- Fetch live data from DB (SSoT) ---
        schedule_data = db.get_schedule_map_for_range(
            badge, hover_date, hover_date, self.source
        )
        day_info = schedule_data.get(hover_date.isoformat())
        pickup, dropoff = db.get_user_location_for_date(badge, hover_date)

        # NEW: Fetch the operation record to get the correct travel dates
        operations = db.get_operations_filtered(
            text=badge, d_from=hover_date, d_to=hover_date
        )
        # Filter for exact badge match and take the latest one if multiple overlap
        operations = [op for op in operations if op.get("badge") == badge]
        operation_info = operations[0] if operations else None


        if not day_info:
            final_html = "<p style='margin:0;'>Information not available</p>"
            self._shift_info_card.show_info(global_cell_rect, final_html)
            return

        # --- Format data for the card (New compact format) ---
        status_code = (day_info.get("status") or "N/A").upper()
        remark = day_info.get("remark")


        # --- BLOQUE A MODIFICAR en _show_shift_tooltip ---

        # Determine Shift Title and Times
        shift_title = status_code
        in_time = day_info.get("in_time")
        out_time = day_info.get("out_time")

        # LOGICA NUEVA: Inyectar horarios por defecto para ON / ON NS si no vienen de BD
        if status_code == "ON":
            shift_title = "ON (Day Shift)"
            if not in_time: 
                in_time = "06:00"
            if not out_time:
                # Regla de negocio: Newmont sale a las 12:00, RGM a las 18:00
                out_time = "12:00" if self.source == "Newmont" else "18:00"

        elif status_code == "ON NS":
            shift_title = "ON NS (Night Shift)"
            if not in_time:
                # Regla de negocio: Newmont entra a las 12:00, RGM a las 18:00
                in_time = "12:00" if self.source == "Newmont" else "18:00"
            if not out_time:
                out_time = "06:00"

        elif status_code in self._custom_shift_map:
            custom_info = self._custom_shift_map[status_code]
            shift_title = custom_info.get("name", status_code)
            # Si faltan horas en el día específico, usar las del tipo de turno general
            if not in_time: in_time = custom_info.get("in_time")
            if not out_time: out_time = custom_info.get("out_time")
            
        elif status_code == "OFF":
            shift_title = "OFF"

        # Build compact HTML content
        title_style = "style='margin: 0 0 2px 0; font-size: 14px; color: #FFFFFF; font-weight: 600;'"
        schedule_style = "style='margin: 0; font-size: 13px; color: #FFFFFF;'"

        content_lines = [f"<p {title_style}>{shift_title}</p>"]

        if in_time and out_time:
            content_lines.append(f"<p {schedule_style}>⏰ {in_time} – {out_time}</p>")

        # MODIFIED: Add travel dates and exact times from the operation record
        if operation_info:
            entry_dt_str = operation_info.get('entry_date', '')
            exit_dt_str = operation_info.get('exit_date', '')
            
            # Check for the presence of time in the string
            if entry_dt_str and exit_dt_str and ' ' in entry_dt_str:
                try:
                    # Parse 'YYYY-MM-DD HH:MM'
                    entry_dt = datetime.strptime(entry_dt_str, '%Y-%m-%d %H:%M')
                    exit_dt = datetime.strptime(exit_dt_str, '%Y-%m-%d %H:%M')

                    entry_date_display = entry_dt.strftime('%Y-%m-%d')
                    entry_time_display = entry_dt.strftime('%H:%M')

                    exit_date_display = exit_dt.strftime('%Y-%m-%d')
                    exit_time_display = exit_dt.strftime('%H:%M')

                    travel_html = "<div style='border-top: 1px solid #4B5563; padding-top: 6px; margin-top: 8px;'>"
                    travel_html += f"<p {schedule_style}>✈️ <b>Entry:</b> {entry_date_display} at {entry_time_display}</p>"
                    travel_html += f"<p {schedule_style}>✈️ <b>Exit:</b> {exit_date_display} at {exit_time_display}</p>"
                    travel_html += "</div>"
                    content_lines.append(travel_html)
                except ValueError:
                    # Fallback if parsing fails, should not happen with new save logic
                    pass


        # Logistic Notes (only if they exist)
        pickup_clean = _clean(pickup)
        dropoff_clean = _clean(dropoff)

        if pickup_clean or dropoff_clean:
            logistics_html = "<div style='border-top: 1px solid #4B5563; padding-top: 6px; margin-top: 8px;'>"
            logistics_html += f"<p {schedule_style}>📍 <b>Pick Up:</b> {pickup_clean or 'Not assigned'}</p>"
            logistics_html += f"<p {schedule_style}>📍 <b>Drop Off:</b> {dropoff_clean or 'Not assigned'}</p>"
            logistics_html += "</div>"
            content_lines.append(logistics_html)

        # NEW: Add remarks section
        if remark:
            remark_style = "style='margin: 0; font-size: 13px; color: #FFFFFF;'"
            # Soft yellow background for the remarks block
            remark_block_style = "style='border-top: 1px solid #1565C0; padding-top: 8px; margin-top: 8px;'"
            remark_html = f"<div {remark_block_style}>"
            remark_html += f"<p {remark_style}><b>Remark:</b> {remark}</p>"
            remark_html += "</div>"
            content_lines.append(remark_html)


        html_body = "".join(content_lines)
        final_html = f"<div style='line-height: 1.3;'>{html_body}</div>"

        # Pass the global cell rect to the show_info method to handle positioning
        self._shift_info_card.show_info(global_cell_rect, final_html)

    def load_users_to_selector(self):
        self.user_selector_combo.blockSignals(True)
        self.user_selector_combo.clear()
        self.users_for_selector = db.get_all_users(self.source)
        self.user_selector_combo.addItem("-- Select a user --")
        for user in self.users_for_selector:
            self.user_selector_combo.addItem(user["name"])
        self.user_selector_combo.setCurrentIndex(0)
        self.user_selector_combo.blockSignals(False)
        # clear dependent fields
        self.role_display.clear()
        self.badge_display.clear()

    def refresh_users_only(self):
        """Repopulate ONLY the users combo (for instant sync)."""
        self.load_users_to_selector()

    def autofill_user_data(self, index):
        if index > 0:
            user = self.users_for_selector[index - 1]
            self.role_display.setText(user["role"])
            self.badge_display.setText(user["badge"])
        else:
            self.role_display.clear()
            self.badge_display.clear()


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

            # Guardar operación de 1 día (entry/exit = día con horas)
            entry_datetime = datetime.combine(start_date, in_time_obj)
            exit_datetime = datetime.combine(end_date, out_time_obj)

      

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
            db.assign_user_location_range(
                badge, start_date, end_date, pickup, dropoff
            )
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




    # ---------- actions ----------
    def save_plan_changes(self):
        # 1. Limpiar estados de error visuales
        mark_error(self.user_selector_combo, False)
        mark_error(self.role_display, False)
        mark_error(self.start_date_edit, False)
        mark_error(self.end_date_edit, False)
        mark_error(self.entry_date_edit, False)
        mark_error(self.exit_date_edit, False)

        # 2. Leer datos básicos del formulario
        username = self.user_selector_combo.currentText()
        badge = self.badge_display.text()
        role = self.role_display.text()
        start_date = self.start_date_edit.date().toPyDate()
        end_date = self.end_date_edit.date().toPyDate()
        
        pickup = self.pickup_combo.currentData() or None
        dropoff = self.dropoff_combo.currentData() or None
        remark = self.remarks_input.text().strip() or None

        # 3. Validaciones básicas
        if not username or username == "-- Select a user --":
            mark_error(self.user_selector_combo, True)
            QMessageBox.warning(self, "Incomplete Data", "Please select an employee.")
            return

        if not role:
            mark_error(self.role_display, True)
            QMessageBox.warning(self, "Incomplete Data", "Please select a role/department.")
            return

        if start_date > end_date:
            mark_error(self.start_date_edit, True)
            mark_error(self.end_date_edit, True)
            QMessageBox.warning(self, "Date Error", "Start date cannot be after end date.")
            return

        # 4. Interpretar selección del turno (Status/Shift)
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
        # 5. CÁLCULO DE FECHAS Y HORAS DE VIAJE (CORRECCIÓN PUNTUAL)
        # ---------------------------------------------------------------------
        entry_datetime = None
        exit_datetime = None

        if self.travel_dates_check.isChecked():
            # OPCIÓN A: El usuario define fechas específicas manualmente
            entry_date = self.entry_date_edit.date().toPyDate()
            entry_time = self.entry_time_edit.time().toPyTime()
            entry_datetime = datetime.combine(entry_date, entry_time)
            
            exit_date = self.exit_date_edit.date().toPyDate()
            exit_time = self.exit_time_edit.time().toPyTime()
            exit_datetime = datetime.combine(exit_date, exit_time)
            
            # Validación lógica de viaje
            if entry_datetime > exit_datetime:
                mark_error(self.entry_date_edit, True)
                mark_error(self.exit_date_edit, True)
                QMessageBox.warning(self, "Date Error", "Entry date/time cannot be after Exit date/time.")
                return

            if entry_datetime.date() > start_date or exit_datetime.date() < end_date:
                reply = QMessageBox.question(self, "Confirm Dates", 
                    "The travel period does not fully encompass the work period. Continue?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
                if reply == QMessageBox.StandardButton.No:
                    return

        elif schedule_status is not None:
            # OPCIÓN B: Automático (Check desmarcado) -> Calcular Defaults
            # Se usa la fecha de inicio del periodo para Entry y fin para Exit.
            # Se inyectan las horas según el tipo de turno o reglas de negocio.
            
            # Hora por defecto base (06:00 / 18:00)
            t_in = datetime.strptime("06:00", "%H:%M").time()
            t_out = datetime.strptime("18:00", "%H:%M").time()

            # Lógica de Horas
            if in_time_raw and out_time_raw:
                # Si el turno (custom) ya trae horas definidas en DB
                try: 
                    t_in = datetime.strptime(str(in_time_raw)[:5], "%H:%M").time()
                    t_out = datetime.strptime(str(out_time_raw)[:5], "%H:%M").time()
                except: pass
            else:
                # Reglas Hardcoded (SSoT fallbacks)
                is_newmont = (self.source == "Newmont")
                if schedule_status == "ON": # Día
                    t_in = datetime.strptime("06:00", "%H:%M").time()
                    t_out = datetime.strptime("12:00" if is_newmont else "18:00", "%H:%M").time()
                elif schedule_status == "ON NS": # Noche
                    t_in = datetime.strptime("12:00" if is_newmont else "18:00", "%H:%M").time()
                    t_out = datetime.strptime("06:00", "%H:%M").time()

            # Crear los datetimes finales para guardar en BD
            entry_datetime = datetime.combine(start_date, t_in)
            exit_datetime = datetime.combine(end_date, t_out)
        
        # ---------------------------------------------------------------------
        # FIN DE LA CORRECCIÓN
        # ---------------------------------------------------------------------

        # 6. Detección de Conflictos (Overwrite check)
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

        # Para auditoría
        prev_map = db.get_schedule_map_for_range(
            badge, start_date, end_date, self.source
        )

        # 7. Guardar en BD (SSoT)
        if schedule_status is not None:
            # Aquí es donde se guardan los datetimes calculados (entry_datetime/exit_datetime)
            db.add_operation(
                username=username,
                role=role,
                badge=badge,
                start_date=start_date,
                end_date=end_date,
                created_by=self.logged_username,
                entry_date=entry_datetime, # Ahora siempre tendrá valor si hay turno
                exit_date=exit_datetime,   # Ahora siempre tendrá valor si hay turno
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
            )
        else:  # Limpiar rango ("Do Not Mark Days")
            db.clear_schedule_range(badge, start_date, end_date, self.source)

        # Guardar ubicación si aplica
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
        QTimer.singleShot(2000, lambda: setattr(self, '_is_internal_update', False))

        # 9. Auditoría y Finalización
        new_map = db.get_schedule_map_for_range(badge, start_date, end_date, self.source)
        db.log_event(
            self.logged_username,
            self.source,
            "SHIFT_MODIFICATION",
            f"{username} ({badge}) {start_date}..{end_date} prev={prev_map} new={new_map} remark={remark}; Excel={'OK' if success else 'ERR'}",
        )

        box = QMessageBox(self)
        box.setIcon(QMessageBox.Icon.Information if success else QMessageBox.Icon.Warning)
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
            QMessageBox.warning(self, "Date Range Error", "Start date cannot be after end date.")
            return

        try:
            # Default filename suggestion
            default_filename = f"Onsite_Stay_Report_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx"

            # Open file dialog to choose where to save
            filePath, _ = QFileDialog.getSaveFileName(
                self,
                "Save Onsite Stay Report",
                default_filename,
                "Excel Files (*.xlsx);;All Files (*)"
            )

            if not filePath:
                return  # User cancelled

            # Generate the report bytes
            report_bytes, msg = excel.generate_stay_period_report(
                self.excel_file,
                start_date,
                end_date
            )

            if report_bytes:
                with open(filePath, "wb") as f:
                    f.write(report_bytes)
                
                QMessageBox.information(self, "Report Generated", f"Successfully generated and saved report:\n{os.path.basename(filePath)}")
                db.log_event(
                    self.logged_username,
                    self.source,
                    "REPORT_STAY_GENERATED",
                    f"Generated for range {start_date} to {end_date}."
                )
            else:
                QMessageBox.critical(self, "Report Generation Failed", msg)

        except Exception as e:
            QMessageBox.critical(self, "Error", f"An unexpected error occurred while generating the report: {e}")
            db.log_event(
                self.logged_username,
                self.source,
                "ERROR_REPORT_STAY",
                f"Failed for range {start_date} to {end_date}: {e}"
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

    def refresh_ui_data(self):
        self.load_shift_type_options()
        self.load_schedule_data()
        self.load_users_to_selector()
        self.load_location_options()  # keep combos in sync with Location admin
        self.remarks_input.clear() # Clear remarks on refresh

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
                self.excel_file = new_path
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

        self.reset_filters() # To initialize and load data

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
        with QSignalBlocker(self.search_input), \
             QSignalBlocker(self.role_combo), \
             QSignalBlocker(self.date_from), \
             QSignalBlocker(self.date_to), \
             QSignalBlocker(self.active_today_check), \
             QSignalBlocker(self.sort_combo):
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
        self._filter_state['text'] = self.search_input.text()
        self._filter_state['role'] = self.role_combo.currentText() if self.role_combo.currentIndex() > 0 else None
        self._filter_state['sort'] = 'name_asc' if self.sort_combo.currentIndex() == 1 else 'start_date_desc'

        d_from = self.date_from.date().toPyDate()
        d_to = self.date_to.date().toPyDate()
        
        records = db.get_operations_filtered(
            text=self._filter_state['text'],
            role=self._filter_state['role'],
            d_from=d_from,
            d_to=d_to,
            sort_by=self._filter_state['sort'],
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
            self.table.setItem(row_idx, 5, QTableWidgetItem(record.get("created_by", "")))

        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)


# -------------------------------------------------------------
# Widget: Users CRUD (with Import from Excel)
# -------------------------------------------------------------
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

        # Left panel: user form
        form_layout = QGridLayout()
        form_layout.setContentsMargins(8, 8, 8, 8)

        self.crud_name_input = QLineEdit()
        self.crud_role_input = QLineEdit()
        self.crud_badge_input = QLineEdit()

        self.crud_save_button = QPushButton("💾 Save User")
        self.crud_save_button.clicked.connect(self.save_crud_user)
        self.crud_new_button = QPushButton("✨ New User")
        self.crud_new_button.clicked.connect(self.clear_crud_form)
        self.crud_delete_button = QPushButton("❌ Delete User")
        self.crud_delete_button.clicked.connect(self.delete_crud_user)

        self.import_button = QPushButton("📥 Import from Excel → DB (validated)")
        self.import_button.clicked.connect(self.import_users_from_excel)

        # Button variants
        self.crud_save_button.setProperty("variant", "primary")
        self.crud_new_button.setProperty("variant", "secondary")
        self.crud_delete_button.setProperty("danger", True)
        self.import_button.setProperty("variant", "secondary")

        form_layout.addWidget(QLabel("Full Name:"), 0, 0)
        form_layout.addWidget(self.crud_name_input, 0, 1)
        form_layout.addWidget(QLabel("Role/Department:"), 1, 0)
        form_layout.addWidget(self.crud_role_input, 1, 1)
        form_layout.addWidget(QLabel("Badge (ID):"), 2, 0)
        form_layout.addWidget(self.crud_badge_input, 2, 1)

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.crud_new_button)
        button_layout.addWidget(self.crud_save_button)
        form_layout.addLayout(button_layout, 3, 0, 1, 2)
        form_layout.addWidget(self.crud_delete_button, 4, 0, 1, 2)
        form_layout.addWidget(self.import_button, 5, 0, 1, 2)

        form_group = create_group_box("Manage User", form_layout)
        form_group.setFixedWidth(400)

        # Right panel: users table
        table_panel = QWidget()
        table_layout = QVBoxLayout(table_panel)
        
        # --- Filter Panel for Users ---
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
        self.users_table.itemClicked.connect(self.load_user_to_crud_form)
        table_layout.addWidget(self.users_table)

        table_group = create_group_box("Registered Users List", table_layout)
        layout.addWidget(form_group)
        layout.addWidget(table_group)

        self.refresh_ui_data()

    def _request_refresh(self):
        self._debounce_timer.start(DEBOUNCE_MS)

    def _populate_role_filter(self):
        self.user_role_combo.blockSignals(True)
        current_role = self.user_role_combo.currentText()
        self.user_role_combo.clear()
        self.user_role_combo.addItem("All Roles")
        all_users = db.get_all_users(self.source)
        roles = sorted(list(set(u['role'] for u in all_users if u.get('role'))))
        self.user_role_combo.addItems(roles)
        idx = self.user_role_combo.findText(current_role)
        if idx > 0:
            self.user_role_combo.setCurrentIndex(idx)
        self.user_role_combo.blockSignals(False)

    def reset_filters(self):
        with QSignalBlocker(self.user_search_input), \
             QSignalBlocker(self.user_role_combo), \
             QSignalBlocker(self.user_badge_prefix_input):
            self.user_search_input.clear()
            self.user_role_combo.setCurrentIndex(0)
            self.user_badge_prefix_input.clear()
        self.load_users_table()
        
    # Table & form CRUD
    def load_users_table(self):
        self._filter_state['text'] = self.user_search_input.text()
        self._filter_state['role'] = self.user_role_combo.currentText() if self.user_role_combo.currentIndex() > 0 else None
        self._filter_state['badge_prefix'] = self.user_badge_prefix_input.text()

        users = db.get_users_filtered(
            source=self.source,
            text=self._filter_state['text'],
            role=self._filter_state['role'],
            badge_prefix=self._filter_state['badge_prefix'],
            active_since=None # Not implemented in UI yet
        )
        
        headers = ["ID", "Name", "Role", "Badge"]
        self.users_table.setRowCount(len(users))
        self.users_table.setColumnCount(len(headers))
        self.users_table.setHorizontalHeaderLabels(headers)

        for row, user in enumerate(users):
            self.users_table.setItem(row, 0, QTableWidgetItem(str(user["id"])))
            self.users_table.setItem(row, 1, QTableWidgetItem(user["name"]))
            self.users_table.setItem(row, 2, QTableWidgetItem(user["role"]))
            self.users_table.setItem(row, 3, QTableWidgetItem(user["badge"]))
        self.users_table.setColumnHidden(0, True)  # hide ID column
        self.users_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch
        )

    def load_user_to_crud_form(self, item):
        row = item.row()
        self.current_user_id = int(self.users_table.item(row, 0).text())
        self.crud_name_input.setText(self.users_table.item(row, 1).text())
        self.crud_role_input.setText(self.users_table.item(row, 2).text())
        self.crud_badge_input.setText(self.users_table.item(row, 3).text())

    def clear_crud_form(self):
        self.current_user_id = None
        self.crud_name_input.clear()
        self.crud_role_input.clear()
        self.crud_badge_input.clear()
        self.users_table.clearSelection()

    def save_crud_user(self):
        name = self.crud_name_input.text().strip()
        role = self.crud_role_input.text().strip()
        badge = self.crud_badge_input.text().strip()

        if not name or not role or not badge:
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Warning)
            box.setWindowTitle("Incomplete Data")
            box.setText("All fields are required.")
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
            box.exec()
            return

        if self.current_user_id:
            success, message = db.update_user(
                self.current_user_id, name, role, badge, self.source
            )
        else:
            success, message = db.add_user(name, role, badge, self.source)

        # --- INICIO DEL CAMBIO ---
        if success:
            # 1. Sincronizar Excel inmediatamente (Escribir el usuario en el archivo físico)
            # Esto permite que el PlanStaffWidget lo lea al instante.
            try:
                excel.refresh_excel_from_db(self.excel_file, self.source)
            except Exception as e:
                print(f"Auto-refresh failed: {e}")

            # 2. Actualizar la UI local
            self._populate_role_filter()
            self.refresh_ui_data()
            
            # 3. Avisar a otras pestañas que hubo cambios
            self.users_changed.emit(self.source)
        # --- FIN DEL CAMBIO ---

        box = QMessageBox(self)
        box.setIcon(
            QMessageBox.Icon.Information if success else QMessageBox.Icon.Warning
        )
        box.setWindowTitle("Success" if success else "Error")
        box.setText(message)
        box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
        box.exec()

   # En main_window.py -> clase CrudWidget

    def delete_crud_user(self):
        if not self.current_user_id:
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Warning)
            box.setWindowTitle("No Selection")
            box.setText("Please select a user in the table to delete.")
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
            box.exec()
            return

        # 1. Capturar el BADGE antes de borrar (lo necesitamos para buscar en Excel)
        badge_to_remove = self.crud_badge_input.text().strip()

        # Configurar cuadro de confirmación correctamente
        confirm = QMessageBox(self)
        confirm.setIcon(QMessageBox.Icon.Question)
        confirm.setWindowTitle("Confirm User Deletion")
        confirm.setText(f"Are you sure you want to delete user '{self.crud_name_input.text()}'?")
        
        # DEFINICIÓN CORRECTA DE LOS BOTONES
        yes_btn = confirm.addButton("Yes", QMessageBox.ButtonRole.YesRole)
        confirm.addButton("No", QMessageBox.ButtonRole.NoRole)
        
        confirm.exec()

        if confirm.clickedButton() == yes_btn:
            # 2. Borrar de BD (SSoT)
            success_db, message_db = db.delete_user(self.current_user_id)
            
            final_msg = message_db
            
            if success_db:
                # 3. Si se borró en BD, borrar también del Excel
                success_excel, msg_excel = excel.remove_user_from_excel(self.excel_file, badge_to_remove)
                
                if success_excel:
                    final_msg += f"\n\nAlso removed from Excel: {badge_to_remove}"
                    # Log de auditoría
                    db.log_event(
                        self.logged_username, 
                        self.source, 
                        "USER_DELETE", 
                        f"Deleted {badge_to_remove} from DB and Excel."
                    )
                else:
                    final_msg += f"\n\nWarning: Could not remove from Excel ({msg_excel})"

                # Emitimos señal para actualizar la UI
                self._populate_role_filter()
                self.refresh_ui_data()
                self.users_changed.emit(self.source)

            box = QMessageBox(self)
            box.setIcon(
                QMessageBox.Icon.Information if success_db else QMessageBox.Icon.Warning
            )
            box.setWindowTitle("Deletion Result")
            box.setText(final_msg)
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
            box.exec()

    def import_users_from_excel(self):
        """
        FR-02: Import users and day-by-day schedules from Excel to DB.
        With strict structure validation first.
        """
        try:
            inserted, skipped, upserts = excel.import_excel_to_db(
                self.excel_file, self.source
            )

            # Audit log (FR-04)
            db.log_event(
                self.logged_username,
                self.source,
                "DATA_IMPORT",
                f"users_inserted={inserted}; users_skipped={skipped}; schedule_upserts={upserts}",
            )

            # Message
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Information)
            box.setWindowTitle("Import Complete")
            box.setText(
                f"Imported {inserted} new users.\n"
                f"Skipped {skipped} users that already existed.\n"
                f"Upserted {upserts} schedule day-entries."
            )
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
            box.exec()

            # Refresh users table immediately
            self.refresh_ui_data()
            # Signals to refresh "Select Employee" in Plan Staff
            self.users_changed.emit(self.source)
            self.import_done.emit(self.source)
        except ValueError as ve:
            # Structure error -> DO NOT save anything
            db.log_event(
                self.logged_username,
                self.source,
                "DATA_IMPORT",
                f"ERROR: {str(ve).replace(chr(10),' | ')}",
            )
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Critical)
            box.setWindowTitle("Invalid Excel")
            box.setText(str(ve))
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
            box.exec()

    def refresh_ui_data(self):
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
        form_layout.addWidget(QLabel("IN time (HH:MM):"), 3, 0)
        form_layout.addWidget(self.in_time_edit, 3, 1)
        form_layout.addWidget(QLabel("OUT time (HH:MM):"), 4, 0)
        form_layout.addWidget(self.out_time_edit, 4, 1)
        actions = QHBoxLayout()
        actions.addWidget(self.new_btn)
        actions.addWidget(self.save_btn)
        actions.addWidget(self.delete_btn)
        form_layout.addLayout(actions, 5, 0, 1, 2)

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

        self.reset_filters()

    def _request_refresh(self):
        self._debounce_timer.start(DEBOUNCE_MS)

    def reset_filters(self):
        with QSignalBlocker(self.st_search_input), \
             QSignalBlocker(self.st_time_from), \
             QSignalBlocker(self.st_time_to), \
             QSignalBlocker(self.st_usage_combo):
            self.st_search_input.clear()
            self.st_time_from.setTime(QTime(0,0))
            self.st_time_to.setTime(QTime(23,59))
            self.st_usage_combo.setCurrentIndex(0)
        self.refresh_table()

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
        # Columns are now shifted due to "Color Preview"
        in_time = self.types_table.item(row, 5).text()
        out_time = self.types_table.item(row, 6).text()
        self.in_time_edit.setTime(QTime.fromString(in_time, "HH:mm"))
        self.out_time_edit.setTime(QTime.fromString(out_time, "HH:mm"))
        self.current_old_code = code

    def clear_form(self):
        self.current_type_id = None
        self.current_old_code = None
        self.name_input.clear()
        self.code_input.clear()
        self.color_display.setText("#FFC000")
        self.in_time_edit.setTime(QTime(8, 0))
        self.out_time_edit.setTime(QTime(17, 0))
        self.types_table.clearSelection()

    def save_type(self):
        name = self.name_input.text().strip()
        code = self.code_input.text().strip().upper()
        color_hex = self.color_display.text().strip() or "#FFC000"
        in_time = self.in_time_edit.time().toString("HH:mm")
        out_time = self.out_time_edit.time().toString("HH:mm")

        if not name or not code:
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Warning)
            box.setWindowTitle("Incomplete Data")
            box.setText("Name and Code are required.")
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
            box.exec()
            return

        if self.current_type_id:
            ok, msg, old_code, new_code = db.update_shift_type(
                self.current_type_id,
                self.source,
                name,
                code,
                color_hex,
                in_time,
                out_time,
            )
            if ok:
                # If the code changed -> update Excel
                if old_code and new_code and old_code != new_code:
                    excel.apply_shift_type_update_to_excel(
                        self.excel_file, self.source, old_code, new_code, color_hex
                    )
                db.log_event(
                    self.logged_username,
                    self.source,
                    "SHIFT_TYPE_UPDATE",
                    f"{old_code} -> {new_code} | {name} {in_time}-{out_time} {color_hex}",
                )
                self.types_changed.emit(self.source)
            box = QMessageBox(self)
            box.setIcon(
                QMessageBox.Icon.Information if ok else QMessageBox.Icon.Warning
            )
            box.setWindowTitle("Save" if ok else "Error")
            box.setText(msg)
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
            box.exec()
        else:
            ok, msg = db.create_shift_type(
                self.source, name, code, color_hex, in_time, out_time
            )
            if ok:
                db.log_event(
                    self.logged_username,
                    self.source,
                    "SHIFT_TYPE_CREATE",
                    f"{code} | {name} {in_time}-{out_time} {color_hex}",
                )
                self.types_changed.emit(self.source)
            box = QMessageBox(self)
            box.setIcon(
                QMessageBox.Icon.Information if ok else QMessageBox.Icon.Warning
            )
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
        self._filter_state['text'] = self.st_search_input.text()
        self._filter_state['in_from'] = self.st_time_from.time().toString("HH:mm")
        self._filter_state['in_to'] = self.st_time_to.time().toString("HH:mm")
        self._filter_state['usage'] = self.st_usage_combo.currentText() if self.st_usage_combo.currentIndex() > 0 else None
        
        types = db.get_shift_types_filtered(
            source=self.source,
            text=self._filter_state['text'],
            in_from=self._filter_state['in_from'],
            in_to=self._filter_state['in_to'],
            usage=self._filter_state['usage']
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
        self.types_table.horizontalHeader().setSectionResizeMode(preview_col_index, QHeaderView.ResizeMode.ResizeToContents)
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
        h.addWidget(btn_new)
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
        sort = 'source' if self.sort_combo.currentIndex() == 1 else 'name'
        
        self._filter_state['source'] = src
        self._filter_state['text'] = text
        self._filter_state['sort'] = sort

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
        self.plan_widget = PlanStaffWidget(
            self.user_role, self.excel_file, self.logged_username
        )
        tabs.addTab(self.plan_widget, "📅 Plan Staff & Reports")

        # 2) Rotation History (new tab, no ID column)
        self.rotation_widget = RotationHistoryWidget(created_by=self.logged_username)  # ← NUEVO: solo mis registros
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
        # Refresh Pick Up / Drop Off dropdowns when locations change
        self.location_widget.locations_changed.connect(
            self.plan_widget.load_location_options
        )

        # 6) Settings Tab
        self.settings_widget = ReportSettingsWidget(
            self.logged_username, self.user_role
        )
        tabs.addTab(self.settings_widget, "⚙️ Settings")

        # Hot sync
        self.crud_widget.users_changed.connect(lambda src: self.plan_widget.refresh_ui_data() if src == self.user_role else None)
        self.crud_widget.import_done.connect(lambda src: self.plan_widget.refresh_ui_data() if src == self.user_role else None)

    def _sync_after_users_changed(self, src: str):
        if src == self.user_role:
            self.plan_widget.refresh_users_only()
            QApplication.processEvents()  # ensure UI repaints

    def handle_logout(self):
        self.logout_signal.emit()
        self.close()


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

        # 5) Rotation History (global; no ID column)
        self.rotation_history = RotationHistoryWidget(created_by=None) # AHORA (modo admin = None → sin filtro)
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
        
        self.rgm_crud.users_changed.connect(
            lambda src: self.rgm_plan.refresh_ui_data()
        )
        self.rgm_crud.import_done.connect(
            lambda src: self.rgm_plan.refresh_ui_data()
        )

        self.nm_crud.users_changed.connect(
            lambda src: self.nm_plan.refresh_ui_data()
        )
        self.nm_crud.import_done.connect(
            lambda src: self.nm_plan.refresh_ui_data()
        )

        # Estos de abajo ya estaban bien, los dejas igual:
        self.rgm_types.types_changed.connect(
            lambda src: self.rgm_plan.refresh_ui_data()
        )
        self.nm_types.types_changed.connect(lambda src: self.nm_plan.refresh_ui_data())

    def handle_logout(self):
        self.logout_signal.emit()
        self.close()
