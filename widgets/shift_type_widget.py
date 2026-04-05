# widgets/shift_type_widget.py
# Extracted from main_window.py — ShiftTypeAdminWidget

import logging

from PyQt6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QLabel,
    QLineEdit,
    QComboBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QGroupBox,
    QMessageBox,
    QColorDialog,
    QTimeEdit,
)
from PyQt6.QtCore import Qt, pyqtSignal, QTime, QTimer, QSignalBlocker, QSize
from PyQt6.QtGui import QColor, QIcon, QPixmap, QPainter, QPen

import database_logic as db
import excel_logic as excel
from constants import DEBOUNCE_MS

logger = logging.getLogger(__name__)


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

        # --- Dropdown de Behavior (reemplaza 3 checkboxes) ---
        self.behavior_combo = QComboBox()
        self.behavior_combo.addItem("Normal (transport required)", "normal")
        self.behavior_combo.addItem("Treat as OFF (Non-working day)", "off")
        self.behavior_combo.addItem("No Transport Required (working, off-site)", "no_transport")
        if self.source == "RGM":
            self.behavior_combo.addItem(
                "Apply 1+D logic (Exit date = End Date + 1)", "apply_1d"
            )
        self.behavior_combo.setToolTip(
            "Normal: standard shift requiring site transport.\n"
            "OFF: non-working day (Vacation, Sick Leave, etc.) — excluded from transport & stay reports.\n"
            "No Transport: person works but does NOT need site transport (Office day, off-site training).\n"
            "1+D (RGM only): exit date = End Date + 1 day, triggers consecutive-shift dialog."
        )
        form_layout.addWidget(QLabel("Behavior:"), 5, 0)
        form_layout.addWidget(self.behavior_combo, 5, 1)

        actions = QHBoxLayout()
        #actions.addWidget(self.new_btn)
        actions.addWidget(self.save_btn)
        actions.addWidget(self.delete_btn)

        # NOTA: Row 8 para dejar espacio al checkbox apply_1d (row 7, solo RGM)
        form_layout.addLayout(actions, 8, 0, 1, 2)

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
        # --- Conectar combo Behavior ---
        self.behavior_combo.currentIndexChanged.connect(self._on_behavior_changed)
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

    def _behavior_flags(self) -> dict:
        """Devuelve los 3 flags booleanos según la selección del dropdown Behavior."""
        mode = self.behavior_combo.currentData()
        return {
            "is_off":        mode == "off",
            "no_transport":  mode == "no_transport",
            "apply_1d":      mode == "apply_1d",
        }

    def _on_behavior_changed(self, _idx: int):
        """Reacciona al cambio del combo: oculta/muestra los campos de tiempo."""
        flags = self._behavior_flags()
        self.toggle_time_inputs(flags["is_off"])

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

        # --- Cargar Behavior desde BD → dropdown ---
        all_types = db.get_shift_types(self.source)
        record = next((t for t in all_types if t['id'] == self.current_type_id), None)

        if record:
            is_off_val       = bool(record.get('is_off', 0))
            no_transport_val = bool(record.get('no_transport', 0))
            apply_1d_val     = bool(record.get('apply_1d', 0))
            if is_off_val:
                mode = "off"
            elif no_transport_val:
                mode = "no_transport"
            elif self.source == "RGM" and apply_1d_val:
                mode = "apply_1d"
            else:
                mode = "normal"
        else:
            mode = "normal"

        with QSignalBlocker(self.behavior_combo):
            i = self.behavior_combo.findData(mode)
            if i >= 0:
                self.behavior_combo.setCurrentIndex(i)

        # Forzar actualización visual de los campos de tiempo
        self.toggle_time_inputs(mode == "off")

        # --- GUARDRAILS: Bloquear controles para System Types ---
        _is_system = bool(record.get("is_system", 0)) if record else False

        # Bloquear el campo Code (ON/ON NS/OFF no pueden renombrarse)
        self.code_input.setReadOnly(_is_system)
        self.code_input.setStyleSheet(
            "background: #F3F4F6; color: #9CA3AF;" if _is_system else ""
        )
        self.code_input.setToolTip(
            "System type - code cannot be changed." if _is_system
            else "Short code used in the schedule (e.g. SOP, STP)"
        )

        # Deshabilitar Delete para system types
        self.delete_btn.setEnabled(not _is_system)
        self.delete_btn.setToolTip(
            "System types cannot be deleted." if _is_system else ""
        )

        # El combo Behavior es invariante en system types
        self.behavior_combo.setEnabled(not _is_system)

    def clear_form(self):
        self.current_type_id = None
        self.current_old_code = None
        self.name_input.clear()
        self.code_input.clear()
        self.color_display.setText("#FFC000")
        self.in_time_edit.setTime(QTime(8, 0))
        self.out_time_edit.setTime(QTime(17, 0))

        # Resetear Behavior al estado "Normal"
        with QSignalBlocker(self.behavior_combo):
            i = self.behavior_combo.findData("normal")
            if i >= 0:
                self.behavior_combo.setCurrentIndex(i)

        # Asegurar que los tiempos sean visibles al limpiar
        self.toggle_time_inputs(False)

        # Restaurar controles que pudieron quedar bloqueados por un system type
        self.code_input.setReadOnly(False)
        self.code_input.setStyleSheet("")
        self.code_input.setToolTip("Short code used in the schedule (e.g. SOP, STP)")
        self.delete_btn.setEnabled(True)
        self.delete_btn.setToolTip("")
        self.behavior_combo.setEnabled(True)

        self.types_table.clearSelection()

    def save_type(self):
        name = self.name_input.text().strip()
        code = self.code_input.text().strip().upper()
        color_hex = self.color_display.text().strip() or "#FFC000"
        in_time = self.in_time_edit.time().toString("HH:mm")
        out_time = self.out_time_edit.time().toString("HH:mm")

        # Leer flags del dropdown Behavior
        flags = self._behavior_flags()
        is_off_val       = flags["is_off"]
        no_transport_val = flags["no_transport"]
        apply_1d_val     = flags["apply_1d"]

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
                is_off=is_off_val,
                no_transport=no_transport_val,
                apply_1d=apply_1d_val,
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
                    f"{old_code} -> {new_code} | Off={is_off_val} | NoTransport={no_transport_val} | Apply1D={apply_1d_val}"
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
                is_off=is_off_val,
                no_transport=no_transport_val,
                apply_1d=apply_1d_val,
            )
            if ok:
                db.log_event(
                    self.logged_username,
                    self.source,
                    "SHIFT_TYPE_CREATE",
                    f"{code} | Off={is_off_val} | NoTransport={no_transport_val} | Apply1D={apply_1d_val}"
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
        headers = ["ID", "Name", "Code", "Color", "Color Preview", "IN", "OUT", "Type"]
        self.types_table.setRowCount(len(types))
        self.types_table.setColumnCount(len(headers))
        self.types_table.setHorizontalHeaderLabels(headers)
        for r, t in enumerate(types):
            self.types_table.setItem(r, 0, QTableWidgetItem(str(t["id"])))
            self.types_table.setItem(r, 1, QTableWidgetItem(t["name"]))
            self.types_table.setItem(r, 2, QTableWidgetItem(t["code"]))
            self.types_table.setItem(r, 3, QTableWidgetItem(t["color_hex"]))

            # Color preview
            preview_item = QTableWidgetItem()
            preview_item.setIcon(_create_color_icon(t["color_hex"]))
            self.types_table.setItem(r, 4, preview_item)

            self.types_table.setItem(r, 5, QTableWidgetItem(t["in_time"]))
            self.types_table.setItem(r, 6, QTableWidgetItem(t["out_time"]))

            # Columna Type: System vs Custom
            _is_sys = bool(t.get("is_system", 0))
            type_item = QTableWidgetItem("🔒 System" if _is_sys else "Custom")
            type_item.setForeground(QColor("#1565C0") if _is_sys else QColor("#374151"))
            if _is_sys:
                _f = type_item.font()
                _f.setBold(True)
                type_item.setFont(_f)
            self.types_table.setItem(r, 7, type_item)

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
