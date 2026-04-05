# common_widgets.py
# Extracted from main_window.py — Shared / common widget classes and helpers.

import logging
from datetime import date as pydate, time as dtime

from PyQt6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QComboBox,
    QDateEdit,
    QTimeEdit,
    QCheckBox,
    QGroupBox,
    QSizePolicy,
    QToolButton,
    QGraphicsDropShadowEffect,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QHeaderView,
    QStyledItemDelegate,
)
from PyQt6.QtCore import (
    Qt,
    QRect,
    QSize,
    QSignalBlocker,
)
from PyQt6.QtGui import QColor, QFont, QIcon, QPixmap, QPainter, QPen

import database_logic as db
from constants import WEEKEND_HEADER_YELLOW

logger = logging.getLogger(__name__)


# -------------------------------------------------------------
# Module-level helper functions
# -------------------------------------------------------------

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
# ShiftCellDelegate
# -------------------------------------------------------------
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
