# plan_staff_widget.py
# Extracted from main_window.py — PlanStaffWidget class and its module-level helpers.

import sqlite3
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment
import json
import os
import logging

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
    QStyledItemDelegate,
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
from PyQt6.QtGui import (
    QColor,
    QFont,
    QCursor,
    QIcon,
    QPixmap,
    QPainter,
    QPen,
    QKeySequence,
    QUndoCommand,
    QUndoStack,
    QAction,
)
from datetime import datetime, date as pydate, timedelta, time as dtime

from clipboard_logic import ScheduleClipboardService
from widgets.undo_commands import UndoCellChangeCommand, UndoScheduleSnapshotCommand
from widgets.common_widgets import (
    ShiftCellDelegate,
    ShiftInfoCard,
    CollapsibleGroupBox,
    DayScheduleEditor,
    WeekendHeader,
)
from ui.theme import mark_error
from constants import (
    WARN_BG_HEX,
    FROZEN_COLUMN_COUNT,
    DEBOUNCE_MS,
    WEEKEND_HEADER_YELLOW,
    NEWMONT_REPORT_HEADERS,
    RGM_REPORT_HEADERS,
    Source,
    ShiftStatus,
    ShiftColor,
)

import database_logic as db
import excel_logic as excel

logger = logging.getLogger(__name__)


# ─── Module-level helpers (originally in main_window.py) ─────────────────────


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


# ─── PlanStaffWidget ─────────────────────────────────────────────────────────


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

        # --- FIX: Puente de datos para hora manual del diálogo Separar ---
        self._last_manual_exit_config = None

        self._is_first_load = True

        # ── Undo / Redo stack (Ctrl+Z / Ctrl+Y) ────────────────────────────────
        self._undo_stack = QUndoStack(self)
        self._undo_stack.setUndoLimit(50)   # limitar memoria en sesiones largas
        self._undo_in_progress = False       # previene re-entrada durante undo/redo

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

        # --- Row reorder: context menu on frozen table ---
        self.frozen_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.frozen_table.customContextMenuRequested.connect(self._show_row_reorder_menu)

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

    # ═══════════════════════════════════════════════════════════════════════════
    # UNDO / REDO SUPPORT
    # ═══════════════════════════════════════════════════════════════════════════

    def keyPressEvent(self, event):
        """
        Unified key handler: Undo/Redo, Copy, Paste.
        • Si el foco está en un widget de texto → deja que ese widget haga su undo nativo.
        • Si el foco está en la tabla de schedule → usa el QUndoStack del sistema.
        """
        from PyQt6.QtWidgets import QLineEdit, QTextEdit, QPlainTextEdit, QApplication
        w = QApplication.focusWidget()
        if isinstance(w, (QLineEdit, QTextEdit, QPlainTextEdit)):
            super().keyPressEvent(event)
            return

        if event.matches(QKeySequence.StandardKey.Undo):
            if self._undo_stack.canUndo():
                self._undo_in_progress = True
                try:
                    self._undo_stack.undo()
                finally:
                    self._undo_in_progress = False
            event.accept()
            return

        if event.matches(QKeySequence.StandardKey.Redo):
            if self._undo_stack.canRedo():
                self._undo_in_progress = True
                try:
                    self._undo_stack.redo()
                finally:
                    self._undo_in_progress = False
            event.accept()
            return

        if event.matches(QKeySequence.StandardKey.Copy):
            ScheduleClipboardService.copy_to_clipboard(
                self.schedule_table,
                self._row_identities,
                self._date_col_dates,
                self._custom_shift_map
            )
            return

        if event.matches(QKeySequence.StandardKey.Paste):
            self._handle_paste_operation()
            return

        super().keyPressEvent(event)

    def _patch_cells_in_table(self, badge: str, start_date, end_date, status: str):
        """
        Actualiza visualmente sólo las celdas afectadas por undo/redo
        sin recargar todo el Excel.  Mucho más rápido que refresh_ui_data().
        Si el badge no está visible en la vista actual, hace un refresh completo
        como fallback seguro.
        """
        # 1. Buscar la fila del empleado
        row_idx = None
        for i, identity in enumerate(self._row_identities):
            if identity.get("badge") == badge:
                row_idx = i
                break

        if row_idx is None:
            # El empleado no está en la vista actual → recarga completa
            self.refresh_ui_data()
            return

        # 2. Actualizar celdas en el rango de fechas
        from datetime import timedelta
        d = start_date
        with QSignalBlocker(self.schedule_table):
            while d <= end_date:
                try:
                    col_idx = self._date_col_dates.index(d)
                except ValueError:
                    d += timedelta(days=1)
                    continue

                item = self.schedule_table.item(row_idx, col_idx)
                if not item:
                    item = QTableWidgetItem()
                    self.schedule_table.setItem(row_idx, col_idx, item)

                display_text = status or ""
                item.setText(display_text)
                self._apply_status_background(item, display_text)
                self._cell_original_values[(row_idx, col_idx)] = display_text
                d += timedelta(days=1)

        self.schedule_table.viewport().update()
        self.rotation_changed.emit()

    def _patch_cells_in_table_from_map(
        self, badge: str, start_date, end_date, schedule_map: dict
    ):
        """
        Versión de _patch_cells_in_table que acepta un mapa {date_iso: {status, ...}}
        en lugar de un único status uniforme.  Necesario para undo/redo de drag-fill
        donde el estado anterior puede tener valores distintos por día.
        """
        row_idx = None
        for i, identity in enumerate(self._row_identities):
            if identity.get("badge") == badge:
                row_idx = i
                break

        if row_idx is None:
            self.refresh_ui_data()
            return

        d = start_date
        with QSignalBlocker(self.schedule_table):
            while d <= end_date:
                try:
                    col_idx = self._date_col_dates.index(d)
                except ValueError:
                    d += timedelta(days=1)
                    continue

                rec = schedule_map.get(d.isoformat())
                display_text = (rec.get("status") or "") if rec else ""

                item = self.schedule_table.item(row_idx, col_idx)
                if not item:
                    item = QTableWidgetItem()
                    self.schedule_table.setItem(row_idx, col_idx, item)

                item.setText(display_text)
                self._apply_status_background(item, display_text)
                self._cell_original_values[(row_idx, col_idx)] = display_text
                d += timedelta(days=1)

        self.schedule_table.viewport().update()
        self.rotation_changed.emit()

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
        """
        Carga todos los shift types desde DB en el combo de Status/Shift.

        ORDEN EN EL COMBO:
        1. Opcion vacia (Do Not Mark Days)
        2. System types (OFF, ON, ON NS) leidos desde DB con sus colores
        3. Separador visual
        4. Custom types (is_system=0)

        FALLBACK: Si los system types no estan en DB aun, muestra opciones
        hardcoded como proteccion para que la UI siempre funcione.
        """
        self.status_selector.blockSignals(True)
        self.status_selector.clear()

        # Opcion vacia
        self.status_selector.addItem(QIcon(), "\u2014 Do Not Mark Days \u2014", {"kind": "none"})

        # Cargar TODOS los shift types desde DB (system + custom)
        all_types = db.get_shift_types(self.source)

        # Actualizar el mapa en memoria (ahora incluye ON/ON NS/OFF)
        self._custom_shift_map = {t["code"].strip().upper(): t for t in all_types}

        # Separar system types de custom types
        system_types = [t for t in all_types if t.get("is_system")]
        custom_types  = [t for t in all_types if not t.get("is_system")]

        # Orden preferido para system types: OFF primero, luego ON, luego ON NS
        _SYSTEM_ORDER = {"OFF": 0, "ON": 1, "ON NS": 2}
        system_types.sort(key=lambda t: _SYSTEM_ORDER.get(t["code"].strip().upper(), 99))

        for t in system_types:
            code = t["code"].strip().upper()
            label = t["name"]
            data = {
                "kind": "base",
                "status": code,
                "shift_type": t["name"],
                "in_time": t.get("in_time"),
                "out_time": t.get("out_time"),
                "is_off": bool(t.get("is_off", 0)),
                "color_hex": t.get("color_hex", "#FFFFFF"),
            }
            self.status_selector.addItem(
                _create_color_icon(t["color_hex"]),
                label,
                data,
            )

        # FALLBACK: si DB esta vacia / seed no ejecutado aun
        if not system_types:
            _fallback = [
                ("#FFC7CE", "OFF",              {"kind":"base","status":"OFF","shift_type":None,"in_time":None,"out_time":None,"is_off":True}),
                ("#C6EFCE", "ON (Day Shift)",   {"kind":"base","status":"ON","shift_type":"Day Shift","in_time":None,"out_time":None}),
                ("#FFFF99", "ON NS (Night Shift)",{"kind":"base","status":"ON NS","shift_type":"Night Shift","in_time":None,"out_time":None}),
            ]
            for color, label, data in _fallback:
                self.status_selector.addItem(_create_color_icon(color), label, data)

        # Custom types: separador + lista
        if custom_types:
            self.status_selector.addItem(
                QIcon(), "\u2014\u2014 Custom Shift Types \u2014\u2014", {"kind": "separator"}
            )
            for t in custom_types:
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
                        "is_off": bool(t.get("is_off", 0)),
                    },
                )

        self.status_selector.setCurrentIndex(0)
        self.status_selector.blockSignals(False)



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

    # ═══════════════════════════════════════════════════════════════════════════
    # ROW REORDER (Move Up / Move Down)
    # ═══════════════════════════════════════════════════════════════════════════

    def _show_row_reorder_menu(self, pos):
        """Context menu on frozen_table for row reorder operations."""
        from PyQt6.QtWidgets import QMenu
        row = self.frozen_table.rowAt(pos.y())
        if row < 0 or row >= len(self._row_identities):
            return

        identity = self._row_identities[row]
        badge = identity.get("badge", "").strip()
        name = identity.get("name", "").strip()
        if not badge:
            return

        menu = QMenu(self)
        menu.setStyleSheet("QMenu { font-size: 13px; }")

        # Move Up (not available if already first row)
        action_up = menu.addAction("⬆️  Move Up")
        action_up.setEnabled(row > 0)

        # Move Down (not available if already last row)
        action_down = menu.addAction("⬇️  Move Down")
        action_down.setEnabled(row < len(self._row_identities) - 1)

        menu.addSeparator()

        # Header with user info (non-interactive)
        info_action = menu.addAction(f"👤 {name} [{badge}]")
        info_action.setEnabled(False)

        chosen = menu.exec(self.frozen_table.viewport().mapToGlobal(pos))
        if chosen is None:
            return

        if chosen == action_up:
            self._move_row_direction(row, badge, direction="up")
        elif chosen == action_down:
            self._move_row_direction(row, badge, direction="down")

    def _move_row_direction(self, current_row: int, badge: str, direction: str):
        """
        Move a row up or down by swapping with the adjacent row.
        Atomic operation: DB swap → Excel swap → UI refresh.
        """
        if direction == "up" and current_row <= 0:
            return
        if direction == "down" and current_row >= len(self._row_identities) - 1:
            return

        # Determine neighbor
        neighbor_row = current_row - 1 if direction == "up" else current_row + 1
        neighbor_identity = self._row_identities[neighbor_row]
        neighbor_badge = neighbor_identity.get("badge", "").strip()

        if not neighbor_badge:
            return

        # 1. Swap in DB (display_order)
        ok, msg = db.swap_user_order(self.source, badge, neighbor_badge)
        if not ok:
            QMessageBox.warning(self, "Move Error", msg)
            return

        # 2. Swap in Excel (full row with styles/values/comments)
        try:
            ok_xl, msg_xl = excel.swap_rows_in_excel(
                self.excel_file, badge, neighbor_badge
            )
            if not ok_xl:
                print(f"Excel swap warning: {msg_xl}")
        except Exception as e:
            print(f"Excel swap failed: {e}")

        # 3. Refresh UI (rebuilds both tables from Excel)
        self.refresh_ui_data()

        # 4. Re-select the moved row in the new position
        new_row = neighbor_row
        if 0 <= new_row < self.frozen_table.rowCount():
            self.frozen_table.selectRow(new_row)

    def refresh_ui_data(self, use_preloaded=False):
        self.load_shift_type_options()
        self.load_schedule_data(use_preloaded=use_preloaded) # Pasar la bandera
        self.load_users_to_selector()
        self.load_location_options()  # keep combos in sync with Location admin
        self.remarks_input.clear()  # Clear remarks on refresh

    # [CAMBIO 5] Lógica crítica para usar los datos en memoria
    def load_schedule_data(self, use_preloaded=False):
        """
        Carga los datos del cronograma en la tabla, optimizado para evitar
        parpadeos (flickering) y saltos de scroll inesperados.
        """
        # -------------------------------------------------------------
        # 1. OPTIMIZACIÓN: GUARDAR POSICIÓN DEL SCROLL (Evitar Salto)
        # -------------------------------------------------------------
        h_scroll_val = 0
        v_scroll_val = 0

        # Guardamos dónde está mirando el usuario antes de tocar nada
        if self.schedule_table.horizontalScrollBar():
            h_scroll_val = self.schedule_table.horizontalScrollBar().value()
        if self.schedule_table.verticalScrollBar():
            v_scroll_val = self.schedule_table.verticalScrollBar().value()

        # -------------------------------------------------------------
        # 2. OPTIMIZACIÓN: CONGELAR ACTUALIZACIONES DE UI (Evitar Parpadeo)
        # -------------------------------------------------------------
        # Bloqueamos el repintado de los widgets principales.
        # Esto evita que el usuario vea la tabla "vaciarse" y volver a llenarse.
        self.setUpdatesEnabled(False)
        self.schedule_table.setUpdatesEnabled(False)
        self.frozen_table.setUpdatesEnabled(False)

        try:
            # -------------------------------------------------------------
            # LÓGICA DE PRE-CARGA DE DATOS (Tu lógica original)
            # -------------------------------------------------------------
            df = None
            if use_preloaded and hasattr(self, '_initial_preloaded_df') and self._initial_preloaded_df is not None:
                print("DEBUG: Using preloaded dataframe (High Performance Mode).")
                df = self._initial_preloaded_df
                self._initial_preloaded_df = None
            else:
                df = excel.get_schedule_preview(self.excel_file)

            # --- Overlay de Usuarios desde BD (Tu lógica original) ---
            try:
                if df is not None and not df.empty:
                    users_db = db.get_all_users(self.source)
                    role_map = {str(u["badge"]).strip(): str(u.get("role") or "").strip() for u in users_db}
                    name_map = {str(u["badge"]).strip(): str(u.get("name") or "").strip() for u in users_db}

                    col_badge = "BADGE" if "BADGE" in df.columns else "Company ID"

                    if col_badge in df.columns:
                        df[col_badge] = df[col_badge].astype(str).str.strip()
                        col_role = "ROLE" if "ROLE" in df.columns else "Discipline"
                        if col_role in df.columns:
                            df[col_role] = df[col_badge].map(role_map).fillna(df[col_role])

                        if "NAME" in df.columns:
                            df["NAME"] = df[col_badge].map(name_map).fillna(df["NAME"])

                        print(f" UI Overlay applied: User metadata synced with DB for display.")
            except Exception as e:
                print(f" Warning: Could not apply DB overlay to preview: {e}")

            self._loading_preview = True
            self._cell_original_values.clear()
            self._row_identities.clear()
            self._date_col_dates.clear()

            # Caso tabla vacía
            if df is None or df.empty:
                self.frozen_table.clear()
                self.schedule_table.clear()
                self.frozen_table.setRowCount(0)
                self.schedule_table.setRowCount(0)
                self._loading_preview = False
                # Aunque esté vacía, debemos asegurar que el 'finally' se ejecute para restaurar UI
                return

            # Auto-Horizon check (Tu lógica original)
            if not use_preloaded:
                try:
                    ok_horizon, msg_horizon = excel.ensure_rolling_horizon_columns(self.excel_file)
                    if ok_horizon:
                        print(f"[AUTO-HORIZON] {msg_horizon}")
                except Exception as e:
                    print(f"[AUTO-HORIZON CRITICAL] Failed: {e}")

            custom_map = db.get_shift_type_map(self.source)

            # -----------------------------------------------------------
            # ORDENAMIENTO Y ESTRUCTURA (Tu lógica original)
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

            # Configuración de Cabeceras
            actual_frozen_count = min(df.shape[1], FROZEN_COLUMN_COUNT)
            frozen_headers = [str(c) for c in cols[:actual_frozen_count]]

            schedule_headers = []
            for d in date_cols:
                schedule_headers.append(f"{d.isoformat()}\n{d.strftime('%a')}")
            self._date_col_dates = list(date_cols)
            self.weekend_header.set_dates(self._date_col_dates)

            # Configurar Tablas (Filas y Columnas)
            self.frozen_table.setRowCount(df.shape[0])
            self.frozen_table.setColumnCount(actual_frozen_count)
            self.frozen_table.setHorizontalHeaderLabels(frozen_headers)

            self.schedule_table.setRowCount(df.shape[0])
            self.schedule_table.setColumnCount(len(schedule_headers))
            for idx, header_text in enumerate(schedule_headers):
                self.schedule_table.setHorizontalHeaderItem(
                    idx, QTableWidgetItem(header_text)
                )

            # Cargar Filas (Loop Principal)
            for i, row in df.iterrows():
                badge_val = row.get("BADGE") if hasattr(row, "get") else (row["BADGE"] if "BADGE" in df.columns else "")
                name_val = row.get("NAME") if hasattr(row, "get") else (row["NAME"] if "NAME" in df.columns else "")
                role_val = row.get("ROLE") if hasattr(row, "get") else (row["ROLE"] if "ROLE" in df.columns else "")

                self._row_identities.append({
                    "badge": str(badge_val) if badge_val is not None else "",
                    "name": str(name_val) if name_val is not None else "",
                    "role": str(role_val) if role_val is not None else "",
                })

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

                        self._apply_status_background(item, val_str)

                        # Protección Legacy: Si el Excel tiene "DAY" o números (1,2,3) y no es un turno custom
                        # forzamos el verde (ON) para que no se vea blanco.
                        if val_str and val_str not in ("ON", "ON NS", "NIGHT", "OFF", "BREAK", "KO", "LEAVE"):
                            if not (self._custom_shift_map or {}).get(val_str):
                                if "DAY" in val_str or val_str.isdigit():
                                    item.setBackground(QColor("#C6EFCE"))

                        self.schedule_table.setItem(i, col_index, item)
                        self._cell_original_values[(i, col_index)] = val_str

                        key = self._warn_key_for(i, col_index)
                        if key in self._warn_highlight_keys:
                            item.setBackground(QColor(WARN_BG_HEX))

            # Estilos y Ajustes (Dentro del try, oculto por setUpdatesEnabled=False)
            self.frozen_table.resizeColumnsToContents()
            self.schedule_table.resizeColumnsToContents()

            compact_height = 35
            compact_font = QFont(); compact_font.setPointSize(7); compact_font.setBold(True)

            header_stylesheet = """
                QHeaderView::section {
                    padding-top: 0px; padding-bottom: 0px;
                    padding-left: 2px; padding-right: 2px;
                    margin: 0px; border-bottom: 1px solid #ccc; border-right: 1px solid #ccc;
                }
            """

            for table in [self.schedule_table, self.frozen_table]:
                h = table.horizontalHeader()
                h.setFont(compact_font)
                h.setFixedHeight(compact_height)
                h.setStyleSheet(header_stylesheet)
                h.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)

            self._update_frozen_width()
            self._loading_preview = False

        finally:
            # -------------------------------------------------------------
            # 3. DESCONGELAR UI (Pintar todo de una sola vez)
            # -------------------------------------------------------------
            # Pase lo que pase (éxito o error), reactivamos la interfaz.
            self.schedule_table.setUpdatesEnabled(True)
            self.frozen_table.setUpdatesEnabled(True)
            self.setUpdatesEnabled(True)

        # -------------------------------------------------------------
        # 4. RESTAURAR SCROLL O CENTRAR (Lógica de "First Load")
        # -------------------------------------------------------------
        if self._is_first_load:
            # Si es la PRIMERA VEZ, centramos en hoy.
            self._center_today_column()
            self._is_first_load = False # Marcamos como ya cargado.
        else:
            # Si es un REFRESCO, restauramos la posición anterior.
            self.schedule_table.horizontalScrollBar().setValue(h_scroll_val)
            self.schedule_table.verticalScrollBar().setValue(v_scroll_val)

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

        # ARREGLO: Solo centrar si es la primera vez absoluta que se muestra
        # Esto previene saltos si el usuario cambia de pestaña y regresa.
        if self._is_first_load:
            self._center_today_column()
            # Nota: No ponemos False aquí inmediatamente, dejamos que load_schedule_data lo maneje

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

    def _apply_status_background(self, item: QTableWidgetItem, value: str | None) -> None:
        """
        UNIFIED background renderer — handles base codes AND custom shift types.
        Must be the ONLY place where cell background color is decided.
        Replaces the old duplicated _apply_base_background().
        """
        from PyQt6.QtGui import QBrush
        s = (value or "").strip().upper()

        # 1) Empty → clear background
        if not s:
            item.setBackground(QBrush())  # NoBrush / transparent
            return

        # 2) Custom shift types first (code → color_hex from DB)
        info = (self._custom_shift_map or {}).get(s)
        if info and info.get("color_hex"):
            item.setBackground(QColor(info["color_hex"]))
            return

        # 3) Base status codes
        if s == "ON":
            item.setBackground(QColor("#C6EFCE"))
        elif s in ("ON NS", "NIGHT"):
            item.setBackground(QColor("#FFFF99"))
        elif s in ("OFF", "BREAK", "KO", "LEAVE"):
            item.setBackground(QColor("#FFC7CE"))
        else:
            # 4) Unknown code → clear (avoid residual colors from previous value)
            item.setBackground(QBrush())

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
        Resuelve el tiempo (IN o OUT) para un status dado.

        ORDEN DE PRIORIDAD:
        1. DB via self._custom_shift_map (incluye system types ON/ON NS/OFF)
        2. Hardcode legacy Newmont/RGM (fallback si DB vacia / seed no ejecutado)
        3. raw_time del arrastre/copia
        4. dtime(0,0) como ultimo recurso
        """
        if not status:
            return dtime(0, 0)

        status_key = status.strip().upper()

        # PRIORIDAD 1: Consultar DB (system types + custom types)
        if status_key in (self._custom_shift_map or {}):
            shift_info = self._custom_shift_map[status_key]
            time_str = shift_info.get("in_time") if kind == "IN" else shift_info.get("out_time")
            if time_str and time_str not in ("00:00", "0:00"):
                return _parse_hhmm_to_time(time_str, dtime(0, 0))
            # time_str es 00:00 (tipico de OFF) o vacio
            return dtime(0, 0)

        # PRIORIDAD 2: Hardcode legacy (solo si status no esta en DB)
        if self.source == "Newmont":
            if status_key == "ON":
                return dtime(6, 0) if kind == "IN" else dtime(12, 0)
            elif status_key == "ON NS":
                return dtime(12, 0) if kind == "IN" else dtime(6, 0)
        elif self.source == "RGM":
            if status_key in ("ON", "ON NS"):
                return dtime(7, 0)

        # PRIORIDAD 3: raw_time del origen (drag-fill/clipboard)
        return _parse_hhmm_to_time(raw_time, dtime(0, 0))
    def _consolidate_and_record_logistics(self, badge, role, username, op_start, op_end, status, cursor=None, manual_split_config=None):
        """
        MOTOR DE CONSOLIDACIÓN v3 — SCAN-BASED.

        En lugar de merge incremental (que fusionaba ON con ONH),
        este motor:
        1. Determina la "zona afectada" (todas las ops que tocan nuestro rango)
        2. Borra TODAS las operaciones de esa zona
        3. Escanea el schedule día a día en esa zona
        4. Crea UNA operación por cada bloque contiguo de MISMO status

        Esto garantiza que:
        - ON nunca se fusiona con ONH
        - Los días huérfanos se reconstruyen
        - Los force_new_entry se respetan como boundaries
        """
        # A. Si es día libre, solo limpiar
        if not db.is_working_status(status, self.source):
            db.delete_operations_in_range(badge, op_start, op_end, cursor=cursor)
            return

        # ---------------------------------------------------------------
        # 1. DETERMINAR LA ZONA AFECTADA
        #    = nuestro rango + cualquier operación existente que lo toque
        # ---------------------------------------------------------------
        zone_start = op_start
        zone_end = op_end

        # Extender zona para cubrir operaciones existentes que se solapan
        for check_date in [op_start, op_end]:
            existing_op = db.get_operation_overlapping(badge, check_date, cursor=cursor)
            if existing_op:
                try:
                    es = datetime.strptime(existing_op['start_date'], "%Y-%m-%d").date()
                    ee = datetime.strptime(existing_op['end_date'], "%Y-%m-%d").date()
                    if es < zone_start: zone_start = es
                    if ee > zone_end:   zone_end = ee
                except (ValueError, KeyError):
                    pass

        # Extender un paso más si el vecino tiene MISMO status (merge natural)
        # — Izquierda
        curr_start_map = db.get_schedule_map_for_range(badge, zone_start, zone_start, self.source, cursor=cursor)
        zone_start_info = curr_start_map.get(zone_start.isoformat(), {})
        zone_start_st = (zone_start_info.get("status") or "").strip().upper()
        is_force_zone_start = (zone_start_info.get("force_new_entry", 0) == 1)

        if not is_force_zone_start:
            prev_day = zone_start - timedelta(days=1)
            prev_map = db.get_schedule_map_for_range(badge, prev_day, prev_day, self.source, cursor=cursor)
            prev_info = prev_map.get(prev_day.isoformat(), {})
            prev_st = (prev_info.get("status") or "").strip().upper()
            if prev_st and db.is_working_status(prev_st, self.source):
                prev_op = db.get_operation_overlapping(badge, prev_day, cursor=cursor)
                if prev_op:
                    try:
                        ps = datetime.strptime(prev_op['start_date'], "%Y-%m-%d").date()
                        if ps < zone_start: zone_start = ps
                    except (ValueError, KeyError):
                        pass

        # — Derecha
        curr_end_map = db.get_schedule_map_for_range(badge, zone_end, zone_end, self.source, cursor=cursor)
        zone_end_info = curr_end_map.get(zone_end.isoformat(), {})
        zone_end_st = (zone_end_info.get("status") or "").strip().upper()

        next_day = zone_end + timedelta(days=1)
        next_map = db.get_schedule_map_for_range(badge, next_day, next_day, self.source, cursor=cursor)
        next_info = next_map.get(next_day.isoformat(), {})
        next_st = (next_info.get("status") or "").strip().upper()
        is_force_next = (next_info.get("force_new_entry", 0) == 1)

        if not is_force_next and next_st and db.is_working_status(next_st, self.source):
            next_op = db.get_operation_overlapping(badge, next_day, cursor=cursor)
            if next_op:
                try:
                    ne = datetime.strptime(next_op['end_date'], "%Y-%m-%d").date()
                    if ne > zone_end: zone_end = ne
                except (ValueError, KeyError):
                    pass

        # ---------------------------------------------------------------
        # 2. PRESERVAR exit_date overrides de TODAS las operaciones en la zona
        # ---------------------------------------------------------------
        # FIX BUG 2: Antes solo capturaba bordes (zone_start, zone_end).
        # Ahora captura TODAS las operaciones que se solapan con la zona
        # para no perder ningún override manual (ej: hora 18:00).
        _saved_exits = {}  # {(start_iso, end_iso): exit_date_str}

        if cursor is not None:
            cursor.execute(
                """
                SELECT start_date, end_date, exit_date
                  FROM operations
                 WHERE badge = ?
                   AND NOT (end_date < ? OR start_date > ?)
                """,
                (badge, zone_start.isoformat(), zone_end.isoformat())
            )
            for row in cursor.fetchall():
                if isinstance(row, dict) or hasattr(row, 'keys'):
                    st, en, ex = row['start_date'], row['end_date'], row['exit_date']
                else:
                    st, en, ex = row[0], row[1], row[2]
                if ex:
                    _saved_exits[(st, en)] = ex
                    print(f"[CONSOL] Preserved exit override: ({st},{en}) -> {ex}")
        else:
            _temp_conn = db.sqlite3.connect(db.DB_FILE)
            _temp_conn.row_factory = db.sqlite3.Row
            _temp_cur = _temp_conn.cursor()
            _temp_cur.execute(
                """
                SELECT start_date, end_date, exit_date
                  FROM operations
                 WHERE badge = ?
                   AND NOT (end_date < ? OR start_date > ?)
                """,
                (badge, zone_start.isoformat(), zone_end.isoformat())
            )
            for row in _temp_cur.fetchall():
                st, en, ex = row['start_date'], row['end_date'], row['exit_date']
                if ex:
                    _saved_exits[(st, en)] = ex
                    print(f"[CONSOL] Preserved exit override: ({st},{en}) -> {ex}")
            _temp_conn.close()

        # ---------------------------------------------------------------
        # 3. BORRAR todas las operaciones en la zona
        # ---------------------------------------------------------------
        db.delete_operations_in_range(badge, zone_start, zone_end, cursor=cursor)

        # ---------------------------------------------------------------
        # 4. ESCANEAR schedule y agrupar en bloques contiguos por status
        # ---------------------------------------------------------------
        full_map = db.get_schedule_map_for_range(badge, zone_start, zone_end, self.source, cursor=cursor)

        blocks = []  # [(block_start, block_end, first_status, last_status), ...]
        blk_start = None
        blk_st_first = None  # Status del primer día (para entry time)
        blk_st_last = None   # Status del último día (para exit time/date)
        current = zone_start

        while current <= zone_end:
            iso = current.isoformat()
            info = full_map.get(iso, {})
            st = (info.get("status") or "").strip().upper()
            is_w = db.is_working_status(st, self.source) if st else False
            is_f = (info.get("force_new_entry", 0) == 1)

            # ¿Este día CONTINÚA el bloque actual?
            # REGLA CLAVE (replica el comportamiento del motor OLD):
            #   - Un día CONTINÚA el bloque si es working y NO tiene force_new_entry=1
            #   - El cambio de status (ej: ON → ON NS) NO rompe el bloque
            #   - SOLO force_new_entry=1 (usuario eligió "Separar") rompe el bloque
            #   - Esto es lo que hace que "Unir (Mismo Viaje)" funcione correctamente
            continues = (is_w and not is_f and blk_start is not None)

            if not continues:
                # Cerrar bloque anterior si existe
                if blk_start is not None:
                    blocks.append((blk_start, current - timedelta(days=1), blk_st_first, blk_st_last))
                # ¿Iniciar nuevo bloque?
                if is_w:
                    blk_start = current
                    blk_st_first = st  # Primer status del bloque (para entry time)
                    blk_st_last = st   # Último status (se actualiza día a día)
                else:
                    blk_start = None
                    blk_st_first = None
                    blk_st_last = None
            else:
                # Actualizar el último status del bloque al día actual
                # Esto asegura que exit time/date se calculen con el turno final
                blk_st_last = st

            current += timedelta(days=1)

        # Cerrar último bloque
        if blk_start is not None:
            blocks.append((blk_start, zone_end, blk_st_first, blk_st_last))
         # ---------------------------------------------------------------
        # 5. CREAR una operación por cada bloque
        # ---------------------------------------------------------------
        for (b_start, b_end, b_first_st, b_last_st) in blocks:
            # --- FIX BUG 3: Inyectar hora manual si este bloque coincide ---
            current_manual_exit = None
            if manual_split_config:
                target_date = manual_split_config.get('date')
                target_time = manual_split_config.get('time')
                if target_date and b_end == target_date:
                    current_manual_exit = target_time
                    print(f"[CONSOL] Injecting manual exit time {target_time} for block ending {b_end}")

            self._create_single_operation(
                badge, role, username, b_start, b_end, b_first_st, b_last_st,
                _saved_exits, cursor,
                manual_exit_time=current_manual_exit
            )

    def _create_single_operation(self, badge, role, username, start, end, first_status, last_status=None, saved_exits=None, cursor=None, manual_exit_time=None):
        """
        Crea UNA operación para un bloque contiguo.

        Replica el comportamiento del motor OLD:
        - Entry time se calcula con el status del PRIMER día (first_status)
        - Exit time/date se calcula con el status del ÚLTIMO día (last_status)

        Esto permite que bloques mixtos (ON + ON NS) se unan correctamente
        cuando el usuario elige "Unir (Mismo Viaje)".
        """
        if last_status is None:
            last_status = first_status

        t_in = self._calculate_time_logic(first_status, "IN")
        t_out = self._calculate_time_logic(last_status, "OUT")
        real_exit_date = self._calculate_rgm_exit_date(last_status, end)

        # Verificar si mañana es force_start → suprimir +1D
        next_day = end + timedelta(days=1)
        next_map = db.get_schedule_map_for_range(badge, next_day, next_day, self.source, cursor=cursor)
        if next_map.get(next_day.isoformat(), {}).get("force_new_entry", 0) == 1:
            real_exit_date = end

        computed_exit_dt = datetime.combine(real_exit_date, t_out)

        # --- FIX BUG 3: PRIORIDAD 1 — Hora manual directa del diálogo ---
        if manual_exit_time is not None:
            computed_exit_dt = datetime.combine(end, manual_exit_time)
            print(f"[CONSOL] Applied MANUAL exit time: {computed_exit_dt} for {badge} {start}-{end}")

        # --- PRIORIDAD 2 — Restaurar exit_date override de BD (saved_exits) ---
        elif saved_exits:
            key = (start.isoformat(), end.isoformat())
            old_exit_str = saved_exits.get(key)
            if old_exit_str:
                try:
                    old_exit_dt = datetime.strptime(old_exit_str, '%Y-%m-%d %H:%M')
                    # Solo restaurar si la hora es distinta (= era un override manual)
                    if old_exit_dt != computed_exit_dt:
                        computed_exit_dt = old_exit_dt
                        print(f"[CONSOL] Restored exit override: {old_exit_str} for {badge} {start}-{end}")
                except (ValueError, TypeError):
                    pass

        db.add_operation(
            username=username,
            role=role,
            badge=badge,
            start_date=start,
            end_date=end,
            created_by=self.logged_username,
            entry_date=datetime.combine(start, t_in),
            exit_date=computed_exit_dt,
            cursor=cursor
        )


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
        macro_started = False  # agrupa todos los empleados del mismo drag en 1 undo-step

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
                        #self._apply_base_background(item, new_status)
                        self._apply_status_background(item, new_status)
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

                        # ── UNDO: capturar estado ANTES del drag (incluye bordes) ──
                        snapshot_start = operation_start_date
                        snapshot_end   = end_fill_date + timedelta(days=1)
                        _drag_old_map  = db.get_schedule_map_for_range(
                            badge, snapshot_start, snapshot_end, self.source
                        )

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
                        # Recuperar config de _resolve_force_new_entry_start (si hubo Separar)
                        start_manual_cfg = getattr(self, '_last_manual_exit_config', None)
                        self._last_manual_exit_config = None  # Consumir

                        self._consolidate_and_record_logistics(
                            badge,
                            role,
                            username,
                            operation_start_date,
                            end_fill_date,
                            new_status,
                            manual_split_config=start_manual_cfg
                        )

                         # 4. RIGHT BOUNDARY CHECK — Borde final del rango arrastrado
                        end_result = self._resolve_force_new_entry_end(
                            badge, end_fill_date, new_status
                        )
                        if end_result == 1:
                            # Recuperar config manual del diálogo (si existe)
                            manual_cfg = getattr(self, '_last_manual_exit_config', None)
                            self._last_manual_exit_config = None  # Consumir

                            # Re-consolidar para recoger el nuevo force_new_entry
                            self._consolidate_and_record_logistics(
                                badge, role, username,
                                end_fill_date, end_fill_date, new_status,
                                manual_split_config=manual_cfg
                            )

                        # ── UNDO: registrar drag como un paso (DB+Excel+UI) ──────
                        if not self._undo_in_progress:
                            _drag_new_map = db.get_schedule_map_for_range(
                                badge, snapshot_start, snapshot_end, self.source
                            )
                            if not macro_started:
                                self._undo_stack.beginMacro("Drag fill")
                                macro_started = True
                            self._undo_stack.push(
                                UndoScheduleSnapshotCommand(
                                    widget=self,
                                    badge=badge,
                                    role=role,
                                    username=username,
                                    snapshot_start=snapshot_start,
                                    snapshot_end=snapshot_end,
                                    op_start=operation_start_date,
                                    op_end=end_fill_date,
                                    old_schedule_map=_drag_old_map,
                                    new_schedule_map=_drag_new_map,
                                    description=f"Drag fill {badge} {start_fill_date}..{end_fill_date}",
                                )
                            )

                    except Exception as e:
                        print(f"Error saving drag-fill for {badge}: {e}")

        finally:
            if macro_started:
                try:
                    self._undo_stack.endMacro()
                except Exception:
                    pass
            self._is_internal_update = False
            self._bulk_editing = False
            self.rotation_changed.emit()
            self.schedule_table.viewport().update()



     # ---------- SAFE REVERT HELPER (protects against zombie QTableWidgetItem) ----------
    def _safe_revert_item(self, item, text):
        """
        Safely revert a QTableWidgetItem's text and background.
        Returns True if the item was still alive, False if it was already destroyed.

        WHY: During modal dialogs (editor.exec(), QMessageBox, _resolve_force_new_entry_*),
        the Qt event loop keeps processing events. If any signal triggers refresh_ui_data(),
        the table is cleared/rebuilt — destroying the C++ object behind 'item'.
        Accessing it after that causes: RuntimeError: wrapped C/C++ object has been deleted.
        """
        try:
            if item is None or item.row() < 0 or item.column() < 0:
                return False
            item.setText(text)
            self._apply_base_background(item, text)
            return True
        except RuntimeError:
            print(f"\u26a0\ufe0f [Recovery] QTableWidgetItem destroyed by concurrent UI refresh. Revert skipped safely.")
            return False




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
                        self._safe_revert_item(item, old_text)
                    return
                else:
                    # Mantener el nuevo valor, marcar warn suave
                    with QSignalBlocker(self.schedule_table):
                        try:
                            item.setText(new_text)  # normalizar
                            item.setBackground(QColor(WARN_BG_HEX))
                            self._warn_highlight_keys.add(self._warn_key_for(r, c))
                        except RuntimeError:
                            pass  # Item destroyed by concurrent refresh
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
                    self._safe_revert_item(item, old_text)
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
                    self._safe_revert_item(item, old_text)
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

            # ── CAPTURAR ESTADO ANTIGUO para Undo Stack (antes de la escritura) ──
            # Necesitamos el rango completo (puede ser >1 día si apply_to_range).
            _undo_old_map = db.get_schedule_map_for_range(
                badge, start_date_block, end_date_block, self.source
            )

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
                        self._safe_revert_item(item, old_text)
                    return
                # ---------------------------------------------------------
                # 1. Guardar Schedule (SSoT) con force_new_entry_start
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

                # [RIGHT BOUNDARY CHECK] " Borde final del bloque editado
                end_result = self._resolve_force_new_entry_end(
                    badge, end_date_block, schedule_status
                )
                if end_result is None:
                    # Usuario canceló → revertir
                    with QSignalBlocker(self.schedule_table):
                        self._safe_revert_item(item, old_text)
                    return
                # 2. SIEMPRE consolidar (maneja JOIN y SEPARATE correctamente)
                #    FIX: Antes solo se llamaba cuando end_result==1 (SEPARATE),
                #    lo que dejaba "Unir (Mismo Viaje)" sin efecto real.

                # Recuperar config manual del diálogo (si existe)
                manual_cfg = getattr(self, '_last_manual_exit_config', None)
                self._last_manual_exit_config = None  # Consumir

                self._consolidate_and_record_logistics(
                    badge, role, username,
                    start_date_block, end_date_block, schedule_status,
                    manual_split_config=manual_cfg
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

                # ── REGISTRAR EN UNDO STACK (sólo si NO estamos en undo/redo) ──
                if not self._undo_in_progress:
                    _undo_cmd = UndoCellChangeCommand(
                        widget=self,
                        badge=badge,
                        role=role,
                        username=username,
                        start_date=start_date_block,
                        end_date=end_date_block,
                        old_schedule_map=_undo_old_map,
                        old_pickup=pickup_init,
                        old_dropoff=dropoff_init,
                        new_status=schedule_status or "",
                        new_shift_type=shift_type,
                        new_in_time=in_time,
                        new_out_time=out_time,
                        new_remark=remark,
                        new_pickup=pickup,
                        new_dropoff=dropoff,
                        description=f"Edit {badge} {start_date_block}",
                    )
                    self._undo_stack.push(_undo_cmd)

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

        # --- LOGICA DE TITULOS Y HORARIOS (data-driven via DB) ---
        # Buscar el type en _custom_shift_map (incluye ON/ON NS/OFF si estan en DB)
        _type_info = (self._custom_shift_map or {}).get(status_code)

        if _type_info:
            shift_title = _type_info.get("name", status_code)
            is_custom_off = bool(_type_info.get("is_off", 0))

            if is_custom_off:
                # Tipo OFF (system u otro): no mostrar horarios
                in_time = None
                out_time = None
            else:
                # Tipo working: usar tiempos de DB si no vienen de la celda
                if not in_time:
                    in_time = _type_info.get("in_time")
                if not out_time:
                    out_time = _type_info.get("out_time")

        elif status_code == "OFF":
            # Fallback hardcode para OFF si no esta en DB
            shift_title = "OFF"
            in_time = None
            out_time = None
            is_custom_off = True

        else:
            # Status desconocido: mostrar tal cual
            shift_title = status_code
            is_custom_off = False

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

        db.delete_operations_in_range(badge, start_date, end_date)
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
    # HELPER: Detección de Colisiones (Working -> Working con CAMBIO DE TURNO)
    # --------------------------------------------------------------------------
    # --------------------------------------------------------------------------
    # HELPER: Detección de Colisiones con UI Avanzada
    # --------------------------------------------------------------------------
    def _resolve_force_new_entry_start(self, badge, start_date, new_status, cursor=None):
        """
        Detecta colisiones (Working -> Working) y gestiona la ruptura de viajes.
        Si el usuario elige SEPARAR, actualiza inmediatamente la hora de salida del viaje ANTERIOR.

        Args:
          cursor: Si se pasa (ej: desde paste), las lecturas usan ESA misma
                  conexión/transacción para ver datos recién escritos.
        Retorna:
          1 -> Separar (force_new_entry_start=1 para hoy)
          0 -> Unir (force_new_entry_start=0 para hoy)
          None -> Cancelar operación
        """
        # 1. Validaciones básicas (sin cambios)
        n_code = str(new_status or "").strip().upper()
        if not n_code or not db.is_working_status(n_code, self.source):
            return 0

        prev_day = start_date - timedelta(days=1)
        prev_map = db.get_schedule_map_for_range(badge, prev_day, prev_day, self.source, cursor=cursor)
        prev_info = prev_map.get(prev_day.isoformat(), {})
        p_code = str(prev_info.get("status") or "").strip().upper()

        if not p_code or not db.is_working_status(p_code, self.source):
            return 0

        # Si son el mismo código (ej: ON -> ON), asumimos continuidad automática
        if p_code == n_code:
            return 0

        # ---------------------------------------------------------
        # 2. REGLA DE NEGOCIO: ¿Mostrar editor de hora de salida?
        #    SOLO para RGM + turno anterior aplica regla 1+D.
        #    Cubre ON, ON NS y cualquier Shift Type custom con apply_1d=True.
        # ---------------------------------------------------------
        ask_prev_exit_time = (
            self.source == "RGM"
            and self._shift_applies_1d(p_code)
        )

        # ---------------------------------------------------------
        # 3. CONSTRUCCIÓN DEL DIÁLOGO PERSONALIZADO
        # ---------------------------------------------------------
        dialog = QDialog(self)
        dialog.setWindowTitle("Gestión de Turnos Consecutivos")
        dialog.setWindowFlags(dialog.windowFlags() & ~Qt.WindowType.WindowContextHelpButtonHint)
        dialog.setFixedWidth(420)

        layout = QVBoxLayout(dialog)

        # A) Panel de Información (Alerta Visual)
        info_frame = QFrame()
        info_frame.setStyleSheet("background-color: #FFF3E0; border: 1px solid #FFE0B2; border-radius: 4px; padding: 6px;")
        info_layout = QVBoxLayout(info_frame)

        lbl_info = QLabel(
            f"<h3 style='color:#E65100; margin:0;'>⚠️ Cambio de Turno Detectado</h3>"
            f"<div style='margin-top:5px;'>"
            f"El usuario <b>{badge}</b> tiene turnos consecutivos diferentes:<br>"
            f"• Ayer ({prev_day.strftime('%d/%m')}): <b>{p_code}</b><br>"
            f"• Hoy ({start_date.strftime('%d/%m')}): <b>{n_code}</b>"
            f"</div>"
        )
        lbl_info.setTextFormat(Qt.TextFormat.RichText)
        info_layout.addWidget(lbl_info)
        layout.addWidget(info_frame)

        layout.addSpacing(10)

        # B) Grupo de hora de salida — SOLO si aplica (RGM + ON/ON NS)
        time_edit = None
        if ask_prev_exit_time:
            grp_split = QGroupBox("Si elige 'Separar (Nuevo Viaje)':")
            grp_split.setStyleSheet("QGroupBox { font-weight: bold; color: #374151; }")
            grp_layout = QVBoxLayout(grp_split)

            grp_layout.addWidget(QLabel(f"Defina la hora real de SALIDA del viaje de ayer ({prev_day.strftime('%d/%m')}):"))

            # Heurística: Si ayer fue turno de noche (NS), sugerir 06:00, sino 18:00
            default_hour = 6 if ("NS" in p_code or "NIGHT" in p_code) else 18
            time_edit = QTimeEdit(QTime(default_hour, 0))
            time_edit.setDisplayFormat("HH:mm")
            time_edit.setStyleSheet("font-size: 13px; padding: 4px;")
            grp_layout.addWidget(time_edit)

            layout.addWidget(grp_split)
            layout.addSpacing(10)

        # C) Botonera
        btn_box = QDialogButtonBox()

        # Botón UNIR (Primary)
        btn_join = btn_box.addButton("Unir (Mismo Viaje)", QDialogButtonBox.ButtonRole.AcceptRole)
        btn_join.setStyleSheet("padding: 6px 15px;")

        # Botón SEPARAR (Action/Destructive) - Color Rojo para resaltar ruptura
        btn_split = btn_box.addButton("Separar (Nuevo Viaje)", QDialogButtonBox.ButtonRole.ActionRole)
        btn_split.setStyleSheet("background-color: #D32F2F; color: white; padding: 6px 15px; font-weight: bold;")

        # Botón CANCELAR
        btn_cancel = btn_box.addButton("Cancelar", QDialogButtonBox.ButtonRole.RejectRole)

        layout.addWidget(btn_box)

        # Conectar señales (Usamos códigos de salida: 10=Join, 20=Split)
        btn_join.clicked.connect(lambda: dialog.done(10))
        btn_split.clicked.connect(lambda: dialog.done(20))
        btn_cancel.clicked.connect(dialog.reject)

        # ---------------------------------------------------------
        # 3. EJECUCIÓN Y PROCESAMIENTO
        # ---------------------------------------------------------
        result = dialog.exec()

        if result == 10: # UNIR
            print("[SCD] User chose: JOIN (0)")
            return 0

        elif result == 20: # SEPARAR
            print("[SCD] User chose: SEPARATE (1)")

            # Actualizar hora de salida del viaje anterior SOLO si se mostró el editor
            if ask_prev_exit_time and time_edit is not None:
                try:
                    custom_exit_time = time_edit.time().toPyTime()
                    success, msg = db.update_operation_exit_time_by_date(badge, prev_day, custom_exit_time, cursor=cursor)

                    if success:
                        print(f"[SCD] Previous exit time updated: {msg}")
                    else:
                        print(f"[SCD] Failed to update previous exit time: {msg}")
                        QMessageBox.warning(self, "Database Warning", f"Could not update previous trip exit time:\n{msg}")

                    # Almacenar config para que el consolidador use la hora manual
                    self._last_manual_exit_config = {
                        'date': prev_day,
                        'time': custom_exit_time
                    }
                except Exception as e:
                    print(f"[SCD] Error processing split: {e}")
                    import traceback; traceback.print_exc()
            else:
                print(f"[SCD] Skip exit time update (source={self.source}, prev={p_code})")
                self._last_manual_exit_config = None

            return 1 # Force split flag

        else: # Cancelar o cerrar ventana
            return None


    def _resolve_force_new_entry_end(self, badge, end_date, end_status, cursor=None):
        """
        Detecta colisiones en el BORDE DERECHO del bloque pegado/arrastrado.

        Escenario: pegaste ON en 16-17, pero el día 18 ya tiene WH.
        Sin esta validación, el ON aplica 1+D (exit=18 07:00) y se cruza con WH (04:00).

        Retorna:
          1 -> Separar (se marcó force_new_entry en next_day + exit override aplicado)
          0 -> Unir / No hay conflicto
          None -> Cancelar
        """
        e_code = str(end_status or "").strip().upper()
        if not e_code or not db.is_working_status(e_code, self.source):
            return 0

        next_day = end_date + timedelta(days=1)
        next_map = db.get_schedule_map_for_range(badge, next_day, next_day, self.source, cursor=cursor)
        next_info = next_map.get(next_day.isoformat(), {})
        n_code = str(next_info.get("status") or "").strip().upper()

        if not n_code or not db.is_working_status(n_code, self.source):
            return 0

        # Mismo código = continuidad, no hay conflicto
        if e_code == n_code:
            return 0

        # Si mañana YA tiene force_new_entry, no necesitamos preguntar de nuevo
        if next_info.get("force_new_entry", 0) == 1:
            return 0

        # ---------------------------------------------------------
        # REGLA DE NEGOCIO: ¿Mostrar editor de hora de salida?
        # SOLO para RGM + el bloque que TERMINA aplica regla 1+D.
        # Cubre ON, ON NS y cualquier Shift Type custom con apply_1d=True.
        # ---------------------------------------------------------
        ask_exit_time = (
            self.source == "RGM"
            and self._shift_applies_1d(e_code)
        )

        # ---------------------------------------------------------
        # DIÁLOGO PERSONALIZADO (Borde Derecho)
        # ---------------------------------------------------------
        dialog = QDialog(self)
        dialog.setWindowTitle("Gestión de Turnos Consecutivos (Borde Final)")
        dialog.setWindowFlags(dialog.windowFlags() & ~Qt.WindowType.WindowContextHelpButtonHint)
        dialog.setFixedWidth(440)

        layout = QVBoxLayout(dialog)

        # A) Panel de Información
        info_frame = QFrame()
        info_frame.setStyleSheet(
            "background-color: #FFF3E0; border: 1px solid #FFE0B2; "
            "border-radius: 4px; padding: 6px;"
        )
        info_layout = QVBoxLayout(info_frame)

        lbl_info = QLabel(
            f"<h3 style='color:#E65100; margin:0;'>⚠️ Cambio de Turno Detectado (Borde Final)</h3>"
            f"<div style='margin-top:5px;'>"
            f"El usuario <b>{badge}</b> tiene turnos consecutivos diferentes:<br>"
            f"• Último día marcado ({end_date.strftime('%d/%m')}): <b>{e_code}</b><br>"
            f"• Día siguiente ({next_day.strftime('%d/%m')}): <b>{n_code}</b>"
            f"</div>"
        )
        lbl_info.setTextFormat(Qt.TextFormat.RichText)
        info_layout.addWidget(lbl_info)
        layout.addWidget(info_frame)

        # Nota sobre 1+D
        if ask_exit_time:
            warn_lbl = QLabel(
                f"<div style='color:#C62828; margin-top:6px;'>"
                f"⛔ El turno <b>{e_code}</b> aplica la regla 1+D: su salida estándar sería "
                f"<b>{next_day.strftime('%d/%m')} 07:00</b>, lo que se cruza con el turno "
                f"<b>{n_code}</b> del mismo día."
                f"</div>"
            )
            warn_lbl.setTextFormat(Qt.TextFormat.RichText)
            warn_lbl.setWordWrap(True)
            layout.addWidget(warn_lbl)

            # FIX BUG 5: Label dinámico que refleja la hora seleccionada
            selected_preview_lbl = QLabel()
            selected_preview_lbl.setStyleSheet(
                "color: #1B5E20; font-weight: bold; font-size: 12px; margin-top: 4px;"
            )
            layout.addWidget(selected_preview_lbl)

        layout.addSpacing(10)

        # B) Grupo de hora de salida (solo RGM + ON/ON NS)
        time_edit = None
        if ask_exit_time:
            grp_split = QGroupBox("Si elige 'Separar (Nuevo Viaje)':")
            grp_split.setStyleSheet("QGroupBox { font-weight: bold; color: #374151; }")
            grp_layout = QVBoxLayout(grp_split)

            grp_layout.addWidget(QLabel(
                f"Defina la hora real de SALIDA del bloque {e_code} "
                f"(último día: {end_date.strftime('%d/%m')}):"
            ))

            default_hour = 6 if ("NS" in e_code or "NIGHT" in e_code) else 18
            time_edit = QTimeEdit(QTime(default_hour, 0))
            time_edit.setDisplayFormat("HH:mm")
            time_edit.setStyleSheet("font-size: 13px; padding: 4px;")
            grp_layout.addWidget(time_edit)

            # FIX BUG 5: Conectar cambio de hora a actualización visual
            def _update_preview():
                if selected_preview_lbl is not None:
                    chosen = time_edit.time().toPyTime()
                    chosen_dt = datetime.combine(end_date, chosen)
                    selected_preview_lbl.setText(
                        f"✅ Salida real seleccionada: {chosen_dt.strftime('%d/%m %H:%M')}"
                    )

            time_edit.timeChanged.connect(lambda _: _update_preview())
            _update_preview()  # Mostrar valor inicial

            layout.addWidget(grp_split)
            layout.addSpacing(10)

        # C) Botonera
        btn_box = QDialogButtonBox()
        btn_join = btn_box.addButton("Unir (Mismo Viaje)", QDialogButtonBox.ButtonRole.AcceptRole)
        btn_join.setStyleSheet("padding: 6px 15px;")
        btn_split = btn_box.addButton("Separar (Nuevo Viaje)", QDialogButtonBox.ButtonRole.ActionRole)
        btn_split.setStyleSheet(
            "background-color: #D32F2F; color: white; padding: 6px 15px; font-weight: bold;"
        )
        btn_cancel = btn_box.addButton("Cancelar", QDialogButtonBox.ButtonRole.RejectRole)

        layout.addWidget(btn_box)

        btn_join.clicked.connect(lambda: dialog.done(10))
        btn_split.clicked.connect(lambda: dialog.done(20))
        btn_cancel.clicked.connect(dialog.reject)

        # ---------------------------------------------------------
        # EJECUCIÓN Y PROCESAMIENTO
        # ---------------------------------------------------------
        result = dialog.exec()

        if result == 10:  # UNIR
            print(f"[SCD-END] User chose: JOIN for {badge} {end_date}->{next_day}")
            return 0

        elif result == 20:  # SEPARAR
            print(f"[SCD-END] User chose: SEPARATE for {badge} {end_date}->{next_day}")

            # --- FIX BUG 1: Garantizar que siempre hay un cursor válido ---
            own_conn = None
            active_cursor = cursor
            if active_cursor is None:
                own_conn = db.sqlite3.connect(db.DB_FILE)
                active_cursor = own_conn.cursor()

            try:
                # 1. Marcar force_new_entry=1 en el día siguiente
                active_cursor.execute(
                    "UPDATE schedules SET force_new_entry = 1 "
                    "WHERE badge = ? AND date = ? AND source = ?",
                    (badge, next_day.isoformat(), self.source)
                )
                rows_updated = active_cursor.rowcount
                if rows_updated == 0:
                    print(f"[SCD-END] WARNING: No schedule row found for {badge} on {next_day}")
                else:
                    print(f"[SCD-END]  force_new_entry=1 set on {next_day}")

                # 2. Si aplica, actualizar exit time del bloque que termina
                if ask_exit_time and time_edit is not None:
                    custom_exit_time = time_edit.time().toPyTime()
                    success, msg = db.update_operation_exit_time_by_date(
                        badge, end_date, custom_exit_time, cursor=active_cursor
                    )
                    if success:
                        print(f"[SCD-END]  Exit time updated: {msg}")
                    else:
                        print(f"[SCD-END]  Failed: {msg}")

                # Commit solo si abrimos conexión propia
                if own_conn:
                    own_conn.commit()

            except Exception as e:
                print(f"[SCD-END] CRITICAL Error in SEPARATE: {e}")
                import traceback; traceback.print_exc()
                if own_conn:
                    own_conn.rollback()
            finally:
                if own_conn:
                    own_conn.close()

            # --- Capturar la hora elegida para pasarla al caller ---
            self._last_manual_exit_config = None
            if ask_exit_time and time_edit is not None:
                chosen_time = time_edit.time().toPyTime()
                self._last_manual_exit_config = {
                    'date': end_date,
                    'time': chosen_time
                }

            return 1

        else:  # Cancelar
            return None

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
            # OPCIÓN B: Automático (Check desmarcado) -> Calcular Defaults
            # Se usa la fecha de inicio del periodo para Entry y fin para Exit.
            # Se inyectan las horas según el tipo de turno o reglas de negocio.

            # Hora por defecto base (07:00 / 07:00) Para RGM
            t_in = datetime.strptime("07:00", "%H:%M").time()
            t_out = datetime.strptime("07:00", "%H:%M").time()

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

                if is_newmont:
                    if schedule_status == "ON": # Día Newmont
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
            # Si es RGM y es ON/ON NS -> Salida es Mañana (end_date + 1)
            # Si es RGM y es Otro (Capacitacion) -> Salida es Hoy (end_date)
            # El método _calculate_rgm_exit_date encapsula esta lógica para no repetir if/else aquí.
            real_exit_date = self._calculate_rgm_exit_date(schedule_status, end_date)

            # 2. Combinamos con las horas (t_in / t_out) que ya calculaste arriba
            entry_datetime = datetime.combine(start_date, t_in)
            exit_datetime = datetime.combine(real_exit_date, t_out) # Usamos real_exit_date

            # --- FIN DEL CAMBIO 1+D ---

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

        # ----------------------------------------------------------
        # 6.5) Shift Collision Detector (Working-Working)
        # ----------------------------------------------------------
        force_new_entry_flag = 0

        try:
            # Debug log to trace the collision check
            print(f"[SCD] Checking collision: schedule_status={schedule_status!r}, badge={badge!r}, start_date={start_date}")

            # Only check for collision if the NEW status is a working status.
            # (Non-working statuses like 'OFF' usually imply a simple overwrite or update, handled elsewhere)
            if schedule_status and db.is_working_status(schedule_status, self.source):

                #  ARCHITECTURAL FIX: Use the unified resolution logic.
                # prevent duplication of logic by reusing _resolve_force_new_entry_start.
                # This method handles the UI dialog AND the database update for the PREVIOUS shift.
                force_result = self._resolve_force_new_entry_start(badge, start_date, schedule_status)

                if force_result is None:
                    # The user clicked "Cancel" or closed the dialog.
                    # We must abort the entire save operation to prevent data corruption.
                    print("[SCD] User cancelled the collision resolution. Aborting save.")
                    return

                force_new_entry_flag = force_result
                print(f"[SCD] Collision resolved. Force Flag: {force_new_entry_flag}")

            else:
                print(f"[SCD] Skipped: '{schedule_status}' is not a working status or is invalid.")

        except Exception as e:
            # Fallback safety: If the detector crashes, assume NO force (Standard Overwrite/Error)
            # but log the stack trace critical for debugging.
            import traceback as _tb
            print(f"[SCD] *** CRITICAL EXCEPTION in Shift Collision Detector: {e}")
            _tb.print_exc()
            force_new_entry_flag = 0

        # ----------------------------------------------------------
        # 6.6) Post-Resolution Date Correction
        # ----------------------------------------------------------
        # If the user chose "Separar" (Force New), the new shift implies a break in continuity.
        # Often, a "Separated" shift should end on the SAME DAY, not D+1, unless specified.
        # This logic ensures the NEW trip's exit_datetime is calculated correctly.

        if force_new_entry_flag == 1 and exit_datetime:
            # Recalculate the end date for the NEW entry specifically for a separated trip.
            # separated_trip=True usually forces the date to be start_date (removes D+1 logic).
            real_exit_date = self._calculate_rgm_exit_date(
                schedule_status, end_date, separated_trip=True
            )
            # Combine the new date with the originally selected time
            exit_datetime = datetime.combine(real_exit_date, exit_datetime.time())
            print(f"[LOGIC] Exit Date Recalculated due to SEPARATION: {exit_datetime}")

        # 7. Guardar en BD (SSoT)
        if schedule_status is not None:
            # Aquí es donde se guardan los datetimes calculados (entry_datetime/exit_datetime)
            db.delete_operations_in_range(badge, start_date, end_date)
            db.add_operation(
                username=username,
                role=role,
                badge=badge,
                start_date=start_date,
                end_date=end_date,
                created_by=self.logged_username,
                entry_date=entry_datetime,  # Ahora siempre tendrá valor si hay turno
                exit_date=exit_datetime,  # Ahora siempre tendrá valor si hay turno
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
               # [RIGHT BOUNDARY CHECK] — Borde final del rango registrado
            end_result = self._resolve_force_new_entry_end(
                badge, end_date, schedule_status
            )
            if end_result is None:
                print("[SCD-END] User cancelled right boundary. Aborting.")
                return
            if end_result == 1:
                manual_cfg = getattr(self, '_last_manual_exit_config', None)
                self._last_manual_exit_config = None

                self._consolidate_and_record_logistics(
                    badge, role, username,
                    start_date, end_date, schedule_status,
                    manual_split_config=manual_cfg
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
            # Si no hay selección (sel is None), asumimos que no se marca (OFF-like)
            is_off_day_guard = True

        if is_off_day_guard:
            pickup = None
            dropoff = None


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
        QTimer.singleShot(2000, lambda: setattr(self, "_is_internal_update", False))

        # 9. Auditoría y Finalización
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

    # ---------- Excel Health / Monitoring ----------
    def check_excel_health(self):
        # This entire block is wrapped in a try/except to prevent a crash
        # if the timer fires after the widget has been destroyed (e.g., on close).
        try:
            exists = os.path.exists(self.excel_file)
            if not exists:
                self.excel_health_label.setText(
                    "Excel status:  Not found (it may have been moved, deleted, or renamed)."
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
                #  Show the signed-in site (RGM/Newmont), not the structural variant
                self.excel_health_label.setText(
                    f"Excel status:  OK ({self.source}) — {os.path.basename(self.excel_file)}"
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
                last_pasted_per_badge = {}  # badge -> (date, status, role, username)
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
                        force_flag = self._resolve_force_new_entry_start(badge, target_date, raw_text, cursor=cursor)

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
                            cursor=cursor,
                            manual_split_config=getattr(self, '_last_manual_exit_config', None)
                        )
                        self._last_manual_exit_config = None  # Consumir

                        # --- Actualización Visual ---
                        item = self.schedule_table.item(abs_row, abs_col)
                        if not item:
                            item = QTableWidgetItem()
                            self.schedule_table.setItem(abs_row, abs_col, item)
                        item.setText(raw_text)
                        #self._apply_base_background(item, raw_text)
                        self._apply_status_background(item, raw_text)
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

                        # Track last pasted per badge (for RIGHT boundary check)
                        prev_entry = last_pasted_per_badge.get(badge)
                        if prev_entry is None or target_date > prev_entry[0]:
                            last_pasted_per_badge[badge] = (target_date, raw_text, role, username)

                # =============================================================
                # RIGHT BOUNDARY CHECK — Borde Final de cada badge pegado
                # =============================================================
                for b_badge, (b_end_date, b_end_status, b_role, b_username) in last_pasted_per_badge.items():
                    end_result = self._resolve_force_new_entry_end(
                        b_badge, b_end_date, b_end_status, cursor=cursor
                    )
                    if end_result is None:
                        # Usuario canceló → rollback
                        cursor.execute("ROLLBACK")
                        return
                    if end_result == 1:
                        # Recuperar config manual del diálogo
                        manual_cfg = getattr(self, '_last_manual_exit_config', None)
                        self._last_manual_exit_config = None  # Consumir

                        self._consolidate_and_record_logistics(
                            b_badge, b_role, b_username,
                            b_end_date, b_end_date, b_end_status,
                            cursor=cursor,
                            manual_split_config=manual_cfg
                        )


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

    def _shift_applies_1d(self, shift_code: str) -> bool:
        """
        Retorna True si el shift debe usar la regla 1+D (salida = end_date + 1).
        Solo aplica para RGM.

        PRIORIDAD:
        1. DB: shift_types.apply_1d (ON/ON NS son system types con apply_1d=1 en RGM)
        2. Hardcode legacy (fallback si DB vacia / seed no ejecutado)
        """
        if self.source != "RGM":
            return False

        code_up = (shift_code or "").strip().upper()

        # PRIORIDAD 1: Consultar DB
        meta = (self._custom_shift_map or {}).get(code_up)
        if meta is not None:
            return bool(meta.get("apply_1d", 0))

        # PRIORIDAD 2: Fallback hardcode (solo si ON/ON NS no estan en DB)
        if code_up in ("ON", "ON NS"):
            return True

        return False
    def _calculate_rgm_exit_date(self, shift_code, end_date, separated_trip=False):
        """
        RGM exit date logic:
        - ON/ON NS or any custom Shift Type with apply_1d=True -> Next Day (1+D rule).
        - Separated trip -> Same Day (trip ends here, no +1).
        - All other shifts / Newmont -> Same Day.
        """
        if self._shift_applies_1d(shift_code):
            if separated_trip:
                return end_date
            from datetime import timedelta
            return end_date + timedelta(days=1)
        return end_date
