# -*- coding: utf-8 -*-
# Ventanas principales con SSoT, validación, regeneración y detección de archivo.
# Integra la UI con database_logic.py y excel_logic.py para:
# - DB como SSoT
# - Validación de estructura de Excel antes de importar
# - Regeneración completa del PlanStaff (RGM o Newmont) desde BD
# - Detección de movimiento/eliminación/renombrado del Excel con QTimer

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel,
    QLineEdit, QComboBox, QDateEdit, QPushButton, QTableWidget,
    QTableWidgetItem, QHeaderView, QGroupBox, QMessageBox, QFileDialog,
    QTabWidget, QApplication, QColorDialog, QTimeEdit
)
from PyQt6.QtCore import QDate, Qt, pyqtSignal, QTime, QTimer
from PyQt6.QtGui import QColor
from datetime import datetime
import os

# App logic
import database_logic as db
import excel_logic as excel


# -------------------------------------------------------------
# Utilidad común
# -------------------------------------------------------------
def create_group_box(title: str, inner_layout) -> QGroupBox:
    box = QGroupBox(title)
    font = box.font()
    font.setBold(True)
    box.setFont(font)
    box.setLayout(inner_layout)
    return box


def _clean(value) -> str:
    """
    FR-07: limpia celdas para la UI: None/NaN/'nan'/'null' -> ''
    """
    if value is None:
        return ""
    s = str(value).strip()
    if s.lower() in ("nan", "none", "null"):
        return ""
    return s


# -------------------------------------------------------------
# Widget: Plan Staff (Preview, Registro, Historial, Reportes)
# -------------------------------------------------------------
class PlanStaffWidget(QWidget):
    def __init__(self, source: str, excel_file: str, logged_username: str):
        super().__init__()
        self.source = source              # "RGM" | "Newmont"
        self.excel_file = excel_file
        self.logged_username = logged_username or "Unknown"
        self._last_excel_mtime = None
        self._missing_prompt_shown = False

        root = QVBoxLayout(self)

        # --- Estado del Excel (salud / SSoT) ---
        status_layout = QHBoxLayout()
        self.excel_health_label = QLabel("Excel status: checking...")
        self.excel_health_label.setStyleSheet("font-weight: bold;")
        self.regen_button = QPushButton("🛠️ Regenerar PlanStaff desde BD")
        self.regen_button.clicked.connect(self.regenerate_excel_from_db)
        self.validate_button = QPushButton("🧪 Validar Estructura Excel")
        self.validate_button.clicked.connect(self.validate_excel_structure_ui)
        self.compare_button = QPushButton("🔎 Comparar Excel vs BD")
        self.compare_button.clicked.connect(self.compare_excel_db_ui)
        status_layout.addWidget(self.excel_health_label)
        status_layout.addStretch()
        status_layout.addWidget(self.validate_button)
        status_layout.addWidget(self.compare_button)
        status_layout.addWidget(self.regen_button)

        # ❗️FIX: agregar QGroupBox con addWidget (no addLayout)
        root.addWidget(create_group_box("Estado del Archivo & SSoT", status_layout))

        # --- Schedule preview ---
        preview_container = QVBoxLayout()
        tables_layout = QHBoxLayout()
        tables_layout.setSpacing(0)
        tables_layout.setContentsMargins(0, 0, 0, 0)

        self.frozen_table = QTableWidget()
        self.schedule_table = QTableWidget()
        self.frozen_table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.schedule_table.verticalScrollBar().valueChanged.connect(self.frozen_table.verticalScrollBar().setValue)
        self.frozen_table.verticalScrollBar().valueChanged.connect(self.schedule_table.verticalScrollBar().setValue)

        tables_layout.addWidget(self.frozen_table)
        tables_layout.addWidget(self.schedule_table, 1)
        preview_container.addLayout(tables_layout)
        preview_title = f"🗓️ Schedule Preview ({os.path.basename(self.excel_file)})"
        root.addWidget(create_group_box(preview_title, preview_container))

        # --- Registro + Historial lado a lado ---
        forms_db_layout = QHBoxLayout()
        # Registro
        self.registration_groupbox = create_group_box("1. Register Employee Schedule (DB is SSoT)", self._build_registration_form())
        forms_db_layout.addWidget(self.registration_groupbox, 1)
        # Historial (operations)
        db_view_layout = QVBoxLayout()
        self.db_table = QTableWidget()
        db_view_layout.addWidget(self.db_table)
        self.db_view_groupbox = create_group_box("Rotation History", db_view_layout)
        forms_db_layout.addWidget(self.db_view_groupbox, 2)
        root.addLayout(forms_db_layout)

        # --- Reportes / Export ---
        report_layout = QHBoxLayout()
        report_layout.addWidget(QLabel("START Date:"))
        self.report_start_date = QDateEdit(QDate.currentDate())
        self.report_start_date.setCalendarPopup(True)
        self.report_start_date.setDisplayFormat("dd/MM/yyyy")
        report_layout.addWidget(self.report_start_date)

        report_layout.addWidget(QLabel("END Date:"))
        self.report_end_date = QDateEdit(QDate.currentDate().addDays(30))
        self.report_end_date.setCalendarPopup(True)
        self.report_end_date.setDisplayFormat("dd/MM/yyyy")
        report_layout.addWidget(self.report_end_date)

        report_button = QPushButton("🚀 Generate Report")
        report_button.clicked.connect(self.generate_report)
        report_layout.addWidget(report_button)

        export_button = QPushButton("📤 Export Plan Staff (.xlsx) desde BD")
        export_button.clicked.connect(self.export_plan_from_db)
        report_layout.addWidget(export_button)

        report_title = f"2. Transportation Report & Export (from {os.path.basename(self.excel_file)})"
        root.addWidget(create_group_box(report_title, report_layout))

        # Cargar datos iniciales
        self.refresh_ui_data()

        # --- Monitor del archivo: detecta mover/eliminar/renombrar ---
        self.file_watch_timer = QTimer(self)
        self.file_watch_timer.setInterval(2000)  # 2s
        self.file_watch_timer.timeout.connect(self.check_excel_health)
        self.file_watch_timer.start()
        self.check_excel_health()

    # ---------- sub-UIs ----------
    def _build_registration_form(self):
        form_layout = QGridLayout()

        self.user_selector_combo = QComboBox()
        self.user_selector_combo.currentIndexChanged.connect(self.autofill_user_data)

        self.role_display = QLineEdit(); self.role_display.setReadOnly(True)
        self.badge_display = QLineEdit(); self.badge_display.setReadOnly(True)

        # NUEVO: selector unificado de estado/turno (base + personalizados)
        self.status_selector = QComboBox()

        self.start_date_edit = QDateEdit(QDate.currentDate()); self.start_date_edit.setCalendarPopup(True); self.start_date_edit.setDisplayFormat("dd/MM/yyyy")
        self.end_date_edit = QDateEdit(QDate.currentDate().addDays(14)); self.end_date_edit.setCalendarPopup(True); self.end_date_edit.setDisplayFormat("dd/MM/yyyy")

        save_button = QPushButton("✅ Save Changes to DB & Excel")
        save_button.clicked.connect(self.save_plan_changes)

        form_layout.addWidget(QLabel("Select Employee:"), 0, 0); form_layout.addWidget(self.user_selector_combo, 0, 1)
        form_layout.addWidget(QLabel("Role/Department:"), 1, 0); form_layout.addWidget(self.role_display, 1, 1)
        form_layout.addWidget(QLabel("Badge (ID):"), 2, 0); form_layout.addWidget(self.badge_display, 2, 1)
        form_layout.addWidget(QLabel("Status/Shift:"), 3, 0); form_layout.addWidget(self.status_selector, 3, 1)
        form_layout.addWidget(QLabel("Period Start Date:"), 4, 0); form_layout.addWidget(self.start_date_edit, 4, 1)
        form_layout.addWidget(QLabel("Period End Date:"), 5, 0); form_layout.addWidget(self.end_date_edit, 5, 1)
        form_layout.addWidget(save_button, 6, 0, 1, 2)

        # Cargar opciones (base + personalizados)
        self.load_shift_type_options()

        return form_layout

    # ---------- data loaders ----------
    def load_shift_type_options(self):
        """Carga estados base + tipos personalizados (de DB) en el combo."""
        self.status_selector.blockSignals(True)
        self.status_selector.clear()

        # Opción vacía
        self.status_selector.addItem("— Do Not Mark Days —", {"kind": "none"})

        # Base
        self.status_selector.addItem("OFF", {"kind": "base", "status": "OFF", "shift_type": None, "in_time": None, "out_time": None})
        self.status_selector.addItem("ON (Day Shift)", {"kind": "base", "status": "ON", "shift_type": "Day Shift", "in_time": None, "out_time": None})
        self.status_selector.addItem("ON NS (Night Shift)", {"kind": "base", "status": "ON NS", "shift_type": "Night Shift", "in_time": None, "out_time": None})

        # Personalizados
        types = db.get_shift_types(self.source)
        if types:
            self.status_selector.addItem("—— Custom Shift Types ——", {"kind": "separator"})
            for t in types:
                label = f"{t['name']} [{t['code']}]  {t['in_time']}-{t['out_time']}"
                self.status_selector.addItem(
                    label,
                    {
                        "kind": "custom",
                        "code": t['code'],
                        "name": t['name'],
                        "in_time": t['in_time'],
                        "out_time": t['out_time']
                    }
                )
        self.status_selector.setCurrentIndex(0)
        self.status_selector.blockSignals(False)

    def load_schedule_data(self):
        FROZEN_COLUMN_COUNT = 3
        df = excel.get_schedule_preview(self.excel_file)
        if df.empty:
            self.frozen_table.clear(); self.schedule_table.clear()
            self.frozen_table.setRowCount(0); self.schedule_table.setRowCount(0)
            return

        # mapa de colores por código personalizado
        custom_map = db.get_shift_type_map(self.source)

        actual_frozen_count = min(df.shape[1], FROZEN_COLUMN_COUNT)
        headers = [str(col.strftime('%Y-%m-%d')) if hasattr(col, "strftime") else str(col) for col in df.columns]

        self.frozen_table.setRowCount(df.shape[0])
        self.frozen_table.setColumnCount(actual_frozen_count)
        self.frozen_table.setHorizontalHeaderLabels(headers[:actual_frozen_count])

        self.schedule_table.setRowCount(df.shape[0])
        self.schedule_table.setColumnCount(max(0, df.shape[1] - actual_frozen_count))
        if df.shape[1] - actual_frozen_count > 0:
            self.schedule_table.setHorizontalHeaderLabels(headers[actual_frozen_count:])

        for i, row in df.iterrows():
            for j, val in enumerate(row):
                text = _clean(val)
                item = QTableWidgetItem(text)

                if j < actual_frozen_count:
                    self.frozen_table.setItem(i, j, item)
                else:
                    col_index = j - actual_frozen_count
                    val_str = text.upper().strip()
                    # Colores base
                    if 'ON NS' in val_str or 'NIGHT' in val_str:
                        item.setBackground(QColor("#FFFF99"))
                    elif val_str == 'ON' or 'DAY' in val_str or val_str.isdigit():
                        item.setBackground(QColor("#C6EFCE"))
                    elif val_str in ('OFF', 'BREAK', 'KO', 'LEAVE'):
                        item.setBackground(QColor("#FFC7CE"))
                    else:
                        # ¿código personalizado?
                        if val_str in custom_map and custom_map[val_str].get("color_hex"):
                            item.setBackground(QColor(custom_map[val_str]["color_hex"]))
                    self.schedule_table.setItem(i, col_index, item)

        self.frozen_table.resizeColumnsToContents()
        self.schedule_table.resizeColumnsToContents()

    def load_db_data(self):
        # Nota: tabla 'operations' no contiene 'source', se lista completo
        records = db.get_all_operations()
        headers = ["ID", "Name", "Role", "Badge", "Start Date", "End Date"]
        self.db_table.setRowCount(len(records))
        self.db_table.setColumnCount(len(headers))
        self.db_table.setHorizontalHeaderLabels(headers)

        for row_idx, record in enumerate(records):
            self.db_table.setItem(row_idx, 0, QTableWidgetItem(str(record['id'])))
            self.db_table.setItem(row_idx, 1, QTableWidgetItem(record['username']))
            self.db_table.setItem(row_idx, 2, QTableWidgetItem(record['role']))
            self.db_table.setItem(row_idx, 3, QTableWidgetItem(record['badge']))
            self.db_table.setItem(row_idx, 4, QTableWidgetItem(record['start_date']))
            self.db_table.setItem(row_idx, 5, QTableWidgetItem(record['end_date']))

        self.db_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

    def load_users_to_selector(self):
        self.user_selector_combo.blockSignals(True)
        self.user_selector_combo.clear()
        self.users_for_selector = db.get_all_users(self.source)
        self.user_selector_combo.addItem("-- Select a user --")
        for user in self.users_for_selector:
            self.user_selector_combo.addItem(user['name'])
        self.user_selector_combo.setCurrentIndex(0)
        self.user_selector_combo.blockSignals(False)
        # limpiar campos dependientes
        self.role_display.clear()
        self.badge_display.clear()

    def refresh_users_only(self):
        """Repuebla SOLO el combo de usuarios (para sincronización instantánea)."""
        self.load_users_to_selector()

    def autofill_user_data(self, index):
        if index > 0:
            user = self.users_for_selector[index - 1]
            self.role_display.setText(user['role'])
            self.badge_display.setText(user['badge'])
        else:
            self.role_display.clear()
            self.badge_display.clear()

    # ---------- actions ----------
    def save_plan_changes(self):
        username = self.user_selector_combo.currentText()
        badge = self.badge_display.text()
        role = self.role_display.text()
        start_date = self.start_date_edit.date().toPyDate()
        end_date = self.end_date_edit.date().toPyDate()

        if not username or username == "-- Select a user --":
            box = QMessageBox(self); box.setIcon(QMessageBox.Icon.Warning)
            box.setWindowTitle("Incomplete Data"); box.setText("Please select an employee.")
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole); box.exec(); return

        if not role:
            box = QMessageBox(self); box.setIcon(QMessageBox.Icon.Warning)
            box.setWindowTitle("Incomplete Data"); box.setText("Please select a role/department.")
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole); box.exec(); return

        if start_date > end_date:
            box = QMessageBox(self); box.setIcon(QMessageBox.Icon.Warning)
            box.setWindowTitle("Date Error"); box.setText("Start date cannot be after end date.")
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole); box.exec(); return

        # Interpretar selección de estado/turno
        sel = self.status_selector.currentData()
        if not sel or sel.get("kind") == "none":
            schedule_status = None; shift_type = None; in_time = None; out_time = None
        elif sel["kind"] == "separator":
            schedule_status = None; shift_type = None; in_time = None; out_time = None
        elif sel["kind"] == "base":
            schedule_status = sel["status"]           # 'ON'|'ON NS'|'OFF'
            shift_type = sel["shift_type"]            # 'Day Shift'|'Night Shift'|None
            in_time = None; out_time = None
        else:
            # custom
            schedule_status = sel["code"]             # código
            shift_type = sel["name"]                  # nombre del tipo
            in_time = sel.get("in_time"); out_time = sel.get("out_time")

        # --- FR-01: Confirmación de sobrescritura (Excel + BD) ---
        conflicts_excel = excel.find_conflicts(self.excel_file, username, badge, start_date, end_date)
        conflicts_db_map = db.get_schedule_map_for_range(badge, start_date, end_date, self.source)
        if conflicts_excel or conflicts_db_map:
            box = QMessageBox(self); box.setIcon(QMessageBox.Icon.Warning)
            box.setWindowTitle("Overwrite Shift Confirmation")
            box.setText("Are you sure you want to modify the existing shift?")
            accept_btn = box.addButton("Accept", QMessageBox.ButtonRole.AcceptRole)
            box.addButton("Cancel", QMessageBox.ButtonRole.RejectRole)
            box.exec()
            if box.clickedButton() != accept_btn:
                return  # abortar

        # --- Datos previos para auditoría (FR-04) ---
        prev_map = db.get_schedule_map_for_range(badge, start_date, end_date, self.source)

        # --- Persistencia en BD (SSoT) ---
        if schedule_status in ("ON", "OFF", "ON NS") or (schedule_status and isinstance(schedule_status, str)):
            # histórico de rango
            if schedule_status is not None:
                db.add_operation(username, role, badge, start_date, end_date)
            # estado día a día
            if schedule_status is None:
                db.clear_schedule_range(badge, start_date, end_date, self.source)
            else:
                db.upsert_schedule_range(
                    badge, start_date, end_date, schedule_status, shift_type, self.source,
                    in_time=in_time, out_time=out_time
                )
        else:
            # limpiar estado día a día en BD cuando se elige "Do Not Mark Days"
            db.clear_schedule_range(badge, start_date, end_date, self.source)

        # --- Excel (artefacto derivado; si no existe, se crea)
        success, message = excel.update_plan_staff_excel(
            self.excel_file, username, role, badge,
            schedule_status, shift_type, start_date, end_date, self.source,
            in_time=in_time, out_time=out_time
        )

        # --- Auditoría (FR-04)
        new_map = db.get_schedule_map_for_range(badge, start_date, end_date, self.source)
        db.log_event(
            self.logged_username,
            self.source,
            "SHIFT_MODIFICATION",
            f"{username} ({badge}) {start_date}..{end_date} prev={prev_map} new={new_map}; Excel={'OK' if success else 'ERR'}"
        )

        # --- Mensaje
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Icon.Information if success else QMessageBox.Icon.Warning)
        box.setWindowTitle("Success" if success else "Warning")
        box.setText(message if success else ("Saved to DB. " + message))
        box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
        box.exec()

        # Refrescar preview/combos/tabla
        self.refresh_ui_data()
        self.check_excel_health()

    def generate_report(self):
        s = self.report_start_date.date().toPyDate()
        e = self.report_end_date.date().toPyDate()
        excel_data, message = excel.generate_transport_report(self.excel_file, s, e)

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Transportation Report",
            f"Transport_Request_{s.strftime('%Y%m%d')}_to_{e.strftime('%Y%m%d')}.xlsx",
            "Excel Files (*.xlsx)"
        )

        if not file_path:
            return
        try:
            with open(file_path, 'wb') as f:
                f.write(excel_data)
            box = QMessageBox(self); box.setIcon(QMessageBox.Icon.Information)
            box.setWindowTitle("Success"); box.setText(f"{message}\n\nReport saved to:\n{file_path}")
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole); box.exec()
            db.log_event(self.logged_username, self.source, "DATA_EXPORT", f"TRANSPORT -> {file_path}")
        except Exception as e:
            box = QMessageBox(self); box.setIcon(QMessageBox.Icon.Critical)
            box.setWindowTitle("Save Error"); box.setText(f"Could not save the file.\nError: {e}")
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole); box.exec()

    def export_plan_from_db(self):
        """FR-03: Exporta una planilla desde el estado actual en la BD (incluye tipos personalizados)."""
        users = db.get_all_users(self.source)
        schedules = db.get_schedules_for_source(self.source)

        if not users:
            box = QMessageBox(self); box.setIcon(QMessageBox.Icon.Warning)
            box.setWindowTitle("No Data"); box.setText("There are no users to export.")
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole); box.exec()
            return

        default_name = f"PlanStaff_{self.source}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        dest_path, _ = QFileDialog.getSaveFileName(self, "Save Plan Staff", default_name, "Excel Files (*.xlsx)")
        if not dest_path:
            return

        ok, msg = excel.export_plan_from_db(self.excel_file, users, schedules, dest_path, self.source)
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Icon.Information if ok else QMessageBox.Icon.Critical)
        box.setWindowTitle("Export" if ok else "Error")
        box.setText(msg)
        box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
        box.exec()

        if ok:
            db.log_event(self.logged_username, self.source, "DATA_EXPORT", f"PLAN_EXPORT -> {dest_path}")

    def refresh_ui_data(self):
        self.load_shift_type_options()
        self.load_schedule_data()
        self.load_db_data()
        self.load_users_to_selector()

    # ---------- NUEVO: Salud/Monitoreo del Excel  ----------
    def check_excel_health(self):
        exists = os.path.exists(self.excel_file)
        if not exists:
            self.excel_health_label.setText("Excel status: ❌ No encontrado (puede haber sido movido, eliminado o renombrado).")
            self.excel_health_label.setStyleSheet("color: #B00020; font-weight: bold;")
            if not self._missing_prompt_shown:
                self._missing_prompt_shown = True
                self.prompt_regenerate_or_locate()
            return

        # Existe -> validar estructura y detectar cambios
        try:
            mtime = os.path.getmtime(self.excel_file)
            structure_ok, errors, meta = excel.validate_excel_structure(self.excel_file)
            if structure_ok:
                self.excel_health_label.setText(f"Excel status: ✅ Correcto ({meta.get('variant','?')}) — {os.path.basename(self.excel_file)}")
                self.excel_health_label.setStyleSheet("color: #1B5E20; font-weight: bold;")
            else:
                self.excel_health_label.setText("Excel status: ⚠️ Estructura inválida. Use 'Regenerar' o corrija el archivo.")
                self.excel_health_label.setStyleSheet("color: #E65100; font-weight: bold;")

            # Si cambió el archivo (mtime), refrescar preview
            if self._last_excel_mtime is None or mtime != self._last_excel_mtime:
                self._last_excel_mtime = mtime
                self.load_schedule_data()
        except Exception:
            self.excel_health_label.setText("Excel status: ⚠️ Error validando archivo.")
            self.excel_health_label.setStyleSheet("color: #E65100; font-weight: bold;")

    def prompt_regenerate_or_locate(self):
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Icon.Warning)
        msg.setWindowTitle("Archivo PlanStaff no disponible")
        msg.setText(f"No se encuentra el archivo:\n{self.excel_file}\n\n"
                    f"El sistema puede regenerarlo desde la BD (SSoT) o puede localizarlo manualmente.")
        regen_btn = msg.addButton("🛠️ Regenerar ahora", QMessageBox.ButtonRole.AcceptRole)
        locate_btn = msg.addButton("📂 Localizar archivo…", QMessageBox.ButtonRole.ActionRole)
        msg.addButton("Cancelar", QMessageBox.ButtonRole.RejectRole)
        msg.exec()

        if msg.clickedButton() == regen_btn:
            self.regenerate_excel_from_db()
        elif msg.clickedButton() == locate_btn:
            new_path, _ = QFileDialog.getOpenFileName(self, "Seleccionar PlanStaff", "", "Excel Files (*.xlsx)")
            if new_path:
                self.excel_file = new_path
                self._missing_prompt_shown = False
                self.check_excel_health()
                self.refresh_ui_data()

    def regenerate_excel_from_db(self):
        ok, msg = excel.regenerate_plan_from_db(self.excel_file, self.source)
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Icon.Information if ok else QMessageBox.Icon.Critical)
        box.setWindowTitle("Regeneración" if ok else "Error")
        box.setText(msg)
        box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
        box.exec()
        if ok:
            db.log_event(self.logged_username, self.source, "DATA_EXPORT", f"PLAN_REGENERATE -> {self.excel_file}")
            self._missing_prompt_shown = False
            self.check_excel_health()
            self.refresh_ui_data()

    def validate_excel_structure_ui(self):
        ok, errors, meta = excel.validate_excel_structure(self.excel_file)
        box = QMessageBox(self)
        if ok:
            box.setIcon(QMessageBox.Icon.Information)
            box.setWindowTitle("Estructura válida")
            box.setText(f"Plantilla: {meta.get('variant','?')} | Columnas de fecha: {meta.get('date_columns',0)}")
        else:
            box.setIcon(QMessageBox.Icon.Warning)
            box.setWindowTitle("Estructura inválida")
            box.setText("Se detectaron problemas:\n\n" + "\n".join(f"- {e}" for e in errors))
        box.addButton("OK", QMessageBox.ButtonRole.AcceptRole); box.exec()

    def compare_excel_db_ui(self):
        report = excel.check_db_sync_with_excel(self.excel_file, self.source)
        mismatches = report.get('schedule_mismatches', [])
        text = []
        text.append(f"Users in Excel: {report.get('users_in_excel',0)}")
        text.append(f"Users in DB:    {report.get('users_in_db',0)}")
        if report.get('missing_badges_in_db'):
            text.append(f"\nMissing in DB (badges): {', '.join(report['missing_badges_in_db'])}")
        if report.get('extra_badges_in_db'):
            text.append(f"Extra in DB (badges not in Excel): {', '.join(report['extra_badges_in_db'])}")
        text.append(f"\nSchedule mismatches: {len(mismatches)}")
        preview = "\n".join(text[:1000])

        # Mostrar en cuadro informativo (resumen)
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Icon.Information)
        box.setWindowTitle("Comparación Excel vs BD")
        box.setText(preview if len(preview) < 1500 else (preview[:1500] + "\n..."))
        box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
        box.exec()


# -------------------------------------------------------------
# Widget: CRUD de usuarios (con Import desde Excel)
# -------------------------------------------------------------
class CrudWidget(QWidget):
    # Señales para sincronización inmediata de UI
    import_done = pyqtSignal(str)   # emite el 'source' cuando termina el import
    users_changed = pyqtSignal(str) # emite el 'source' cuando cambia el listado de usuarios (crear/editar/elim)

    def __init__(self, source: str, excel_file: str, logged_username: str):
        super().__init__()
        self.source = source
        self.excel_file = excel_file
        self.logged_username = logged_username or "Unknown"
        self.current_user_id = None

        layout = QHBoxLayout(self)

        # Left panel: user form
        form_layout = QGridLayout()
        self.crud_name_input = QLineEdit()
        self.crud_role_input = QLineEdit()
        self.crud_badge_input = QLineEdit()

        self.crud_save_button = QPushButton("💾 Save User")
        self.crud_save_button.clicked.connect(self.save_crud_user)
        self.crud_new_button = QPushButton("✨ New User")
        self.crud_new_button.clicked.connect(self.clear_crud_form)
        self.crud_delete_button = QPushButton("❌ Delete User")
        self.crud_delete_button.clicked.connect(self.delete_crud_user)

        self.import_button = QPushButton("📥 Import from Excel → DB (validado)")
        self.import_button.clicked.connect(self.import_users_from_excel)

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
        table_layout = QVBoxLayout()
        self.users_table = QTableWidget()
        self.users_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.users_table.itemClicked.connect(self.load_user_to_crud_form)
        table_layout.addWidget(self.users_table)

        table_group = create_group_box("Registered Users List", table_layout)
        layout.addWidget(form_group)
        layout.addWidget(table_group)

        self.refresh_ui_data()

    # Tabla y formulario CRUD
    def load_users_table(self):
        users = db.get_all_users(self.source)
        headers = ["ID", "Name", "Role", "Badge"]
        self.users_table.setRowCount(len(users))
        self.users_table.setColumnCount(len(headers))
        self.users_table.setHorizontalHeaderLabels(headers)

        for row, user in enumerate(users):
            self.users_table.setItem(row, 0, QTableWidgetItem(str(user['id'])))
            self.users_table.setItem(row, 1, QTableWidgetItem(user['name']))
            self.users_table.setItem(row, 2, QTableWidgetItem(user['role']))
            self.users_table.setItem(row, 3, QTableWidgetItem(user['badge']))
        self.users_table.setColumnHidden(0, True) # El '0' es el índice de la columna 'ID'
        self.users_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

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
            success, message = db.update_user(self.current_user_id, name, role, badge, self.source)
        else:
            success, message = db.add_user(name, role, badge, self.source)

        box = QMessageBox(self)
        box.setIcon(QMessageBox.Icon.Information if success else QMessageBox.Icon.Warning)
        box.setWindowTitle("Success" if success else "Error")
        box.setText(message)
        box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
        box.exec()

        self.refresh_ui_data()
        # 🔔 sincronizar combo en Plan Staff
        self.users_changed.emit(self.source)

    def delete_crud_user(self):
        if not self.current_user_id:
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Warning)
            box.setWindowTitle("No Selection")
            box.setText("Please select a user in the table to delete.")
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
            box.exec()
            return

        confirm = QMessageBox(self)
        confirm.setIcon(QMessageBox.Icon.Question)
        confirm.setWindowTitle("Confirm Deletion")
        confirm.setText(f"Are you sure you want to delete {self.crud_name_input.text()}?")
        yes_btn = confirm.addButton("Yes", QMessageBox.ButtonRole.YesRole)
        confirm.addButton("No", QMessageBox.ButtonRole.NoRole)
        confirm.exec()

        if confirm.clickedButton() == yes_btn:
            success, message = db.delete_user(self.current_user_id)
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Information if success else QMessageBox.Icon.Warning)
            box.setWindowTitle("Success" if success else "Error")
            box.setText(message)
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
            box.exec()
            self.refresh_ui_data()
            # 🔔 sincronizar combo en Plan Staff
            self.users_changed.emit(self.source)

    def import_users_from_excel(self):
        """
        FR-02: Import users and day-by-day schedules from Excel to DB.
        Con validación previa de estructura. Si inválido -> no escribe y muestra error detallado.
        """
        try:
            inserted, skipped, upserts = excel.import_excel_to_db(self.excel_file, self.source)

            # Log de auditoría (FR-04)
            db.log_event(self.logged_username, self.source, "DATA_IMPORT",
                         f"users_inserted={inserted}; users_skipped={skipped}; schedule_upserts={upserts}")

            # Mensaje al usuario
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Information)
            box.setWindowTitle("Import Complete")
            box.setText(f"Imported {inserted} new users.\n"
                        f"Skipped {skipped} users that already existed.\n"
                        f"Upserted {upserts} schedule day-entries.")
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
            box.exec()

            # Refrescar tabla de usuarios inmediatamente
            self.refresh_ui_data()
            # 🔔 Señales para refrescar el combo "Select Employee" en Plan Staff
            self.users_changed.emit(self.source)
            self.import_done.emit(self.source)
        except ValueError as ve:
            # Error de estructura -> NO guardar nada
            db.log_event(self.logged_username, self.source, "DATA_IMPORT",
                         f"ERROR: {str(ve).replace(chr(10),' | ')}")
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Critical)
            box.setWindowTitle("Invalid Excel")
            box.setText(str(ve))
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
            box.exec()

    def refresh_ui_data(self):
        self.load_users_table()
        self.clear_crud_form()


# -------------------------------------------------------------
# Widget: Shift Types Admin (Admin y Site Managers)
# -------------------------------------------------------------
class ShiftTypeAdminWidget(QWidget):
    types_changed = pyqtSignal(str)  # emite source

    def __init__(self, source: str, excel_file: str, logged_username: str):
        super().__init__()
        self.source = source
        self.excel_file = excel_file
        self.logged_username = logged_username or "Unknown"
        self.current_type_id = None
        self.current_old_code = None

        layout = QHBoxLayout(self)

        # Left: form
        form_layout = QGridLayout()
        self.name_input = QLineEdit()
        self.code_input = QLineEdit()
        self.color_display = QLineEdit(); self.color_display.setReadOnly(True)
        self.pick_color_btn = QPushButton("🎨 Pick Color")
        self.pick_color_btn.clicked.connect(self.pick_color)

        self.in_time_edit = QTimeEdit(); self.in_time_edit.setDisplayFormat("HH:mm"); self.in_time_edit.setTime(QTime(8, 0))
        self.out_time_edit = QTimeEdit(); self.out_time_edit.setDisplayFormat("HH:mm"); self.out_time_edit.setTime(QTime(17, 0))

        self.new_btn = QPushButton("✨ New Shift Type")
        self.new_btn.clicked.connect(self.clear_form)
        self.save_btn = QPushButton("💾 Save")
        self.save_btn.clicked.connect(self.save_type)
        self.delete_btn = QPushButton("❌ Delete")
        self.delete_btn.clicked.connect(self.delete_type)

        form_layout.addWidget(QLabel("Name:"), 0, 0); form_layout.addWidget(self.name_input, 0, 1)
        form_layout.addWidget(QLabel("Code (short):"), 1, 0); form_layout.addWidget(self.code_input, 1, 1)
        form_layout.addWidget(QLabel("Color:"), 2, 0)
        h_color = QHBoxLayout(); h_color.addWidget(self.color_display); h_color.addWidget(self.pick_color_btn)
        form_layout.addLayout(h_color, 2, 1)
        form_layout.addWidget(QLabel("IN time (HH:MM):"), 3, 0); form_layout.addWidget(self.in_time_edit, 3, 1)
        form_layout.addWidget(QLabel("OUT time (HH:MM):"), 4, 0); form_layout.addWidget(self.out_time_edit, 4, 1)
        actions = QHBoxLayout(); actions.addWidget(self.new_btn); actions.addWidget(self.save_btn); actions.addWidget(self.delete_btn)
        form_layout.addLayout(actions, 5, 0, 1, 2)

        form_group = create_group_box(f"Configure Shift Types ({self.source})", form_layout)
        form_group.setFixedWidth(450)

        # Right: table
        table_layout = QVBoxLayout()
        self.types_table = QTableWidget()
        self.types_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.types_table.itemClicked.connect(self.load_type_to_form)
        table_layout.addWidget(self.types_table)

        table_group = create_group_box("Defined Shift Types", table_layout)

        layout.addWidget(form_group)
        layout.addWidget(table_group)

        self.refresh_table()

    def pick_color(self):
        initial = QColor(self.color_display.text() or "#FFC000")
        color = QColorDialog.getColor(initial, self, "Pick Color")
        if color.isValid():
            self.color_display.setText(color.name())

    def load_type_to_form(self, item):
        row = item.row()
        self.current_type_id = int(self.types_table.item(row, 0).text())
        name = self.types_table.item(row, 1).text()
        code = self.types_table.item(row, 2).text()
        color = self.types_table.item(row, 3).text()
        in_time = self.types_table.item(row, 4).text()
        out_time = self.types_table.item(row, 5).text()
        self.name_input.setText(name)
        self.code_input.setText(code)
        self.color_display.setText(color)
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
            box = QMessageBox(self); box.setIcon(QMessageBox.Icon.Warning)
            box.setWindowTitle("Incomplete Data"); box.setText("Name and Code are required.")
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole); box.exec(); return

        if self.current_type_id:
            ok, msg, old_code, new_code = db.update_shift_type(self.current_type_id, self.source, name, code, color_hex, in_time, out_time)
            if ok:
                # Si cambió el código -> actualizar Excel
                if old_code and new_code and old_code != new_code:
                    excel.apply_shift_type_update_to_excel(self.excel_file, self.source, old_code, new_code, color_hex)
                db.log_event(self.logged_username, self.source, "SHIFT_TYPE_UPDATE", f"{old_code} -> {new_code} | {name} {in_time}-{out_time} {color_hex}")
                self.types_changed.emit(self.source)
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Information if ok else QMessageBox.Icon.Warning)
            box.setWindowTitle("Save" if ok else "Error")
            box.setText(msg)
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
            box.exec()
        else:
            ok, msg = db.create_shift_type(self.source, name, code, color_hex, in_time, out_time)
            if ok:
                db.log_event(self.logged_username, self.source, "SHIFT_TYPE_CREATE", f"{code} | {name} {in_time}-{out_time} {color_hex}")
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
                db.log_event(self.logged_username, self.source, "SHIFT_TYPE_DELETE", f"{code}")
                self.types_changed.emit(self.source)
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Information if ok else QMessageBox.Icon.Warning)
            box.setWindowTitle("Delete" if ok else "Cannot delete")
            box.setText(msg)
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
            box.exec()

            self.refresh_table()
            self.clear_form()

    def refresh_table(self):
        types = db.get_shift_types(self.source)
        headers = ["ID", "Name", "Code", "Color", "IN", "OUT"]
        self.types_table.setRowCount(len(types))
        self.types_table.setColumnCount(len(headers))
        self.types_table.setHorizontalHeaderLabels(headers)
        for r, t in enumerate(types):
            self.types_table.setItem(r, 0, QTableWidgetItem(str(t['id'])))
            self.types_table.setItem(r, 1, QTableWidgetItem(t['name']))
            self.types_table.setItem(r, 2, QTableWidgetItem(t['code']))
            self.types_table.setItem(r, 3, QTableWidgetItem(t['color_hex']))
            self.types_table.setItem(r, 4, QTableWidgetItem(t['in_time']))
            self.types_table.setItem(r, 5, QTableWidgetItem(t['out_time']))
        self.types_table.setColumnHidden(0, True)
        self.types_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)


# -------------------------------------------------------------
# Ventana principal (perfil normal: RGM o Newmont)
# -------------------------------------------------------------
class MainWindow(QMainWindow):
    logout_signal = pyqtSignal()

    def __init__(self, user_role, excel_file, logged_username=None, can_manage_shift_types: bool = False):
        super().__init__()
        self.user_role = user_role  # RGM or Newmont. Used as 'source'
        self.excel_file = excel_file
        self.logged_username = logged_username or "Unknown"
        self.can_manage_shift_types = bool(can_manage_shift_types)

        db.setup_database()
        db.log_event(self.logged_username, self.user_role, "USER_LOGIN", f"Excel={self.excel_file}")

        self.setWindowTitle(f"👨‍✈️ Operations Manager - Profile: {self.user_role} | User: {self.logged_username}")
        self.setGeometry(100, 100, 1200, 800)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Top bar
        top_layout = QHBoxLayout()
        title_label = QLabel(f"Transport & Operations Manager ({self.user_role})")
        font = title_label.font(); font.setPointSize(20); font.setBold(True); title_label.setFont(font)

        self.logged_user_label = QLabel(f"👤 {self.logged_username}")
        lu_font = self.logged_user_label.font(); lu_font.setBold(True); self.logged_user_label.setFont(lu_font)
        self.logged_user_label.setStyleSheet("padding: 0 12px;")

        logout_button = QPushButton("🔒 Sign Out")
        logout_button.setFixedWidth(150)
        logout_button.clicked.connect(self.handle_logout)

        top_layout.addWidget(title_label)
        top_layout.addStretch()
        top_layout.addWidget(self.logged_user_label)
        top_layout.addWidget(logout_button)
        main_layout.addLayout(top_layout)

        # Tabs
        tabs = QTabWidget()
        main_layout.addWidget(tabs)

        # 1) Plan Staff (con preview/registro/reportes)
        self.plan_widget = PlanStaffWidget(self.user_role, self.excel_file, self.logged_username)
        tabs.addTab(self.plan_widget, "📅 Plan Staff & Reports")

        # 2) Users CRUD
        self.crud_widget = CrudWidget(self.user_role, self.excel_file, self.logged_username)
        tabs.addTab(self.crud_widget, "👥 Users (CRUD)")

        # 3) Shift Types (solo si el usuario puede gestionarlos)
        if self.can_manage_shift_types:
            self.shift_types_widget = ShiftTypeAdminWidget(self.user_role, self.excel_file, self.logged_username)
            tabs.addTab(self.shift_types_widget, f"⚙️ {self.user_role} Shift Types")
            # Refrescar combos/preview cuando cambien tipos
            self.shift_types_widget.types_changed.connect(lambda src: self.plan_widget.refresh_ui_data())

        # 🔗 Conexiones para refrescar "Select Employee" inmediatamente
        self.crud_widget.users_changed.connect(self._sync_after_users_changed)
        self.crud_widget.import_done.connect(self._sync_after_users_changed)

    def _sync_after_users_changed(self, src: str):
        if src == self.user_role:
            self.plan_widget.refresh_users_only()
            QApplication.processEvents()  # asegura que la UI se repinta

    def handle_logout(self):
        self.logout_signal.emit()
        self.close()


# -------------------------------------------------------------
# Ventana de Administrador (acceso unificado)
# -------------------------------------------------------------
class AdminMainWindow(QMainWindow):
    logout_signal = pyqtSignal()

    def __init__(self, logged_username: str, rgm_excel: str, newmont_excel: str):
        super().__init__()
        self.logged_username = logged_username or "admin"

        db.setup_database()
        db.log_event(self.logged_username, "Administrator", "USER_LOGIN",
                     f"Access to admin console | RGM={rgm_excel} | Newmont={newmont_excel}")

        self.setWindowTitle(f"🛡️ Administrator Console | User: {self.logged_username}")
        self.setGeometry(100, 100, 1400, 900)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Top bar
        top_layout = QHBoxLayout()
        title_label = QLabel("Unified Access — RGM & Newmont")
        font = title_label.font(); font.setPointSize(20); font.setBold(True); title_label.setFont(font)

        self.logged_user_label = QLabel(f"👤 {self.logged_username} (Administrator)")
        lu_font = self.logged_user_label.font(); lu_font.setBold(True); self.logged_user_label.setFont(lu_font)
        self.logged_user_label.setStyleSheet("padding: 0 12px;")

        logout_button = QPushButton("🔒 Sign Out")
        logout_button.setFixedWidth(150)
        logout_button.clicked.connect(self.handle_logout)

        top_layout.addWidget(title_label)
        top_layout.addStretch()
        top_layout.addWidget(self.logged_user_label)
        top_layout.addWidget(logout_button)
        main_layout.addLayout(top_layout)

        # Tabs unificadas
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

        # 5) Audit Log (global)
        audit_all = AuditLogWidget(source=None)
        self.tabs.addTab(audit_all, "📝 Audit Log")

        # 6) Shift Types (Admin para ambos sites)
        self.rgm_types = ShiftTypeAdminWidget("RGM", rgm_excel, self.logged_username)
        self.tabs.addTab(self.rgm_types, "⚙️ RGM Shift Types")

        self.nm_types = ShiftTypeAdminWidget("Newmont", newmont_excel, self.logged_username)
        self.tabs.addTab(self.nm_types, "⚙️ Newmont Shift Types")

        # 🔗 Sincronización en caliente
        self.rgm_crud.users_changed.connect(lambda src: self.rgm_plan.refresh_users_only())
        self.rgm_crud.import_done.connect(lambda src: self.rgm_plan.refresh_users_only())

        self.nm_crud.users_changed.connect(lambda src: self.nm_plan.refresh_users_only())
        self.nm_crud.import_done.connect(lambda src: self.nm_plan.refresh_users_only())

        self.rgm_types.types_changed.connect(lambda src: self.rgm_plan.refresh_ui_data())
        self.nm_types.types_changed.connect(lambda src: self.nm_plan.refresh_ui_data())

    def handle_logout(self):
        self.logout_signal.emit()
        self.close()


# -------------------------------------------------------------
# Widget: Audit Log (visible para Admin; reutilizable si se desea)
# -------------------------------------------------------------
class AuditLogWidget(QWidget):
    def __init__(self, source: str | None):
        super().__init__()
        self.source = source
        layout = QVBoxLayout(self)
        self.audit_table = QTableWidget()
        layout.addWidget(self.audit_table)

        refresh_btn = QPushButton("🔄 Refresh")
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
            self.audit_table.setItem(r, 0, QTableWidgetItem(ev.get('ts', '')))
            self.audit_table.setItem(r, 1, QTableWidgetItem(ev.get('username', '')))
            self.audit_table.setItem(r, 2, QTableWidgetItem(ev.get('source', '')))
            self.audit_table.setItem(r, 3, QTableWidgetItem(ev.get('action_type', '')))
            self.audit_table.setItem(r, 4, QTableWidgetItem(ev.get('detail', '')))

        self.audit_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
