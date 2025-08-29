from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel,
    QLineEdit, QComboBox, QDateEdit, QPushButton, QTableWidget,
    QTableWidgetItem, QHeaderView, QGroupBox, QRadioButton, QMessageBox, QFileDialog,
    QTabWidget, QApplication
)
from PyQt6.QtCore import QDate, Qt, pyqtSignal
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

        root = QVBoxLayout(self)

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
        self.registration_groupbox = create_group_box("1. Register Employee Schedule", self._build_registration_form())
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

        export_button = QPushButton("📤 Export Plan Staff (.xlsx)")
        export_button.clicked.connect(self.export_plan_from_db)
        report_layout.addWidget(export_button)

        report_title = f"2. Generate Transportation Report (from {os.path.basename(self.excel_file)})"
        root.addWidget(create_group_box(report_title, report_layout))

        # Cargar datos iniciales
        self.refresh_ui_data()

    # ---------- sub-UIs ----------
    def _build_registration_form(self):
        form_layout = QGridLayout()

        self.user_selector_combo = QComboBox()
        self.user_selector_combo.currentIndexChanged.connect(self.autofill_user_data)

        self.role_display = QLineEdit(); self.role_display.setReadOnly(True)
        self.badge_display = QLineEdit(); self.badge_display.setReadOnly(True)

        self.shift_combo = QComboBox()
        self.shift_combo.addItems(["Day Shift", "Night Shift"])

        self.on_radio = QRadioButton("ON")
        self.off_radio = QRadioButton("OFF")
        self.none_radio = QRadioButton("Do Not Mark Days")
        self.on_radio.setChecked(True)
        # habilitar shift solo si ON
        self.on_radio.toggled.connect(lambda: self.shift_combo.setEnabled(self.on_radio.isChecked()))
        self.off_radio.toggled.connect(lambda: self.shift_combo.setEnabled(self.on_radio.isChecked()))
        self.none_radio.toggled.connect(lambda: self.shift_combo.setEnabled(self.on_radio.isChecked()))

        self.start_date_edit = QDateEdit(QDate.currentDate()); self.start_date_edit.setCalendarPopup(True); self.start_date_edit.setDisplayFormat("dd/MM/yyyy")
        self.end_date_edit = QDateEdit(QDate.currentDate().addDays(14)); self.end_date_edit.setCalendarPopup(True); self.end_date_edit.setDisplayFormat("dd/MM/yyyy")

        save_button = QPushButton("✅ Save Changes to DB & Excel")
        save_button.clicked.connect(self.save_plan_changes)

        form_layout.addWidget(QLabel("Select Employee:"), 0, 0); form_layout.addWidget(self.user_selector_combo, 0, 1)
        form_layout.addWidget(QLabel("Role/Department:"), 1, 0); form_layout.addWidget(self.role_display, 1, 1)
        form_layout.addWidget(QLabel("Badge (ID):"), 2, 0); form_layout.addWidget(self.badge_display, 2, 1)
        form_layout.addWidget(QLabel("Shift:"), 3, 0); form_layout.addWidget(self.shift_combo, 3, 1)

        status_layout = QHBoxLayout()
        status_layout.addWidget(self.on_radio); status_layout.addWidget(self.off_radio); status_layout.addWidget(self.none_radio)

        form_layout.addWidget(QLabel("Status to Mark:"), 4, 0)
        form_layout.addLayout(status_layout, 4, 1)
        form_layout.addWidget(QLabel("Period Start Date:"), 5, 0); form_layout.addWidget(self.start_date_edit, 5, 1)
        form_layout.addWidget(QLabel("Period End Date:"), 6, 0); form_layout.addWidget(self.end_date_edit, 6, 1)
        form_layout.addWidget(save_button, 7, 0, 1, 2)

        return form_layout

    # ---------- data loaders ----------
    def load_schedule_data(self):
        FROZEN_COLUMN_COUNT = 3
        df = excel.get_schedule_preview(self.excel_file)
        if df.empty:
            self.frozen_table.clear(); self.schedule_table.clear()
            self.frozen_table.setRowCount(0); self.schedule_table.setRowCount(0)
            return

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
                text = _clean(val)  # FR-07
                item = QTableWidgetItem(text)

                if j < actual_frozen_count:
                    self.frozen_table.setItem(i, j, item)
                else:
                    col_index = j - actual_frozen_count
                    val_str = text.upper()  # Usar texto ya saneado
                    if 'ON NS' in val_str or 'NIGHT' in val_str:
                        item.setBackground(QColor("#FFFF99"))
                    elif 'ON' in val_str or 'DAY' in val_str or val_str.isdigit():
                        item.setBackground(QColor("#C6EFCE"))
                    elif 'OFF' in val_str or 'BREAK' in val_str or 'KO' in val_str:
                        item.setBackground(QColor("#FFC7CE"))
                    elif 'LEAVE' in val_str:
                        item.setBackground(QColor("#D9D9D9"))
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
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Warning)
            box.setWindowTitle("Incomplete Data")
            box.setText("Please select an employee.")
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
            box.exec()
            return

        if not role:
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Warning)
            box.setWindowTitle("Incomplete Data")
            box.setText("Please select a role/department.")
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
            box.exec()
            return

        if start_date > end_date:
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Warning)
            box.setWindowTitle("Date Error")
            box.setText("Start date cannot be after end date.")
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
            box.exec()
            return

        schedule_status = "OFF"
        shift_type = None
        if self.on_radio.isChecked():
            schedule_status = "ON"
            shift_type = self.shift_combo.currentText()
        elif self.none_radio.isChecked():
            schedule_status = None

        # --- FR-01: Confirmación de sobrescritura (Excel + BD) ---
        conflicts_excel = excel.find_conflicts(self.excel_file, username, badge, start_date, end_date)
        conflicts_db_map = db.get_schedule_map_for_range(badge, start_date, end_date, self.source)
        if conflicts_excel or conflicts_db_map:
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Warning)
            box.setWindowTitle("Overwrite Shift Confirmation")
            box.setText("Are you sure you want to modify the existing shift?")
            accept_btn = box.addButton("Accept", QMessageBox.ButtonRole.AcceptRole)
            box.addButton("Cancel", QMessageBox.ButtonRole.RejectRole)
            box.exec()
            if box.clickedButton() != accept_btn:
                return  # abortar

        # --- Datos previos para auditoría (FR-04) ---
        prev_map = db.get_schedule_map_for_range(badge, start_date, end_date, self.source)

        # --- Persistencia en BD ---
        if schedule_status in ("ON", "OFF"):
            # histórico de rango
            db.add_operation(username, role, badge, start_date, end_date)
            # estado día a día
            db.upsert_schedule_range(badge, start_date, end_date, schedule_status, shift_type, self.source)
        else:
            # limpiar estado día a día en BD cuando se elige "Do Not Mark Days"
            db.clear_schedule_range(badge, start_date, end_date, self.source)

        # --- Actualizar Excel
        success, message = excel.update_plan_staff_excel(
            self.excel_file, username, role, badge, schedule_status, shift_type, start_date, end_date
        )

        # --- Auditoría (FR-04): detalle con prev y nuevo ---
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

    def generate_report(self):
        s = self.report_start_date.date().toPyDate()
        e = self.report_end_date.date().toPyDate()
        excel_data, message = excel.generate_transport_report(self.excel_file, s, e)

        file_path, _ = QFileDialog.getSaveFileName(self, "Save Transportation Report", "transport_request.xlsx", "Excel Files (*.xlsx)")
        if not file_path:
            return
        try:
            with open(file_path, 'wb') as f:
                f.write(excel_data)
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Information)
            box.setWindowTitle("Success")
            box.setText(f"{message}\n\nReport saved to:\n{file_path}")
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
            box.exec()

            # FR-04: log exportación
            db.log_event(self.logged_username, self.source, "DATA_EXPORT", f"TRANSPORT -> {file_path}")
        except Exception as e:
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Critical)
            box.setWindowTitle("Save Error")
            box.setText(f"Could not save the file.\nError: {e}")
            box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
            box.exec()

    def export_plan_from_db(self):
        """FR-03: Exporta una planilla desde el estado actual en la BD, manteniendo formato del template actual."""
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

        default_name = f"PlanStaff_{self.source}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        dest_path, _ = QFileDialog.getSaveFileName(self, "Save Plan Staff", default_name, "Excel Files (*.xlsx)")
        if not dest_path:
            return

        ok, msg = excel.export_plan_from_db(self.excel_file, users, schedules, dest_path)
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Icon.Information if ok else QMessageBox.Icon.Critical)
        box.setWindowTitle("Export" if ok else "Error")
        box.setText(msg)
        box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
        box.exec()

        if ok:
            db.log_event(self.logged_username, self.source, "DATA_EXPORT", f"PLAN_EXPORT -> {dest_path}")

    def refresh_ui_data(self):
        self.load_schedule_data()
        self.load_db_data()
        self.load_users_to_selector()


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

        self.import_button = QPushButton("📥 Import from Excel")
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
        Sincroniza UI en caliente (emite señal users_changed + import_done).
        """
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

    def refresh_ui_data(self):
        self.load_users_table()
        self.clear_crud_form()


# -------------------------------------------------------------
# Widget: Audit Log (visible solo para Admin)
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


# -------------------------------------------------------------
# Ventana principal (perfil normal)
# -------------------------------------------------------------
class MainWindow(QMainWindow):
    logout_signal = pyqtSignal()

    def __init__(self, user_role, excel_file, logged_username=None):
        super().__init__()
        self.user_role = user_role  # RGM or Newmont. Used as 'source'
        self.excel_file = excel_file
        self.logged_username = logged_username or "Unknown"

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

        # Pestaña Plan Staff (con preview/registro/reportes)
        self.plan_widget = PlanStaffWidget(self.user_role, self.excel_file, self.logged_username)
        tabs.addTab(self.plan_widget, "📅 Plan Staff & Reports")

        # Pestaña CRUD
        self.crud_widget = CrudWidget(self.user_role, self.excel_file, self.logged_username)
        tabs.addTab(self.crud_widget, "👥 Users (CRUD)")

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
# Ventana de Administrador (perfil con acceso unificado)
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

        # FR-05.3: Pestañas visibles para Admin
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

        # 🔗 Sincronización en caliente para Admin (cada CRUD refresca su Plan Staff)
        self.rgm_crud.users_changed.connect(lambda src: self.rgm_plan.refresh_users_only())
        self.rgm_crud.import_done.connect(lambda src: self.rgm_plan.refresh_users_only())

        self.nm_crud.users_changed.connect(lambda src: self.nm_plan.refresh_users_only())
        self.nm_crud.import_done.connect(lambda src: self.nm_plan.refresh_users_only())

    def handle_logout(self):
        self.logout_signal.emit()
        self.close()
