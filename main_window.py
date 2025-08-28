from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, 
    QLineEdit, QComboBox, QDateEdit, QPushButton, QTableWidget, 
    QTableWidgetItem, QHeaderView, QGroupBox, QRadioButton, QMessageBox, QFileDialog,
    QTabWidget
)
from PyQt6.QtCore import QDate, Qt, pyqtSignal
from PyQt6.QtGui import QColor
from datetime import datetime
import os

# Importar nuestra lógica separada
import database_logic as db
import excel_logic as excel

class MainWindow(QMainWindow):
    logout_signal = pyqtSignal()

    def __init__(self, user_role, excel_file):
        super().__init__()
        self.user_role = user_role # RGM o Newmont. Se usará como 'source'
        self.excel_file = excel_file
        self.setWindowTitle(f"👨‍✈️ Gestor de Operaciones - Perfil: {self.user_role}")
        self.setGeometry(100, 100, 1200, 800)
        self.current_user_id = None

        db.setup_database()

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        top_layout = QHBoxLayout()
        title_label = QLabel(f"Gestor de Transporte y Operaciones ({self.user_role})")
        font = title_label.font(); font.setPointSize(20); font.setBold(True)
        title_label.setFont(font)
        
        logout_button = QPushButton("🔒 Cerrar Sesión")
        logout_button.setFixedWidth(150)
        logout_button.clicked.connect(self.handle_logout)

        top_layout.addWidget(title_label)
        top_layout.addStretch()
        top_layout.addWidget(logout_button)
        main_layout.addLayout(top_layout)

        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        self.plan_staff_tab = QWidget()
        self.tabs.addTab(self.plan_staff_tab, "📅 Plan Staff y Reportes")
        self.setup_plan_staff_ui()

        self.crud_tab = QWidget()
        self.tabs.addTab(self.crud_tab, "👥 CRUD Usuarios")
        self.setup_crud_ui()

        self.refresh_ui_data()

    def setup_plan_staff_ui(self):
        layout = QVBoxLayout(self.plan_staff_tab)
        self.setup_schedule_preview_ui(layout)
        forms_db_layout = QHBoxLayout()
        self.setup_registration_form_ui(forms_db_layout)
        self.setup_db_view_ui(forms_db_layout)
        layout.addLayout(forms_db_layout)
        self.setup_report_generator_ui(layout)

    def setup_crud_ui(self):
        layout = QHBoxLayout(self.crud_tab)
        
        form_layout = QGridLayout()
        self.crud_name_input = QLineEdit()
        self.crud_role_input = QLineEdit()
        self.crud_badge_input = QLineEdit()
        
        self.crud_save_button = QPushButton("💾 Guardar Usuario")
        self.crud_save_button.clicked.connect(self.save_crud_user)
        self.crud_new_button = QPushButton("✨ Nuevo Usuario")
        self.crud_new_button.clicked.connect(self.clear_crud_form)
        self.crud_delete_button = QPushButton("❌ Eliminar Usuario")
        self.crud_delete_button.clicked.connect(self.delete_crud_user)
        
        self.import_button = QPushButton("📥 Importar desde Excel")
        self.import_button.clicked.connect(self.import_users_from_excel)

        form_layout.addWidget(QLabel("Nombre y Apellido:"), 0, 0)
        form_layout.addWidget(self.crud_name_input, 0, 1)
        form_layout.addWidget(QLabel("Rol/Departamento:"), 1, 0)
        form_layout.addWidget(self.crud_role_input, 1, 1)
        form_layout.addWidget(QLabel("Badge (ID):"), 2, 0)
        form_layout.addWidget(self.crud_badge_input, 2, 1)
        
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.crud_new_button)
        button_layout.addWidget(self.crud_save_button)
        form_layout.addLayout(button_layout, 3, 0, 1, 2)
        form_layout.addWidget(self.crud_delete_button, 4, 0, 1, 2)
        form_layout.addWidget(self.import_button, 5, 0, 1, 2)

        form_group = self.create_group_box("Gestionar Usuario", form_layout)
        form_group.setFixedWidth(400)
        
        table_layout = QVBoxLayout()
        self.users_table = QTableWidget()
        self.users_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.users_table.itemClicked.connect(self.load_user_to_crud_form)
        table_layout.addWidget(self.users_table)
        
        table_group = self.create_group_box("Lista de Usuarios Registrados", table_layout)
        layout.addWidget(form_group)
        layout.addWidget(table_group)

    def handle_logout(self):
        self.logout_signal.emit()
        self.close()

    def create_group_box(self, title, layout):
        box = QGroupBox(title)
        font = box.font(); font.setBold(True)
        box.setFont(font)
        box.setLayout(layout)
        return box

    def setup_schedule_preview_ui(self, parent_layout):
        schedule_container_layout = QVBoxLayout()
        tables_layout = QHBoxLayout()
        tables_layout.setSpacing(0); tables_layout.setContentsMargins(0, 0, 0, 0)
        self.frozen_table = QTableWidget()
        self.schedule_table = QTableWidget()
        tables_layout.addWidget(self.frozen_table); tables_layout.addWidget(self.schedule_table, 1)
        schedule_container_layout.addLayout(tables_layout)
        self.frozen_table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.schedule_table.verticalScrollBar().valueChanged.connect(self.frozen_table.verticalScrollBar().setValue)
        self.frozen_table.verticalScrollBar().valueChanged.connect(self.schedule_table.verticalScrollBar().setValue)
        preview_title = f"🗓️ Vista Previa del Cronograma ({os.path.basename(self.excel_file)})"
        parent_layout.addWidget(self.create_group_box(preview_title, schedule_container_layout))

    def load_schedule_data(self):
        FROZEN_COLUMN_COUNT = 3
        df = excel.get_schedule_preview(self.excel_file)
        if df.empty: return
        actual_frozen_count = min(df.shape[1], FROZEN_COLUMN_COUNT)
        headers = [str(col.strftime('%Y-%m-%d')) if isinstance(col, datetime) else str(col) for col in df.columns]
        self.frozen_table.setRowCount(df.shape[0]); self.frozen_table.setColumnCount(actual_frozen_count)
        self.frozen_table.setHorizontalHeaderLabels(headers[:actual_frozen_count])
        self.schedule_table.setRowCount(df.shape[0]); self.schedule_table.setColumnCount(df.shape[1] - actual_frozen_count)
        self.schedule_table.setHorizontalHeaderLabels(headers[actual_frozen_count:])
        for i, row in df.iterrows():
            for j, val in enumerate(row):
                item = QTableWidgetItem(str(val))
                if j < actual_frozen_count: self.frozen_table.setItem(i, j, item)
                else:
                    col_index = j - actual_frozen_count
                    val_str = str(val).upper()
                    if 'ON NS' in val_str or 'NIGHT' in val_str: item.setBackground(QColor("#FFFF99"))
                    elif 'ON' in val_str or 'DAY' in val_str: item.setBackground(QColor("#C6EFCE"))
                    elif 'OFF' in val_str: item.setBackground(QColor("#FFC7CE"))
                    elif 'LEAVE' in val_str: item.setBackground(QColor("#D9D9D9"))
                    self.schedule_table.setItem(i, col_index, item)
        self.frozen_table.resizeColumnsToContents(); self.schedule_table.resizeColumnsToContents()

    def setup_registration_form_ui(self, parent_layout):
        form_layout = QGridLayout()
        self.user_selector_combo = QComboBox()
        self.user_selector_combo.currentIndexChanged.connect(self.autofill_user_data)
        self.role_display = QLineEdit(); self.role_display.setReadOnly(True)
        self.badge_display = QLineEdit(); self.badge_display.setReadOnly(True)
        self.shift_combo = QComboBox(); self.shift_combo.addItems(["Day Shift", "Night Shift"])
        self.on_radio = QRadioButton("ON"); self.off_radio = QRadioButton("OFF")
        self.none_radio = QRadioButton("No Marcar Días"); self.on_radio.setChecked(True)
        self.on_radio.toggled.connect(lambda: self.shift_combo.setEnabled(self.on_radio.isChecked()))
        self.off_radio.toggled.connect(lambda: self.shift_combo.setEnabled(self.on_radio.isChecked()))
        self.none_radio.toggled.connect(lambda: self.shift_combo.setEnabled(self.on_radio.isChecked()))
        self.start_date_edit = QDateEdit(QDate.currentDate()); self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDisplayFormat("dd/MM/yyyy")
        self.end_date_edit = QDateEdit(QDate.currentDate().addDays(14)); self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setDisplayFormat("dd/MM/yyyy")
        save_button = QPushButton("✅ Guardar Cambios en DB y Excel")
        save_button.clicked.connect(self.save_plan_changes)
        form_layout.addWidget(QLabel("Seleccionar Empleado:"), 0, 0); form_layout.addWidget(self.user_selector_combo, 0, 1)
        form_layout.addWidget(QLabel("Rol/Departamento:"), 1, 0); form_layout.addWidget(self.role_display, 1, 1)
        form_layout.addWidget(QLabel("Badge (ID):"), 2, 0); form_layout.addWidget(self.badge_display, 2, 1)
        form_layout.addWidget(QLabel("Turno:"), 3, 0); form_layout.addWidget(self.shift_combo, 3, 1)
        status_layout = QHBoxLayout(); status_layout.addWidget(self.on_radio); status_layout.addWidget(self.off_radio); status_layout.addWidget(self.none_radio)
        form_layout.addWidget(QLabel("Estado a Marcar:"), 4, 0); form_layout.addLayout(status_layout, 4, 1)
        form_layout.addWidget(QLabel("Fecha Inicio Período:"), 5, 0); form_layout.addWidget(self.start_date_edit, 5, 1)
        form_layout.addWidget(QLabel("Fecha Final Período:"), 6, 0); form_layout.addWidget(self.end_date_edit, 6, 1)
        form_layout.addWidget(save_button, 7, 0, 1, 2)
        self.registration_groupbox = self.create_group_box("1. Registrar Horario de Empleado", form_layout)
        parent_layout.addWidget(self.registration_groupbox, 1)

    def save_plan_changes(self):
        username = self.user_selector_combo.currentText()
        badge = self.badge_display.text()
        role = self.role_display.text()
        start_date = self.start_date_edit.date().toPyDate()
        end_date = self.end_date_edit.date().toPyDate()
        if not username or username == "-- Seleccione un usuario --":
            QMessageBox.warning(self, "Datos incompletos", "Por favor, seleccione un empleado.")
            return
        if start_date > end_date:
            QMessageBox.warning(self, "Error de Fechas", "La fecha de inicio no puede ser posterior a la final.")
            return
        schedule_status = "OFF"; shift_type = None
        if self.on_radio.isChecked():
            schedule_status = "ON"; shift_type = self.shift_combo.currentText()
        elif self.none_radio.isChecked():
            schedule_status = None
        db.add_operation(username, role, badge, start_date, end_date)
        success, message = excel.update_plan_staff_excel(
            self.excel_file, username, role, badge, schedule_status, shift_type, start_date, end_date
        )
        if success: QMessageBox.information(self, "Éxito", f"DB y Excel actualizados para {username}.")
        else: QMessageBox.warning(self, "Error de Excel", f"Guardado en DB, pero error en Excel: {message}")
        self.refresh_ui_data()
    
    def setup_db_view_ui(self, parent_layout):
        db_view_layout = QVBoxLayout()
        self.db_table = QTableWidget()
        db_view_layout.addWidget(self.db_table)
        self.db_view_groupbox = self.create_group_box("Historial de Rotaciones", db_view_layout)
        parent_layout.addWidget(self.db_view_groupbox, 2)

    def setup_report_generator_ui(self, parent_layout):
        report_layout = QHBoxLayout()
        report_layout.addWidget(QLabel("Fecha INICIO:"))
        self.report_start_date = QDateEdit(QDate.currentDate()); self.report_start_date.setCalendarPopup(True)
        self.report_start_date.setDisplayFormat("dd/MM/yyyy"); report_layout.addWidget(self.report_start_date)
        report_layout.addWidget(QLabel("Fecha FINAL:"))
        self.report_end_date = QDateEdit(QDate.currentDate().addDays(30)); self.report_end_date.setCalendarPopup(True)
        self.report_end_date.setDisplayFormat("dd/MM/yyyy"); report_layout.addWidget(self.report_end_date)
        report_button = QPushButton("🚀 Generar Reporte"); report_button.clicked.connect(self.generate_report)
        report_layout.addWidget(report_button)
        report_title = f"2. Generar Reporte de Transporte (desde {os.path.basename(self.excel_file)})"
        parent_layout.addWidget(self.create_group_box(report_title, report_layout))

    def load_db_data(self):
        records = db.get_all_operations()
        headers = ["ID", "Nombre", "Rol", "Badge", "Fecha Inicio", "Fecha Fin"]
        self.db_table.setRowCount(len(records)); self.db_table.setColumnCount(len(headers))
        self.db_table.setHorizontalHeaderLabels(headers)
        for row_idx, record in enumerate(records):
            self.db_table.setItem(row_idx, 0, QTableWidgetItem(str(record['id'])))
            self.db_table.setItem(row_idx, 1, QTableWidgetItem(record['username']))
            self.db_table.setItem(row_idx, 2, QTableWidgetItem(record['role']))
            self.db_table.setItem(row_idx, 3, QTableWidgetItem(record['badge']))
            self.db_table.setItem(row_idx, 4, QTableWidgetItem(record['start_date']))
            self.db_table.setItem(row_idx, 5, QTableWidgetItem(record['end_date']))
        self.db_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

    def refresh_ui_data(self):
        self.load_schedule_data()
        self.load_db_data()
        self.load_crud_users_table()
        self.load_users_to_selector()
        self.clear_crud_form()

    def generate_report(self):
        start_date = self.report_start_date.date().toPyDate()
        end_date = self.report_end_date.date().toPyDate()
        if start_date > end_date:
            QMessageBox.warning(self, "Error de Fechas", "La fecha de inicio no puede ser posterior a la final.")
            return
        excel_data, message = excel.generate_transport_excel_from_planstaff(self.excel_file, start_date, end_date)
        if not excel_data:
            QMessageBox.information(self, "Información", message)
            return
        default_filename = f"Transport_Request_{start_date.strftime('%Y%m%d')}_to_{end_date.strftime('%Y%m%d')}.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(self, "Guardar Reporte", default_filename, "Archivos de Excel (*.xlsx)")
        if file_path:
            try:
                with open(file_path, 'wb') as f: f.write(excel_data)
                QMessageBox.information(self, "Éxito", f"{message}\n\nReporte guardado en:\n{file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Error al Guardar", f"No se pudo guardar el archivo.\nError: {e}")

    def load_crud_users_table(self):
        self.users = db.get_all_users(self.user_role)
        headers = ["ID", "Nombre", "Rol", "Badge"]
        self.users_table.setRowCount(len(self.users)); self.users_table.setColumnCount(len(headers))
        self.users_table.setHorizontalHeaderLabels(headers)
        for row, user in enumerate(self.users):
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
        name = self.crud_name_input.text()
        role = self.crud_role_input.text()
        badge = self.crud_badge_input.text()
        if not name or not role or not badge:
            QMessageBox.warning(self, "Datos Incompletos", "Todos los campos son obligatorios.")
            return
        if self.current_user_id:
            success, message = db.update_user(self.current_user_id, name, role, badge, self.user_role)
        else:
            success, message = db.add_user(name, role, badge, self.user_role)
        if success: QMessageBox.information(self, "Éxito", message)
        else: QMessageBox.warning(self, "Error", message)
        self.refresh_ui_data()

    def delete_crud_user(self):
        if not self.current_user_id:
            QMessageBox.warning(self, "Sin Selección", "Por favor, seleccione un usuario de la tabla para eliminar.")
            return
        reply = QMessageBox.question(self, "Confirmar Eliminación", 
            f"¿Está seguro de que desea eliminar a {self.crud_name_input.text()}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            success, message = db.delete_user(self.current_user_id)
            if success: QMessageBox.information(self, "Éxito", message)
            else: QMessageBox.warning(self, "Error", message)
            self.refresh_ui_data()

    def load_users_to_selector(self):
        self.user_selector_combo.clear()
        self.users_for_selector = db.get_all_users(self.user_role)
        self.user_selector_combo.addItem("-- Seleccione un usuario --")
        for user in self.users_for_selector:
            self.user_selector_combo.addItem(user['name'])
    
    def autofill_user_data(self, index):
        if index > 0:
            user = self.users_for_selector[index - 1]
            self.role_display.setText(user['role'])
            self.badge_display.setText(user['badge'])
        else:
            self.role_display.clear()
            self.badge_display.clear()

    def import_users_from_excel(self):
        """
        Importa usuarios desde el archivo Excel a la base de datos usando una
        operación en bloque para evitar bloqueos de la base de datos.
        """
        users = excel.get_users_from_excel(self.excel_file)
        if not users:
            QMessageBox.warning(self, "Importación Fallida", "No se encontraron usuarios en el archivo Excel o el archivo no se pudo leer.")
            return
        
        added_count = db.add_users_bulk(users, self.user_role)
        skipped_count = len(users) - added_count
        
        QMessageBox.information(self, "Importación Completa", 
            f"Se importaron {added_count} nuevos usuarios.\nSe omitieron {skipped_count} usuarios que ya existían.")
        
        self.refresh_ui_data()
