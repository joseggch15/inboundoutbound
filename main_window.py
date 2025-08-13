import sys
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QLabel, QLineEdit, QComboBox, QDateEdit, QPushButton,
    QTableWidget, QTableWidgetItem, QHeaderView, QGroupBox, QRadioButton,
    QMessageBox, QFileDialog
)
from PyQt6.QtCore import QDate, Qt
from PyQt6.QtGui import QColor, QFont
from datetime import date, timedelta, datetime

# Importar nuestra lógica separada
import database_logic as db
import excel_logic as excel

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("👨‍✈️ Gestor de Transporte y Operaciones")
        self.setGeometry(100, 100, 1200, 800)

        db.setup_database()

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        title_label = QLabel("Gestor de Transporte y Operaciones")
        font = title_label.font()
        font.setPointSize(20)
        font.setBold(True)
        title_label.setFont(font)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(title_label)

        self.setup_schedule_preview_ui(main_layout)

        forms_db_layout = QHBoxLayout()
        self.setup_registration_form_ui(forms_db_layout)
        self.setup_db_view_ui(forms_db_layout)
        main_layout.addLayout(forms_db_layout)

        self.setup_report_generator_ui(main_layout)
        
        self.refresh_ui_data()

    def create_group_box(self, title, layout):
        box = QGroupBox(title)
        font = box.font()
        font.setBold(True)
        box.setFont(font)
        box.setLayout(layout)
        return box

    def setup_schedule_preview_ui(self, parent_layout):
        schedule_container_layout = QVBoxLayout()
        
        tables_layout = QHBoxLayout()
        tables_layout.setSpacing(0)

        self.frozen_table = QTableWidget()
        self.schedule_table = QTableWidget()
        
        tables_layout.addWidget(self.frozen_table)
        tables_layout.addWidget(self.schedule_table, 1) 
        
        schedule_container_layout.addLayout(tables_layout)

        self.frozen_table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.frozen_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.schedule_table.verticalHeader().setVisible(False)

        self.schedule_table.verticalScrollBar().valueChanged.connect(
            self.frozen_table.verticalScrollBar().setValue
        )
        self.frozen_table.verticalScrollBar().valueChanged.connect(
            self.schedule_table.verticalScrollBar().setValue
        )

        parent_layout.addWidget(self.create_group_box("🗓️ Vista Previa del Cronograma (PlanStaff.xlsx)", schedule_container_layout))

    def load_schedule_data(self):
        FROZEN_COLUMN_COUNT = 3

        df = excel.get_schedule_preview()
        if df.empty:
            return

        if df.shape[1] < FROZEN_COLUMN_COUNT:
            actual_frozen_count = df.shape[1]
        else:
            actual_frozen_count = FROZEN_COLUMN_COUNT

        headers = [str(col.strftime('%Y-%m-%d')) if isinstance(col, datetime) else str(col) for col in df.columns]

        self.frozen_table.setRowCount(df.shape[0])
        self.frozen_table.setColumnCount(actual_frozen_count)
        self.frozen_table.setHorizontalHeaderLabels(headers[:actual_frozen_count])

        self.schedule_table.setRowCount(df.shape[0])
        self.schedule_table.setColumnCount(df.shape[1] - actual_frozen_count)
        self.schedule_table.setHorizontalHeaderLabels(headers[actual_frozen_count:])

        for i, row in df.iterrows():
            for j, val in enumerate(row):
                item = QTableWidgetItem(str(val))
                if j < actual_frozen_count:
                    self.frozen_table.setItem(i, j, item)
                else:
                    col_index = j - actual_frozen_count
                    
                    val_str = str(val).upper()
                    # <<< CAMBIO: Lógica de colores actualizada para reconocer "ON NS"
                    if val_str == 'ON' or 'DAY SHIFT' in val_str or 'FA DAY' in val_str:
                        item.setBackground(QColor("#C6EFCE"))  # Verde
                    elif val_str == 'ON NS' or 'NIGHT SHIFT' in val_str or 'FA NIGHT' in val_str:
                        item.setBackground(QColor("#FFFF99"))  # Amarillo
                    elif val_str == 'OFF':
                        item.setBackground(QColor("#FFC7CE"))  # Rojo
                    elif 'LEAVE' in val_str:
                         item.setBackground(QColor("#D9D9D9")) # Gris para C. Leave

                    self.schedule_table.setItem(i, col_index, item)
        
        self.frozen_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.frozen_table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        
        self.schedule_table.resizeColumnsToContents()
        self.schedule_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.schedule_table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

        total_width = self.frozen_table.verticalHeader().width() + self.frozen_table.frameWidth() * 2
        for i in range(self.frozen_table.columnCount()):
            total_width += self.frozen_table.columnWidth(i)
        
        total_width += 2 
        self.frozen_table.setFixedWidth(total_width)

    def setup_registration_form_ui(self, parent_layout):
        form_layout = QGridLayout()
        
        self.username_input = QLineEdit()
        self.role_combo = QComboBox()
        self.badge_input = QLineEdit()
        
        self.shift_combo = QComboBox()
        self.shift_combo.addItems(["Day Shift", "Night Shift"])

        self.on_radio = QRadioButton("ON")
        self.off_radio = QRadioButton("OFF")
        self.none_radio = QRadioButton("No Marcar Días")
        self.on_radio.setChecked(True)

        self.on_radio.toggled.connect(self.update_shift_combo_state)
        self.off_radio.toggled.connect(self.update_shift_combo_state)
        self.none_radio.toggled.connect(self.update_shift_combo_state)

        self.start_date_edit = QDateEdit(QDate.currentDate())
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDisplayFormat("dd/MM/yyyy")
        self.end_date_edit = QDateEdit(QDate.currentDate().addDays(14))
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setDisplayFormat("dd/MM/yyyy")

        save_button = QPushButton("✅ Guardar Cambios en DB y Excel")
        save_button.clicked.connect(self.save_changes)

        form_layout.addWidget(QLabel("Nombre y Apellido:"), 0, 0)
        form_layout.addWidget(self.username_input, 0, 1)
        form_layout.addWidget(QLabel("Rol/Departamento:"), 1, 0)
        form_layout.addWidget(self.role_combo, 1, 1)
        form_layout.addWidget(QLabel("Badge (ID):"), 2, 0)
        form_layout.addWidget(self.badge_input, 2, 1)
        form_layout.addWidget(QLabel("Turno:"), 3, 0)
        form_layout.addWidget(self.shift_combo, 3, 1)
        
        status_layout = QHBoxLayout()
        status_layout.addWidget(self.on_radio)
        status_layout.addWidget(self.off_radio)
        status_layout.addWidget(self.none_radio)
        form_layout.addWidget(QLabel("Estado a Marcar:"), 4, 0)
        form_layout.addLayout(status_layout, 4, 1)

        form_layout.addWidget(QLabel("Fecha Inicio Período:"), 5, 0)
        form_layout.addWidget(self.start_date_edit, 5, 1)
        form_layout.addWidget(QLabel("Fecha Final Período:"), 6, 0)
        form_layout.addWidget(self.end_date_edit, 6, 1)
        form_layout.addWidget(save_button, 7, 0, 1, 2)
        
        parent_layout.addWidget(self.create_group_box("1. Registrar / Actualizar Empleado", form_layout), 1)
        self.update_shift_combo_state()

    def update_shift_combo_state(self):
        if self.on_radio.isChecked():
            self.shift_combo.setEnabled(True)
        else:
            self.shift_combo.setEnabled(False)

    def save_changes(self):
        username = self.username_input.text()
        badge = self.badge_input.text()
        role = self.role_combo.currentText()
        start_date = self.start_date_edit.date().toPyDate()
        end_date = self.end_date_edit.date().toPyDate()

        if not username or not badge:
            QMessageBox.warning(self, "Datos incompletos", "El Nombre y el Badge son campos obligatorios.")
            return
        if start_date > end_date:
            QMessageBox.warning(self, "Error de Fechas", "La fecha de inicio no puede ser posterior a la fecha final.")
            return
            
        schedule_status = "OFF"
        shift_type = None
        if self.on_radio.isChecked():
            schedule_status = "ON"
            shift_type = self.shift_combo.currentText()
        elif self.none_radio.isChecked():
            schedule_status = None

        db.add_operation(username, role, badge, start_date, end_date)
        
        success, message = excel.update_plan_staff_excel(username, role, badge, schedule_status, shift_type, start_date, end_date)

        if success:
            QMessageBox.information(self, "Éxito", f"¡Éxito! DB y Excel actualizados para {username}.")
        else:
            QMessageBox.warning(self, "Error de Excel", f"Se guardó en la DB, pero no se pudo actualizar PlanStaff.xlsx.\nCausa: {message}")
        
        self.refresh_ui_data()
    
    def setup_db_view_ui(self, parent_layout):
        db_view_layout = QVBoxLayout()
        self.db_table = QTableWidget()
        db_view_layout.addWidget(self.db_table)
        parent_layout.addWidget(self.create_group_box("Historial de Rotaciones (Vista de la DB)", db_view_layout), 2)

    def setup_report_generator_ui(self, parent_layout):
        report_layout = QHBoxLayout()
        report_layout.addWidget(QLabel("Fecha INICIO del reporte:"))
        self.report_start_date = QDateEdit(QDate.currentDate())
        self.report_start_date.setCalendarPopup(True)
        self.report_start_date.setDisplayFormat("dd/MM/yyyy")
        report_layout.addWidget(self.report_start_date)
        report_layout.addWidget(QLabel("Fecha FINAL del reporte:"))
        self.report_end_date = QDateEdit(QDate.currentDate().addDays(30))
        self.report_end_date.setCalendarPopup(True)
        self.report_end_date.setDisplayFormat("dd/MM/yyyy")
        report_layout.addWidget(self.report_end_date)
        report_button = QPushButton("🚀 Generar y Descargar Reporte")
        report_button.clicked.connect(self.generate_report)
        report_layout.addWidget(report_button)
        parent_layout.addWidget(self.create_group_box("2. Generar Reporte de Transporte", report_layout))

    def load_db_data(self):
        records = db.get_all_operations()
        headers = ["ID", "Nombre", "Rol", "Badge", "Fecha Inicio", "Fecha Fin"]
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

    def load_role_options(self):
        roles = excel.get_roles_from_excel()
        self.role_combo.clear()
        self.role_combo.addItems(roles)

    def refresh_ui_data(self):
        self.load_schedule_data()
        self.load_db_data()
        self.load_role_options()

    def generate_report(self):
        start_date = self.report_start_date.date().toPyDate()
        end_date = self.report_end_date.date().toPyDate()
        if start_date > end_date:
            QMessageBox.warning(self, "Error de Fechas", "La fecha de inicio del reporte no puede ser posterior a la fecha final.")
            return
        records = db.get_operations_for_report(start_date, end_date)
        if not records:
            QMessageBox.information(self, "Sin Datos", "No se encontraron rotaciones activas en el rango de fechas seleccionado.")
            return
        excel_data = excel.generate_transport_excel_from_db(records)
        default_filename = f"Transport_Request_{start_date}_to_{end_date}.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(self, "Guardar Reporte", default_filename, "Archivos de Excel (*.xlsx)")
        if file_path:
            try:
                with open(file_path, 'wb') as f:
                    f.write(excel_data)
                QMessageBox.information(self, "Éxito", f"Reporte guardado exitosamente en:\n{file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Error al Guardar", f"No se pudo guardar el archivo.\nError: {e}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())