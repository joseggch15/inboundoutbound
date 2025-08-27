from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QLineEdit, QMessageBox, QDialog, 
    QProgressBar, QDialogButtonBox
)
from PyQt6.QtCore import Qt, QTimer

# --- CREDENCIALES DE ACCESO ---
CREDENTIALS = {
    "javier teheran": {
        "password": "123", 
        "role": "RGM", 
        "excel_file": "PlanStaffRGM.xlsx"
    },
    "miguel venegas": {
        "password": "456", 
        "role": "Newmont", 
        "excel_file": "PlanStaffNewmont.xlsx"
    }
}

class LoginWindow(QDialog):
    """Ventana de inicio de sesión."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Inicio de Sesión")
        self.setModal(True)
        self.user_role = None
        self.excel_file = None

        layout = QVBoxLayout(self)
        
        self.username_input = QLineEdit()
        self.username_input.setPlaceholderText("Usuario")
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("Contraseña")
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.check_login)
        buttons.rejected.connect(self.reject)

        layout.addWidget(QLabel("Por favor, ingrese sus credenciales:"))
        layout.addWidget(self.username_input)
        layout.addWidget(self.password_input)
        layout.addWidget(buttons)

    def check_login(self):
        username = self.username_input.text().lower()
        password = self.password_input.text()
        user_data = CREDENTIALS.get(username)
        
        if user_data and user_data["password"] == password:
            self.user_role = user_data["role"]
            self.excel_file = user_data["excel_file"]
            self.accept()
        else:
            QMessageBox.warning(self, "Error de Acceso", "Usuario o contraseña incorrectos.")
            self.password_input.clear()

class LoadingWindow(QWidget):
    """Ventana de carga que se muestra después de un login exitoso."""
    def __init__(self, role):
        super().__init__()
        self.role = role
        self.setWindowTitle("Cargando...")
        self.setFixedSize(400, 150)
        
        layout = QVBoxLayout(self)
        self.label = QLabel()
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        font = self.label.font()
        font.setPointSize(14)
        self.label.setFont(font)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        
        layout.addWidget(self.label)
        layout.addWidget(self.progress_bar)
        
        self.setup_ui_for_role()

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_progress)
        self.timer.start(30)
        self.progress_value = 0

    def setup_ui_for_role(self):
        if self.role == "RGM":
            self.label.setText("Cargando Módulo Plan Staff RGM...")
            self.setStyleSheet("background-color: #E6F3FF;")
        elif self.role == "Newmont":
            self.label.setText("Cargando Módulo de Reportes Newmont...")
            self.setStyleSheet("background-color: #E8F8F5;")
        else:
            self.label.setText("Cargando aplicación...")

    def update_progress(self):
        self.progress_value += 1
        self.progress_bar.setValue(self.progress_value)
        if self.progress_value >= 100:
            self.timer.stop()
            self.close()
