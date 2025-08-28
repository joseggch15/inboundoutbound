from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QLineEdit, QMessageBox, QDialog,
    QProgressBar, QDialogButtonBox
)
from PyQt6.QtCore import Qt, QTimer

# --- ACCESS CREDENTIALS ---
CREDENTIALS = {
    "javierteheran": {
        "password": "123",
        "role": "RGM",
        "excel_file": "PlanStaffRGM.xlsx"
    },
    "miguelvenegas": {
        "password": "456",
        "role": "Newmont",
        "excel_file": "PlanStaffNewmont.xlsx"
    }
}

class LoginWindow(QDialog):
    """Sign-in dialog."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Sign In")
        self.setModal(True)

        # Valores que leerá main.py tras un login exitoso
        self.username = None
        self.user_role = None
        self.excel_file = None

        layout = QVBoxLayout(self)

        self.username_input = QLineEdit()
        self.username_input.setPlaceholderText("Username")
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("Password")
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.check_login)
        buttons.rejected.connect(self.reject)
        # Forzar etiquetas en inglés
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Sign in")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancel")

        layout.addWidget(QLabel("Please enter your credentials:"))
        layout.addWidget(self.username_input)
        layout.addWidget(self.password_input)
        layout.addWidget(buttons)

    def check_login(self):
        # Guardamos el texto tal como lo escribió el usuario para mostrarlo en la UI
        typed_username = self.username_input.text().strip()
        # Para validar credenciales, usamos lower()
        username_lookup = typed_username.lower()
        password = self.password_input.text()
        user_data = CREDENTIALS.get(username_lookup)

        if user_data and user_data["password"] == password:
            # Datos que usará el resto de la app
            self.username = typed_username if typed_username else username_lookup
            self.user_role = user_data["role"]
            self.excel_file = user_data["excel_file"]
            self.accept()
        else:
            QMessageBox.warning(self, "Login Error", "Invalid username or password.")
            self.password_input.clear()

class LoadingWindow(QWidget):
    """Loading splash shown after a successful login."""
    def __init__(self, role):
        super().__init__()
        self.role = role
        self.setWindowTitle("Loading...")
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
            self.label.setText("Loading RGM Plan Staff Module...")
            self.setStyleSheet("background-color: #E6F3FF;")
        elif self.role == "Newmont":
            self.label.setText("Loading Newmont Reports Module...")
            self.setStyleSheet("background-color: #E8F8F5;")
        else:
            self.label.setText("Loading application...")

    def update_progress(self):
        self.progress_value += 1
        self.progress_bar.setValue(self.progress_value)
        if self.progress_value >= 100:
            self.timer.stop()
            self.close()
