import sys
from PyQt6.QtWidgets import QApplication, QDialog, QWidget, QVBoxLayout, QPushButton, QLabel
from PyQt6.QtCore import Qt
from datetime import datetime

# Importar las ventanas
from main_window import MainWindow
from ui_login import LoginWindow, LoadingWindow

class LauncherWindow(QWidget):
    """
    Ventana inicial que muestra un botón para iniciar sesión.
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Bienvenido al Gestor de Operaciones")
        self.setGeometry(400, 400, 400, 200)
        self.main_app_window = None

        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        title = QLabel("Gestor de Transporte y Operaciones")
        font = title.font()
        font.setPointSize(16)
        font.setBold(True)
        title.setFont(font)

        login_button = QPushButton("▶️ Iniciar Sesión")
        login_button.setFixedSize(200, 50)
        font = login_button.font()
        font.setPointSize(12)
        login_button.setFont(font)
        login_button.clicked.connect(self.start_login_process)

        layout.addWidget(title)
        layout.addSpacing(20)
        layout.addWidget(login_button, alignment=Qt.AlignmentFlag.AlignCenter)

        self.setLayout(layout)

    def start_login_process(self):
        """
        Maneja el flujo completo de login y pasa la información del usuario
        y el archivo excel a la ventana principal.
        """
        login_dialog = LoginWindow(self)
        
        if login_dialog.exec() == QDialog.DialogCode.Accepted:
            user_role = login_dialog.user_role
            excel_file = login_dialog.excel_file
            
            self.hide()

            loading_screen = LoadingWindow(role=user_role)
            loading_screen.show()
            
            start_time = datetime.now()
            while (datetime.now() - start_time).total_seconds() < 3:
                QApplication.instance().processEvents()

            loading_screen.close()

            self.main_app_window = MainWindow(user_role=user_role, excel_file=excel_file)
            self.main_app_window.logout_signal.connect(self.handle_logout)
            self.main_app_window.show()

    def handle_logout(self):
        """
        Muestra esta ventana de lanzador de nuevo al cerrar sesión.
        """
        self.main_app_window = None
        self.show()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    launcher = LauncherWindow()
    launcher.show()
    sys.exit(app.exec())
