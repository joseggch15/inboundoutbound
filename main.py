# main.py  (PyQt6, fix HiDPI attrs)
import sys
from PyQt6.QtWidgets import QApplication, QDialog, QWidget, QVBoxLayout, QPushButton, QLabel
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
from datetime import datetime

# Ventanas principales
from main_window import MainWindow, AdminMainWindow
from ui_login import LoginWindow, LoadingWindow

# Tema UI
from ui.theme import apply_app_theme


class LauncherWindow(QWidget):
    """
    Initial window with a single Sign In button.
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Inbound - Outbound PLG")
        self.setMinimumSize(420, 220)
        self.main_app_window = None

        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setSpacing(12)
        layout.setContentsMargins(24, 24, 24, 24)

        title = QLabel("Transport & Operations Manager")
        font = title.font(); font.setPointSize(20); font.setBold(True)
        title.setFont(font)

        login_button = QPushButton("Sign In")
        login_button.setFixedSize(220, 48)
        login_button.setProperty("variant", "primary")  # << estilo coherente

        layout.addWidget(title, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addSpacing(20)
        layout.addWidget(login_button, alignment=Qt.AlignmentFlag.AlignCenter)

        self.setLayout(layout)
        login_button.clicked.connect(self.start_login_process)

    def start_login_process(self):
        """
        Handles the complete sign-in flow and opens the main window
        depending on the user's role.
        """
        login_dialog = LoginWindow(self)

        if login_dialog.exec() == QDialog.DialogCode.Accepted:
            # Data returned by the login dialog
            user_role = login_dialog.user_role
            excel_file = login_dialog.excel_file
            logged_username = login_dialog.username
            can_manage_shift_types = getattr(login_dialog, "can_manage_shift_types", False)

            self.hide()

            loading_screen = LoadingWindow(role=user_role)
            loading_screen.show()

            start_time = datetime.now()
            # Small non-blocking splash
            while (datetime.now() - start_time).total_seconds() < 1.8:
                QApplication.instance().processEvents()

            loading_screen.close()

            # Open the appropriate main window
            if user_role == "Administrator":
                self.main_app_window = AdminMainWindow(
                    logged_username=logged_username,
                    rgm_excel="PlanStaffRGM.xlsx",
                    newmont_excel="PlanStaffNewmont.xlsx"
                )
            else:
                self.main_app_window = MainWindow(
                    user_role=user_role,
                    excel_file=excel_file,
                    logged_username=logged_username,
                    can_manage_shift_types=can_manage_shift_types
                )

            self.main_app_window.logout_signal.connect(self.handle_logout)
            self.main_app_window.show()

    def handle_logout(self):
        """Shows the launcher again after signing out."""
        self.main_app_window = None
        self.show()


if __name__ == '__main__':
    # Qt6 ya maneja HiDPI por defecto; solo ajustamos la política de redondeo (si está disponible)
    try:
        QApplication.setHighDpiScaleFactorRoundingPolicy(
            Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
        )
    except Exception:
        pass

    app = QApplication(sys.argv)
    app.setApplicationName("Inbound - Outbound PLG")

    # Aplicar tema global
    apply_app_theme(app)

    launcher = LauncherWindow()
    launcher.show()
    sys.exit(app.exec())
