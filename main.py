import sys
from PyQt6.QtWidgets import QApplication, QDialog, QWidget, QVBoxLayout, QPushButton, QLabel
from PyQt6.QtCore import Qt
from datetime import datetime

# Ventanas principales
from main_window import MainWindow, AdminMainWindow
from ui_login import LoginWindow, LoadingWindow


class LauncherWindow(QWidget):
    """
    Initial window with a single Sign In button.
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Welcome to Operations Manager")
        self.setGeometry(400, 400, 400, 200)
        self.main_app_window = None

        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        title = QLabel("Transport & Operations Manager")
        font = title.font()
        font.setPointSize(16)
        font.setBold(True)
        title.setFont(font)

        login_button = QPushButton("▶️ Sign In")
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
        Handles the complete sign-in flow and opens the main window
        depending on the user's role.
        """
        login_dialog = LoginWindow(self)

        if login_dialog.exec() == QDialog.DialogCode.Accepted:
            # Data returned by the login dialog
            user_role = login_dialog.user_role
            excel_file = login_dialog.excel_file
            logged_username = login_dialog.username
            can_manage_shift_types = getattr(login_dialog, "can_manage_shift_types", False)  # <— NUEVO

            self.hide()

            loading_screen = LoadingWindow(role=user_role)
            loading_screen.show()

            start_time = datetime.now()
            # Small non-blocking splash (mantiene comportamiento actual)
            while (datetime.now() - start_time).total_seconds() < 3:
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
                    can_manage_shift_types=can_manage_shift_types  # <— NUEVO
                )

            self.main_app_window.logout_signal.connect(self.handle_logout)
            self.main_app_window.show()

    def handle_logout(self):
        """Shows the launcher again after signing out."""
        self.main_app_window = None
        self.show()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    launcher = LauncherWindow()
    launcher.show()
    sys.exit(app.exec())
