# main.py  (PyQt6, sin AA_UseHighDpiPixmaps y con splash no bloqueante)
import sys
from PyQt6.QtWidgets import QApplication, QDialog, QWidget, QVBoxLayout, QPushButton, QLabel
from PyQt6.QtCore import Qt, QTimer
from datetime import datetime

from main_window import MainWindow, AdminMainWindow
from ui_login import LoginWindow, LoadingWindow
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
        self._login_payload = None
        self._loading = None

        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setSpacing(12)
        layout.setContentsMargins(24, 24, 24, 24)

        title = QLabel("Transport & Operations Manager")
        font = title.font(); font.setPointSize(20); font.setBold(True)
        title.setFont(font)

        login_button = QPushButton("Sign In")
        login_button.setFixedSize(220, 48)
        login_button.setProperty("variant", "primary")

        layout.addWidget(title, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addSpacing(20)
        layout.addWidget(login_button, alignment=Qt.AlignmentFlag.AlignCenter)

        self.setLayout(layout)
        login_button.clicked.connect(self.start_login_process)

    def start_login_process(self):
        """
        Handles the sign-in flow and opens the main window depending on the user's role.
        """
        login_dialog = LoginWindow(self)
        if login_dialog.exec() == QDialog.DialogCode.Accepted:
            # Cache login data
            self._login_payload = {
                "user_role": login_dialog.user_role,
                "excel_file": login_dialog.excel_file,
                "logged_username": login_dialog.username,
                "can_manage_shift_types": bool(getattr(login_dialog, "can_manage_shift_types", False)),
            }

            self.hide()

            # Non-blocking splash (no bucles con processEvents)
            self._loading = LoadingWindow(role=self._login_payload["user_role"])
            self._loading.show()
            QTimer.singleShot(1200, self._open_main_after_loading)  # ~1.2s de splash

    def _open_main_after_loading(self):
        if self._loading:
            self._loading.close()
            self._loading = None

        p = self._login_payload or {}
        role = p.get("user_role")
        if role == "Administrator":
            self.main_app_window = AdminMainWindow(
                logged_username=p.get("logged_username") or "",
                rgm_excel="PlanStaffRGM.xlsx",
                newmont_excel="PlanStaffNewmont.xlsx",
            )
        else:
            self.main_app_window = MainWindow(
                user_role=role,
                excel_file=p.get("excel_file") or "",
                logged_username=p.get("logged_username") or "",
                can_manage_shift_types=p.get("can_manage_shift_types", False),
            )

        self.main_app_window.logout_signal.connect(self.handle_logout)
        self.main_app_window.show()


    def handle_logout(self):
        """Shows the launcher again after signing out."""
        self.main_app_window = None
        self.show()
        


if __name__ == '__main__':
    # Qt6 ya maneja HiDPI por defecto; solo ajustamos la política de redondeo si está disponible
    try:
        QApplication.setHighDpiScaleFactorRoundingPolicy(
            Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
        )
    except Exception:
        pass

    app = QApplication(sys.argv)
    app.setApplicationName("Inbound - Outbound PLG")

    apply_app_theme(app)

    launcher = LauncherWindow()
    launcher.show()
    sys.exit(app.exec())
