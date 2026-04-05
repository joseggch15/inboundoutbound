# ui_login.py — Login screen with credentials stored in database (hashed passwords).
# No more hardcoded plaintext passwords.

import logging

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QLineEdit, QMessageBox, QDialog,
    QProgressBar, QDialogButtonBox, QFrame, QApplication
)
from PyQt6.QtCore import Qt, QTimer

import database_logic as db

logger = logging.getLogger(__name__)


class LoginWindow(QDialog):
    """Sign-in dialog. Authenticates against the credentials table in the DB."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Log In")
        self.setModal(True)

        # Values read by main.py after a successful login
        self.username = None
        self.user_role = None
        self.excel_file = None
        self.can_manage_shift_types = False

        layout = QVBoxLayout(self)

        self.username_input = QLineEdit()
        self.username_input.setPlaceholderText("Username")
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("Password")
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.check_login)
        buttons.rejected.connect(self.reject)

        # Button labels and theme variants
        ok_btn = buttons.button(QDialogButtonBox.StandardButton.Ok)
        cancel_btn = buttons.button(QDialogButtonBox.StandardButton.Cancel)
        if ok_btn:
            ok_btn.setText("Log in")
            ok_btn.setProperty("variant", "primary")
        if cancel_btn:
            cancel_btn.setText("Cancel")
            cancel_btn.setProperty("variant", "text")

        layout.addWidget(QLabel("Please enter your credentials:"))
        layout.addWidget(self.username_input)
        layout.addWidget(self.password_input)
        layout.addWidget(buttons)

    def check_login(self):
        typed_username = self.username_input.text().strip()
        password = self.password_input.text()

        # Ensure DB and credentials table exist
        try:
            db.setup_database()
        except Exception as e:
            logger.error("Failed to setup database during login: %s", e)

        user_data = db.authenticate_user(typed_username, password)

        if user_data:
            self.username = typed_username if typed_username else user_data["username"]
            self.user_role = user_data["role"]
            self.excel_file = user_data["excel_file"]
            self.can_manage_shift_types = user_data["can_manage_shift_types"]
            logger.info("User '%s' logged in as %s", self.username, self.user_role)
            self.accept()
        else:
            QMessageBox.warning(self, "Login Error", "Invalid username or password.")
            self.password_input.clear()


class LoadingWindow(QDialog):
    """
    Loading splash shown after a successful login.
    Controlled externally by real application events.
    """
    def __init__(self, role, parent=None):
        super().__init__(parent)
        self.role = role
        self.setWindowTitle("Loading...")
        self.setFixedSize(420, 180)

        # Flags for visual stability
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.Dialog |
            Qt.WindowType.WindowStaysOnTopHint
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setModal(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)

        # Frame with styled border
        self.frame = QFrame()
        self.frame.setStyleSheet("""
            QFrame {
                background-color: #FFFFFF;
                border: 1px solid #CFD8DC;
                border-radius: 12px;
            }
            QLabel { color: #374151; border: none; }
        """)

        inner_layout = QVBoxLayout(self.frame)
        inner_layout.setContentsMargins(20, 20, 20, 20)

        self.label = QLabel("Initializing System...")
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        font = self.label.font()
        font.setPointSize(13)
        font.setBold(True)
        self.label.setFont(font)

        self.status_label = QLabel("Please wait...")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet("color: #546E7A; font-size: 11px; margin-top: 5px;")

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(8)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                background-color: #ECEFF1;
                border-radius: 4px;
            }
            QProgressBar::chunk {
                background-color: #0288D1;
                border-radius: 4px;
            }
        """)

        inner_layout.addStretch()
        inner_layout.addWidget(self.label)
        inner_layout.addSpacing(15)
        inner_layout.addWidget(self.progress_bar)
        inner_layout.addWidget(self.status_label)
        inner_layout.addStretch()

        layout.addWidget(self.frame)
        self.setup_ui_for_role()

    def setup_ui_for_role(self):
        from constants import Source
        if self.role == Source.RGM.value:
            self.label.setText("Transport Manager: RGM")
        elif self.role == Source.NEWMONT.value:
            self.label.setText("Transport Manager: Newmont")
        elif self.role == Source.ADMINISTRATOR.value:
            self.label.setText("Administrator Console")

    def update_progress(self, value, message):
        self.progress_bar.setValue(value)
        if message:
            self.status_label.setText(message)
        # Force immediate repaint to avoid visual "freezing"
        QApplication.processEvents()
