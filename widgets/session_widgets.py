# session_widgets.py
# Extracted from main_window.py — Multi-user session awareness widgets.

import sqlite3
import socket
import logging

from PyQt6.QtWidgets import QWidget, QHBoxLayout, QLabel
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QColor, QFont, QPainter, QPen

import database_logic as db
from constants import AVATAR_COLORS

logger = logging.getLogger(__name__)


class UserAvatarWidget(QWidget):
    """Circular badge showing user initials, like Office co-authoring indicators."""

    def __init__(self, username: str, color: str, is_self: bool = False,
                 is_editor: bool = False, parent=None):
        super().__init__(parent)
        self._username = username
        self._color = QColor(color)
        self._is_self = is_self
        self._is_editor = is_editor
        self._initials = self._get_initials(username)
        self.setFixedSize(36, 36)
        self.setToolTip(
            f"{username} {'(You)' if is_self else ''}"
            f" — {'Editing' if is_editor else 'Viewing'}"
        )

    @staticmethod
    def _get_initials(name: str) -> str:
        parts = name.strip().split()
        if len(parts) >= 2:
            return (parts[0][0] + parts[-1][0]).upper()
        return name[:2].upper() if name else "?"

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect = self.rect().adjusted(2, 2, -2, -2)

        # Circle fill
        painter.setBrush(self._color)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawEllipse(rect)

        # Editor gets a green border ring
        if self._is_editor:
            pen = QPen(QColor("#4CAF50"), 2)
            painter.setPen(pen)
            painter.setBrush(Qt.BrushStyle.NoBrush)
            painter.drawEllipse(rect)

        # Initials text
        painter.setPen(QColor("#FFFFFF"))
        font = QFont()
        font.setPointSize(10)
        font.setBold(True)
        painter.setFont(font)
        painter.drawText(rect, Qt.AlignmentFlag.AlignCenter, self._initials)
        painter.end()


class SessionBarWidget(QWidget):
    """
    Horizontal bar that shows circular avatars for all active sessions.
    Refreshes periodically by polling the active_sessions table.
    """

    def __init__(self, current_username: str, parent=None):
        super().__init__(parent)
        self._current_username = current_username
        self._avatar_layout = QHBoxLayout(self)
        self._avatar_layout.setContentsMargins(0, 0, 0, 0)
        self._avatar_layout.setSpacing(4)

        # Mode label (Editing / View Only)
        self._mode_label = QLabel()
        font = self._mode_label.font()
        font.setBold(True)
        font.setPointSize(9)
        self._mode_label.setFont(font)

        self._avatars_container = QWidget()
        self._avatars_layout = QHBoxLayout(self._avatars_container)
        self._avatars_layout.setContentsMargins(0, 0, 0, 0)
        self._avatars_layout.setSpacing(-8)  # overlap like Office

        self._avatar_layout.addWidget(self._avatars_container)
        self._avatar_layout.addWidget(self._mode_label)

        # Refresh timer (every 15 seconds)
        self._refresh_timer = QTimer(self)
        self._refresh_timer.timeout.connect(self.refresh)
        self._refresh_timer.start(15_000)

        self._is_editor = True  # default until first refresh

    @property
    def is_editor(self) -> bool:
        return self._is_editor

    def refresh(self):
        """Poll active_sessions and rebuild avatar widgets."""
        try:
            sessions = db.get_all_active_sessions()
        except (sqlite3.Error, OSError) as e:
            logger.debug("Could not fetch active sessions: %s", e)
            return

        # Clear existing avatars
        while self._avatars_layout.count():
            item = self._avatars_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        if not sessions:
            self._is_editor = True
            self._mode_label.setText("Editing")
            self._mode_label.setStyleSheet("color: #388E3C; padding-left: 8px;")
            return

        # First session (oldest login_time) is the editor
        editor_username = sessions[0]["username"]
        editor_machine = sessions[0]["machine_name"]

        current_machine = socket.gethostname()
        self._is_editor = (
            editor_username == self._current_username
            and editor_machine == current_machine
        )

        for i, session in enumerate(sessions):
            color = AVATAR_COLORS[i % len(AVATAR_COLORS)]
            is_self = (
                session["username"] == self._current_username
                and session["machine_name"] == current_machine
            )
            is_session_editor = (i == 0)  # first session = editor
            avatar = UserAvatarWidget(
                session["username"], color, is_self=is_self,
                is_editor=is_session_editor
            )
            self._avatars_layout.addWidget(avatar)

        if self._is_editor:
            self._mode_label.setText("Editing")
            self._mode_label.setStyleSheet("color: #388E3C; padding-left: 8px;")
        else:
            self._mode_label.setText("View Only")
            self._mode_label.setStyleSheet("color: #D32F2F; padding-left: 8px;")

    def stop(self):
        """Stop the refresh timer (call on close/logout)."""
        self._refresh_timer.stop()
