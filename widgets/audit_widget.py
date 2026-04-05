# widgets/audit_widget.py
# Extracted from main_window.py — AuditLogWidget

import logging
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QPushButton, QTableWidget, QTableWidgetItem,
    QHeaderView,
)
from PyQt6.QtCore import Qt

import database_logic as db

logger = logging.getLogger(__name__)


class AuditLogWidget(QWidget):
    def __init__(self, source: str | None):
        super().__init__()
        self.source = source
        layout = QVBoxLayout(self)
        self.audit_table = QTableWidget()
        layout.addWidget(self.audit_table)

        refresh_btn = QPushButton("Refresh")
        refresh_btn.setProperty("variant", "secondary")
        refresh_btn.clicked.connect(self.load_audit_log_data)
        layout.addWidget(refresh_btn, alignment=Qt.AlignmentFlag.AlignRight)

        self.load_audit_log_data()

    def load_audit_log_data(self):
        events = db.get_audit_log(source=self.source)
        headers = ["Timestamp", "User", "Source", "Action", "Detail"]
        self.audit_table.setRowCount(len(events))
        self.audit_table.setColumnCount(len(headers))
        self.audit_table.setHorizontalHeaderLabels(headers)

        for r, ev in enumerate(events):
            self.audit_table.setItem(r, 0, QTableWidgetItem(ev.get("ts", "")))
            self.audit_table.setItem(r, 1, QTableWidgetItem(ev.get("username", "")))
            self.audit_table.setItem(r, 2, QTableWidgetItem(ev.get("source", "")))
            self.audit_table.setItem(r, 3, QTableWidgetItem(ev.get("action_type", "")))
            self.audit_table.setItem(r, 4, QTableWidgetItem(ev.get("detail", "")))

        self.audit_table.setAlternatingRowColors(True)
        self.audit_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch
        )
