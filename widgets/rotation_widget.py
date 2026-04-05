# widgets/rotation_widget.py
# Extracted from main_window.py — RotationHistoryWidget

import logging

from PyQt6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QComboBox,
    QDateEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QAbstractItemView,
    QCheckBox,
)
from PyQt6.QtCore import QDate, QTimer, QSignalBlocker

import database_logic as db
from constants import DEBOUNCE_MS

logger = logging.getLogger(__name__)


class RotationHistoryWidget(QWidget):
    def __init__(self, created_by: str | None = None):
        """
        Si created_by es None => modo 'admin' (ve todo).
        En caso contrario, solo muestra operaciones creadas por ese usuario.
        """
        super().__init__()
        self._created_by = (created_by or "").strip() or None
        self._filter_state = {}
        self._debounce_timer = QTimer(self)
        self._debounce_timer.setSingleShot(True)
        self._debounce_timer.timeout.connect(self.refresh_data)

        layout = QVBoxLayout(self)

        # -- Filter Panel --
        filter_layout = QHBoxLayout()
        filter_layout.setContentsMargins(0, 0, 0, 8)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search Name/Badge/Role...")
        self.search_input.textChanged.connect(self._request_refresh)

        self.role_combo = QComboBox()
        self.role_combo.currentIndexChanged.connect(self._request_refresh)

        self.date_from = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_from.setDisplayFormat("yyyy-MM-dd")
        self.date_from.dateChanged.connect(self._request_refresh)

        self.date_to = QDateEdit()
        self.date_to.setCalendarPopup(True)
        self.date_to.setDisplayFormat("yyyy-MM-dd")
        self.date_to.dateChanged.connect(self._request_refresh)

        self.active_today_check = QCheckBox("Active today")
        self.active_today_check.toggled.connect(self._toggle_active_today)

        self.sort_combo = QComboBox()
        self.sort_combo.addItems(["Start Date (desc)", "Name (asc)"])
        self.sort_combo.currentIndexChanged.connect(self._request_refresh)

        reset_button = QPushButton("Reset")
        reset_button.clicked.connect(self.reset_filters)

        filter_layout.addWidget(self.search_input, 2)
        filter_layout.addWidget(QLabel("Role:"))
        filter_layout.addWidget(self.role_combo, 1)
        filter_layout.addWidget(QLabel("Date Range:"))
        filter_layout.addWidget(self.date_from)
        filter_layout.addWidget(QLabel("–"))
        filter_layout.addWidget(self.date_to)
        filter_layout.addWidget(self.active_today_check)
        filter_layout.addStretch()
        filter_layout.addWidget(QLabel("Sort by:"))
        filter_layout.addWidget(self.sort_combo)
        filter_layout.addWidget(reset_button)

        layout.addLayout(filter_layout)

        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        layout.addWidget(self.table)

        self.reset_filters()  # To initialize and load data

    def _request_refresh(self):
        self._debounce_timer.start(DEBOUNCE_MS)

    def _toggle_active_today(self, checked):
        self.date_from.setEnabled(not checked)
        self.date_to.setEnabled(not checked)
        if checked:
            today = QDate.currentDate()
            self.date_from.setDate(today)
            self.date_to.setDate(today)
        self._request_refresh()

    def _populate_role_filter(self):
        self.role_combo.blockSignals(True)
        current_role = self.role_combo.currentText()
        self.role_combo.clear()
        self.role_combo.addItem("All Roles", None)
        # roles solo de operaciones creadas por el usuario (si aplica)
        all_records = db.get_operations_filtered(created_by=self._created_by)
        roles = sorted(list(set(r["role"] for r in all_records if r.get("role"))))
        self.role_combo.addItems(roles)

        idx = self.role_combo.findText(current_role)
        if idx != -1:
            self.role_combo.setCurrentIndex(idx)
        self.role_combo.blockSignals(False)

    def reset_filters(self):
        with QSignalBlocker(self.search_input), QSignalBlocker(
            self.role_combo
        ), QSignalBlocker(self.date_from), QSignalBlocker(self.date_to), QSignalBlocker(
            self.active_today_check
        ), QSignalBlocker(
            self.sort_combo
        ):
            self.search_input.clear()
            self._populate_role_filter()
            self.role_combo.setCurrentIndex(0)
            self.date_from.setDate(QDate(2000, 1, 1))
            self.date_to.setDate(QDate.currentDate().addYears(5))
            self.active_today_check.setChecked(False)
            self.sort_combo.setCurrentIndex(0)
        self.refresh_data()

    def refresh_data(self):
        # Persist filter state
        self._filter_state["text"] = self.search_input.text()
        self._filter_state["role"] = (
            self.role_combo.currentText()
            if self.role_combo.currentIndex() > 0
            else None
        )
        self._filter_state["sort"] = (
            "name_asc" if self.sort_combo.currentIndex() == 1 else "start_date_desc"
        )

        d_from = self.date_from.date().toPyDate()
        d_to = self.date_to.date().toPyDate()

        records = db.get_operations_filtered(
            text=self._filter_state["text"],
            role=self._filter_state["role"],
            d_from=d_from,
            d_to=d_to,
            sort_by=self._filter_state["sort"],
            created_by=self._created_by,  # NUEVO: solo mis rotaciones
        )

        headers = ["Name", "Role", "Badge", "Start Date", "End Date", "Created By"]
        self.table.setRowCount(len(records))
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)

        for row_idx, record in enumerate(records):
            self.table.setItem(row_idx, 0, QTableWidgetItem(record["username"]))
            self.table.setItem(row_idx, 1, QTableWidgetItem(record["role"]))
            self.table.setItem(row_idx, 2, QTableWidgetItem(record["badge"]))
            self.table.setItem(row_idx, 3, QTableWidgetItem(record["start_date"]))
            self.table.setItem(row_idx, 4, QTableWidgetItem(record["end_date"]))
            self.table.setItem(
                row_idx, 5, QTableWidgetItem(record.get("created_by", ""))
            )

        self.table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch
        )
