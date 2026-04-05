# widgets/location_widget.py
# Extracted from main_window.py — LocationAdminWidget

import logging
from PyQt6.QtWidgets import (
    QWidget, QHBoxLayout, QVBoxLayout, QGridLayout, QLabel, QLineEdit,
    QComboBox, QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
    QGroupBox, QMessageBox,
)
from PyQt6.QtCore import pyqtSignal, QTimer, QSignalBlocker

import database_logic as db
from constants import DEBOUNCE_MS, Source

logger = logging.getLogger(__name__)


def _create_group_box(title: str, inner_layout) -> QGroupBox:
    box = QGroupBox(title)
    font = box.font()
    font.setBold(True)
    box.setFont(font)
    box.setLayout(inner_layout)
    return box


class LocationAdminWidget(QWidget):
    """
    Si scope_source es 'RGM' o 'Newmont' -> modo estandar (solo su empresa).
    Si scope_source es None -> modo Admin (todas, con filtro y selector de dueno).
    """

    locations_changed = pyqtSignal()

    def __init__(self, scope_source: str | None = None):
        super().__init__()
        self.scope_source = scope_source
        self.loc_id = None
        self._filter_state = {}
        self._debounce_timer = QTimer(self)
        self._debounce_timer.setSingleShot(True)
        self._debounce_timer.timeout.connect(self._reload_table)

        layout = QHBoxLayout(self)

        # --- Form ---
        form = QGridLayout()
        row = 0

        self.loc_input = QLineEdit()
        form.addWidget(QLabel("Location name:"), row, 0)
        form.addWidget(self.loc_input, row, 1)
        row += 1

        self.owner_combo = None
        if self.scope_source is None:
            self.owner_combo = QComboBox()
            self.owner_combo.addItems([Source.RGM.value, Source.NEWMONT.value])
            form.addWidget(QLabel("Owner (Source):"), row, 0)
            form.addWidget(self.owner_combo, row, 1)
            row += 1

        btn_save = QPushButton("Save")
        btn_del = QPushButton("Delete")
        btn_save.setProperty("variant", "primary")
        btn_del.setProperty("danger", True)
        h = QHBoxLayout()
        h.addWidget(btn_save)
        h.addWidget(btn_del)
        form.addLayout(h, row, 0, 1, 2)

        form_group = _create_group_box("Location", form)
        form_group.setFixedWidth(420)

        # --- Table ---
        table_panel = QWidget()
        table_box = QVBoxLayout(table_panel)

        controls_row = QHBoxLayout()

        self.loc_search_input = QLineEdit()
        self.loc_search_input.setPlaceholderText("Search location...")
        self.loc_search_input.textChanged.connect(self._request_refresh)
        controls_row.addWidget(self.loc_search_input, 1)

        self.filter_combo = None
        if self.scope_source is None:
            self.filter_combo = QComboBox()
            self.filter_combo.addItem("All Sources", None)
            self.filter_combo.addItem(Source.RGM.value, Source.RGM.value)
            self.filter_combo.addItem(Source.NEWMONT.value, Source.NEWMONT.value)
            self.filter_combo.currentIndexChanged.connect(self._request_refresh)
            controls_row.addWidget(QLabel("Source:"))
            controls_row.addWidget(self.filter_combo)

        self.sort_combo = QComboBox()
        self.sort_combo.addItems(["Sort by Name", "Sort by Source"])
        self.sort_combo.currentIndexChanged.connect(self._request_refresh)
        controls_row.addWidget(self.sort_combo)

        reset_btn = QPushButton("Reset")
        reset_btn.clicked.connect(self._reset_filters)
        controls_row.addWidget(reset_btn)

        table_box.addLayout(controls_row)

        self.loc_table = QTableWidget()
        self.loc_table.setAlternatingRowColors(True)
        table_box.addWidget(self.loc_table)

        table_group = _create_group_box("Locations", table_box)

        layout.addWidget(form_group)
        layout.addWidget(table_group)

        # Events
        btn_save.clicked.connect(self._save_loc)
        btn_del.clicked.connect(self._delete_loc)
        self.loc_table.itemClicked.connect(self._load_to_form)

        self._reset_filters()

    def _request_refresh(self):
        self._debounce_timer.start(DEBOUNCE_MS)

    def _reset_filters(self):
        with QSignalBlocker(self.loc_search_input), QSignalBlocker(self.sort_combo):
            self.loc_search_input.clear()
            self.sort_combo.setCurrentIndex(0)
            if self.filter_combo:
                with QSignalBlocker(self.filter_combo):
                    self.filter_combo.setCurrentIndex(0)
        self._reload_table()

    def _effective_filter_source(self) -> str | None:
        if self.scope_source is not None:
            return self.scope_source
        if self.filter_combo is None:
            return None
        return self.filter_combo.currentData()

    def _reload_table(self):
        src = self._effective_filter_source()
        text = self.loc_search_input.text()
        sort = "source" if self.sort_combo.currentIndex() == 1 else "name"

        self._filter_state["source"] = src
        self._filter_state["text"] = text
        self._filter_state["sort"] = sort

        rows = db.get_locations_filtered(source=src, text=text, sort_by=sort)

        if self.scope_source is None:
            headers = ["ID", "Source", "Location"]
        else:
            headers = ["ID", "Location"]
        self.loc_table.setRowCount(len(rows))
        self.loc_table.setColumnCount(len(headers))
        self.loc_table.setHorizontalHeaderLabels(headers)
        for r, row in enumerate(rows):
            self.loc_table.setItem(r, 0, QTableWidgetItem(str(row["id"])))
            if self.scope_source is None:
                self.loc_table.setItem(r, 1, QTableWidgetItem(row["source"]))
                self.loc_table.setItem(r, 2, QTableWidgetItem(row["pickup_location"]))
            else:
                self.loc_table.setItem(r, 1, QTableWidgetItem(row["pickup_location"]))
        self.loc_table.setColumnHidden(0, True)
        self.loc_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch
        )

    def _new_loc(self):
        self.loc_id = None
        self.loc_input.clear()
        if self.owner_combo is not None:
            self.owner_combo.setCurrentIndex(0)
        self.loc_table.clearSelection()

    def _save_loc(self):
        name = (self.loc_input.text() or "").strip()
        if not name:
            QMessageBox.warning(self, "Input Error", "Location name cannot be empty.")
            return

        if self.scope_source is None:
            dest_source = self.owner_combo.currentText() if self.owner_combo else Source.RGM.value
        else:
            dest_source = self.scope_source

        if self.loc_id:
            if self.scope_source is None:
                ok, msg = db.update_location_admin(self.loc_id, name, dest_source)
            else:
                ok, msg = db.update_location(self.loc_id, name, dest_source)
        else:
            ok, msg = db.create_location(name, dest_source)

        QMessageBox.information(self, "Location", msg)
        self._reload_table()
        self.locations_changed.emit()
        self._new_loc()

    def _delete_loc(self):
        if not self.loc_id:
            QMessageBox.warning(self, "Location", "Please select a row.")
            return

        reply = QMessageBox.question(
            self,
            "Confirm Deletion",
            f"Are you sure you want to delete this Location?\n\nThis will fail if users are currently assigned to it.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        if self.scope_source is None:
            ok, msg = db.delete_location_admin(self.loc_id)
        else:
            ok, msg = db.delete_location(self.loc_id, self.scope_source)

        if ok:
            QMessageBox.information(self, "Success", msg)
            self._reload_table()
            self.locations_changed.emit()
            self._new_loc()
        else:
            QMessageBox.warning(self, "Cannot Delete", msg)

    def _load_to_form(self, item):
        row = item.row()
        self.loc_id = int(self.loc_table.item(row, 0).text())
        if self.scope_source is None:
            self.loc_input.setText(self.loc_table.item(row, 2).text())
            if self.owner_combo is not None:
                src = self.loc_table.item(row, 1).text()
                idx = self.owner_combo.findText(src)
                if idx >= 0:
                    self.owner_combo.setCurrentIndex(idx)
        else:
            self.loc_input.setText(self.loc_table.item(row, 1).text())
