# role_widget.py
# Extracted from main_window.py — Role administration widget.

import logging

from PyQt6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QLabel,
    QLineEdit,
    QComboBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QMessageBox,
    QAbstractItemView,
)
from PyQt6.QtCore import pyqtSignal, QTimer

import database_logic as db
from constants import DEBOUNCE_MS
from widgets.common_widgets import create_group_box

logger = logging.getLogger(__name__)


class RoleAdminWidget(QWidget):
    """
    Manages the master list of Roles.
    Similar to LocationAdminWidget but for 'roles' table.
    """
    roles_changed = pyqtSignal()

    def __init__(self, scope_source: str | None = None):
        super().__init__()
        self.scope_source = scope_source
        self.role_id = None
        self._debounce_timer = QTimer(self)
        self._debounce_timer.setSingleShot(True)
        self._debounce_timer.timeout.connect(self._reload_table)

        layout = QHBoxLayout(self)

        # --- Form ---
        form = QGridLayout()
        row = 0
        self.role_input = QLineEdit()
        form.addWidget(QLabel("Role / Dept Name:"), row, 0)
        form.addWidget(self.role_input, row, 1)
        row += 1

        # Admin Logic (Optional owner selection)
        self.owner_combo = None
        if self.scope_source is None:
            self.owner_combo = QComboBox()
            self.owner_combo.addItems(["RGM", "Newmont"])
            form.addWidget(QLabel("Owner:"), row, 0)
            form.addWidget(self.owner_combo, row, 1)
            row += 1

        btn_new = QPushButton("\u2728 New")
        btn_save = QPushButton("\U0001f4be Save")
        btn_del = QPushButton("\u274c Delete")
        btn_save.setProperty("variant", "primary")
        btn_del.setProperty("danger", True)

        h = QHBoxLayout()
        #h.addWidget(btn_new);
        h.addWidget(btn_save); h.addWidget(btn_del)
        form.addLayout(h, row, 0, 1, 2)

        form_group = create_group_box("Manage Roles", form)
        form_group.setFixedWidth(400)

        # --- Table ---
        table_panel = QWidget()
        t_layout = QVBoxLayout(table_panel)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search role...")
        self.search_input.textChanged.connect(lambda: self._debounce_timer.start(DEBOUNCE_MS))
        t_layout.addWidget(self.search_input)

        self.role_table = QTableWidget()
        self.role_table.setAlternatingRowColors(True)
        self.role_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.role_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.role_table.itemClicked.connect(self._load_to_form)
        t_layout.addWidget(self.role_table)

        table_group = create_group_box("Existing Roles", t_layout)

        layout.addWidget(form_group)
        layout.addWidget(table_group)

        # Connects
        btn_new.clicked.connect(self._new_role)
        btn_save.clicked.connect(self._save_role)
        btn_del.clicked.connect(self._delete_role)

        self._reload_table()

    def _reload_table(self):
        text = self.search_input.text()
        rows = db.get_roles_filtered(self.scope_source, text)

        headers = ["ID", "Source", "Role Name"] if self.scope_source is None else ["ID", "Role Name"]
        self.role_table.setRowCount(len(rows))
        self.role_table.setColumnCount(len(headers))
        self.role_table.setHorizontalHeaderLabels(headers)

        for r, row in enumerate(rows):
            self.role_table.setItem(r, 0, QTableWidgetItem(str(row["id"])))
            if self.scope_source is None:
                self.role_table.setItem(r, 1, QTableWidgetItem(row["source"]))
                self.role_table.setItem(r, 2, QTableWidgetItem(row["name"]))
            else:
                self.role_table.setItem(r, 1, QTableWidgetItem(row["name"]))

        self.role_table.setColumnHidden(0, True)
        self.role_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

    def _new_role(self):
        self.role_id = None
        self.role_input.clear()
        self.role_table.clearSelection()

    def _save_role(self):
        name = self.role_input.text().strip()
        if not name: return

        src = self.scope_source
        if src is None and self.owner_combo:
            src = self.owner_combo.currentText()

        if self.role_id:
            ok, msg = db.update_role(self.role_id, name, src)
        else:
            ok, msg = db.create_role(name, src)

        QMessageBox.information(self, "Role", msg)
        self._reload_table()
        self.roles_changed.emit()
        self._new_role()

    def _delete_role(self):
        if not self.role_id:
            return

        # Determinar el 'source' (RGM o Newmont)
        src = self.scope_source
        if src is None and self.owner_combo: # Caso especial para Admin
             # Intenta leer la columna oculta o el texto de la tabla
             current_row = self.role_table.currentRow()
             if current_row >= 0:
                 src = self.role_table.item(current_row, 1).text()

        # 1. Confirmación de seguridad (UI)
        confirm = QMessageBox.question(
            self,
            "Confirm",
            "Delete this role?\n\nThis will fail if currently assigned to users."
        )

        if confirm == QMessageBox.StandardButton.Yes:
            # 2. Lógica de Base de Datos
            # Llamamos a la nueva función protegida en db
            ok, msg = db.delete_role(self.role_id, src)

            # 3. Feedback
            if ok:
                QMessageBox.information(self, "Role", msg)
                self._reload_table()
                self.roles_changed.emit()
                self._new_role()
            else:
                # Aquí mostramos el aviso si hay usuarios con este rol
                QMessageBox.warning(self, "Cannot Delete", msg)

    def _load_to_form(self, item):
        row = item.row()
        self.role_id = int(self.role_table.item(row, 0).text())
        if self.scope_source is None:
            self.role_input.setText(self.role_table.item(row, 2).text())
            src = self.role_table.item(row, 1).text()
            if self.owner_combo: self.owner_combo.setCurrentText(src)
        else:
            self.role_input.setText(self.role_table.item(row, 1).text())
