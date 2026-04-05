# widgets/crud_widget.py
# Extracted from main_window.py — CrudWidget (Users CRUD with Import from Excel)

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
    QGroupBox,
    QMessageBox,
    QAbstractItemView,
    QFrame,
)
from PyQt6.QtCore import Qt, pyqtSignal, QTimer, QSignalBlocker
from PyQt6.QtGui import QColor

import database_logic as db
import excel_logic as excel
from constants import DEBOUNCE_MS

logger = logging.getLogger(__name__)


def create_group_box(title: str, inner_layout) -> QGroupBox:
    box = QGroupBox(title)
    font = box.font()
    font.setBold(True)
    box.setFont(font)
    box.setLayout(inner_layout)
    return box


class CrudWidget(QWidget):
    # Signals for immediate UI sync
    import_done = pyqtSignal(str)  # emits 'source' when import finishes
    users_changed = pyqtSignal(str)  # emits 'source' when user list changes

    def __init__(self, source: str, excel_file: str, logged_username: str):
        super().__init__()
        self.source = source
        self.excel_file = excel_file
        self.logged_username = logged_username or "Unknown"
        self.current_user_id = None

        self._filter_state = {}
        self._debounce_timer = QTimer(self)
        self._debounce_timer.setSingleShot(True)
        self._debounce_timer.timeout.connect(self.load_users_table)

        layout = QHBoxLayout(self)

        # ---------------------------------------------------------
        # Left panel: User Form
        # ---------------------------------------------------------
        form_layout = QGridLayout()
        form_layout.setContentsMargins(8, 8, 8, 8)

        self.crud_name_input = QLineEdit()
        self.crud_role_input = QComboBox()
        self.crud_role_input.setEditable(False)
        self.crud_badge_input = QLineEdit()

        # --- NEW: Default Location Combos ---
        self.def_pickup_combo = QComboBox()
        self.def_dropoff_combo = QComboBox()

        # Buttons
        self.crud_save_button = QPushButton("💾 Save User")
        self.crud_save_button.clicked.connect(self.save_crud_user)
        self.crud_new_button = QPushButton("✨ New User")
        self.crud_new_button.clicked.connect(self.clear_crud_form)
        self.crud_delete_button = QPushButton("❌ Delete User")
        self.crud_delete_button.clicked.connect(self.delete_crud_user)

        self.import_button = QPushButton("📥 Import from Excel → DB (validated)")
        self.import_button.clicked.connect(self.import_users_from_excel)

        # Styling
        self.crud_save_button.setProperty("variant", "primary")
        self.crud_new_button.setProperty("variant", "secondary")
        self.crud_delete_button.setProperty("danger", True)
        self.import_button.setProperty("variant", "secondary")

        # Layout Setup
        row = 0
        form_layout.addWidget(QLabel("Full Name:"), row, 0)
        form_layout.addWidget(self.crud_name_input, row, 1)
        row += 1
        form_layout.addWidget(QLabel("Role/Department:"), row, 0)
        form_layout.addWidget(self.crud_role_input, row, 1) # Adding the combo
        row += 1
        form_layout.addWidget(QLabel("Badge (ID):"), row, 0)
        form_layout.addWidget(self.crud_badge_input, row, 1)
        row += 1

        # Separator for Logistics
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setFrameShadow(QFrame.Shadow.Sunken)
        form_layout.addWidget(sep, row, 0, 1, 2)
        row += 1

        lbl_logistics = QLabel("Default Logistics (Pre-defined):")
        font_l = lbl_logistics.font(); font_l.setBold(True)
        lbl_logistics.setFont(font_l)
        form_layout.addWidget(lbl_logistics, row, 0, 1, 2)
        row += 1

        form_layout.addWidget(QLabel("Default Pick Up:"), row, 0)
        form_layout.addWidget(self.def_pickup_combo, row, 1)
        row += 1
        form_layout.addWidget(QLabel("Default Drop Off:"), row, 0)
        form_layout.addWidget(self.def_dropoff_combo, row, 1)
        row += 1

        # Action Buttons
        button_layout = QHBoxLayout()
        #button_layout.addWidget(self.crud_new_button)
        button_layout.addWidget(self.crud_save_button)
        form_layout.addLayout(button_layout, row, 0, 1, 2)
        row += 1
        form_layout.addWidget(self.crud_delete_button, row, 0, 1, 2)
        row += 1
        #form_layout.addWidget(self.import_button, row, 0, 1, 2)

        form_group = create_group_box("Manage User", form_layout)
        form_group.setFixedWidth(400)

        # ---------------------------------------------------------
        # Right panel: Users Table
        # ---------------------------------------------------------
        table_panel = QWidget()
        table_layout = QVBoxLayout(table_panel)

        # Filter Panel
        filter_bar = QHBoxLayout()
        self.user_search_input = QLineEdit()
        self.user_search_input.setPlaceholderText("Search Name/Badge...")
        self.user_search_input.textChanged.connect(self._request_refresh)

        self.user_role_combo = QComboBox()
        self.user_role_combo.currentIndexChanged.connect(self._request_refresh)

        self.user_badge_prefix_input = QLineEdit()
        self.user_badge_prefix_input.setPlaceholderText("Badge prefix...")
        self.user_badge_prefix_input.textChanged.connect(self._request_refresh)

        user_reset_btn = QPushButton("Reset")
        user_reset_btn.clicked.connect(self.reset_filters)

        filter_bar.addWidget(self.user_search_input, 2)
        filter_bar.addWidget(QLabel("Role:"))
        filter_bar.addWidget(self.user_role_combo, 1)
        filter_bar.addWidget(QLabel("Badge Prefix:"))
        filter_bar.addWidget(self.user_badge_prefix_input, 1)
        filter_bar.addWidget(user_reset_btn)
        table_layout.addLayout(filter_bar)

        self.users_table = QTableWidget()
        self.users_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.users_table.setAlternatingRowColors(True)
        self.users_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.users_table.itemClicked.connect(self.load_user_to_crud_form)
        table_layout.addWidget(self.users_table)

        table_group = create_group_box("Registered Users List", table_layout)
        layout.addWidget(form_group)
        layout.addWidget(table_group)

        # Initial Load
        self.refresh_ui_data()
    def populate_role_combo(self):
        """Refreshes the Role dropdown from the DB Master List."""
        current_text = self.crud_role_input.currentText()
        self.crud_role_input.blockSignals(True)
        self.crud_role_input.clear()

        # Add a default blank
        self.crud_role_input.addItem("", None)

        # Fetch from roles table
        roles = db.get_roles(self.source)
        print(f"DEBUG: Cargando roles para {self.source}. Encontrados: {len(roles)}")
        for r in roles:
            self.crud_role_input.addItem(r["name"], r["id"])

        # Try to restore selection if text matches
        idx = self.crud_role_input.findText(current_text)
        if idx >= 0:
            self.crud_role_input.setCurrentIndex(idx)

        self.crud_role_input.blockSignals(False)

    def _request_refresh(self):
        self._debounce_timer.start(DEBOUNCE_MS)

    def _populate_role_filter(self):
        # Populate Role Filter
        self.user_role_combo.blockSignals(True)
        current_role = self.user_role_combo.currentText()
        self.user_role_combo.clear()
        self.user_role_combo.addItem("All Roles")
        all_users = db.get_all_users(self.source)
        roles = sorted(list(set(u["role"] for u in all_users if u.get("role"))))
        self.user_role_combo.addItems(roles)
        idx = self.user_role_combo.findText(current_role)
        if idx > 0:
            self.user_role_combo.setCurrentIndex(idx)
        self.user_role_combo.blockSignals(False)

        # Populate Location Combos (Default Pick/Drop)
        self._populate_location_combos(self.def_pickup_combo)
        self._populate_location_combos(self.def_dropoff_combo)

    def _populate_location_combos(self, combo: QComboBox):
        """Helper to fill location combos keeping current selection if possible."""
        current_data = combo.currentData()
        combo.blockSignals(True)
        combo.clear()
        combo.addItem("— None —", None)
        locations = db.get_locations(self.source)
        for loc in locations:
            combo.addItem(loc["pickup_location"], loc["pickup_location"])

        if current_data:
            idx = combo.findData(current_data)
            if idx >= 0:
                combo.setCurrentIndex(idx)
        combo.blockSignals(False)

    def reset_filters(self):
        with QSignalBlocker(self.user_search_input), QSignalBlocker(
            self.user_role_combo
        ), QSignalBlocker(self.user_badge_prefix_input):
            self.user_search_input.clear()
            self.user_role_combo.setCurrentIndex(0)
            self.user_badge_prefix_input.clear()
        self.load_users_table()

    def load_users_table(self):
        self._filter_state["text"] = self.user_search_input.text()
        self._filter_state["role"] = (
            self.user_role_combo.currentText()
            if self.user_role_combo.currentIndex() > 0
            else None
        )
        self._filter_state["badge_prefix"] = self.user_badge_prefix_input.text()

        # --- MODIFIED: Use the JOINED query to get defaults ---
        # Note: Filtering by defaults isn't requested, but we need to display them.
        # If filtering is active, we might need a more complex query or filter in Python.
        # Since get_users_filtered doesn't support defaults, we use get_users_with_defaults
        # and filter in Python for simplicity (assuming < 2000 users).
        # For production with 50k+ users, we'd add defaults to get_users_filtered SQL.

        all_users_with_defaults = db.get_users_with_defaults(self.source)

        # Apply filters in Python
        filtered_users = []
        txt = (self._filter_state["text"] or "").lower()
        role_filter = self._filter_state["role"]
        pfx = (self._filter_state["badge_prefix"] or "").lower()

        for u in all_users_with_defaults:
            if txt and (txt not in u["name"].lower() and txt not in u["badge"].lower()):
                continue
            if role_filter and u["role"] != role_filter:
                continue
            if pfx and not u["badge"].lower().startswith(pfx):
                continue
            filtered_users.append(u)

        # Columns: ID, Name, Role, Badge, Default Pick, Default Drop
        headers = ["ID", "Name", "Role", "Badge", "Def. Pick Up", "Def. Drop Off"]
        self.users_table.setRowCount(len(filtered_users))
        self.users_table.setColumnCount(len(headers))
        self.users_table.setHorizontalHeaderLabels(headers)

        for row, user in enumerate(filtered_users):
            self.users_table.setItem(row, 0, QTableWidgetItem(str(user["id"])))
            self.users_table.setItem(row, 1, QTableWidgetItem(user["name"]))
            self.users_table.setItem(row, 2, QTableWidgetItem(user["role"]))
            self.users_table.setItem(row, 3, QTableWidgetItem(user["badge"]))

            # New Columns
            pu = user.get("pickup_location") or ""
            do = user.get("dropoff_location") or ""

            item_pu = QTableWidgetItem(pu)
            item_pu.setForeground(QColor("#1565C0") if pu else QColor("#9E9E9E")) # Blue if set
            self.users_table.setItem(row, 4, item_pu)

            item_do = QTableWidgetItem(do)
            item_do.setForeground(QColor("#1565C0") if do else QColor("#9E9E9E"))
            self.users_table.setItem(row, 5, item_do)

        self.users_table.setColumnHidden(0, True)  # hide ID column
        self.users_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        # Give more space to Logistics columns if needed
        self.users_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        self.users_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)

    def load_user_to_crud_form(self, item):
        row = item.row()
        self.current_user_id = int(self.users_table.item(row, 0).text())
        self.crud_name_input.setText(self.users_table.item(row, 1).text())

        # Load Role into Combo
        role_text = self.users_table.item(row, 2).text()
        idx = self.crud_role_input.findText(role_text)
        if idx >= 0:
            self.crud_role_input.setCurrentIndex(idx)
        else:
            # If the role isn't in the list (legacy data), we can't select it easily
            # if editable is False.
            # Ideally, migration handled this. Defaults to blank if not found.
            self.crud_role_input.setCurrentIndex(0)

        self.crud_badge_input.setText(self.users_table.item(row, 3).text())

        # Load Defaults from table (columns 4 and 5)
        # Note: We rely on the text in the table matching the data in the combo
        def_pickup = self.users_table.item(row, 4).text()
        def_dropoff = self.users_table.item(row, 5).text()

        idx_pu = self.def_pickup_combo.findData(def_pickup if def_pickup else None)
        self.def_pickup_combo.setCurrentIndex(idx_pu if idx_pu >= 0 else 0)

        idx_do = self.def_dropoff_combo.findData(def_dropoff if def_dropoff else None)
        self.def_dropoff_combo.setCurrentIndex(idx_do if idx_do >= 0 else 0)

    def clear_crud_form(self):
        self.current_user_id = None
        self.crud_name_input.clear()
        self.crud_role_input.setCurrentIndex(0)
        self.crud_badge_input.clear()
        self.def_pickup_combo.setCurrentIndex(0)
        self.def_dropoff_combo.setCurrentIndex(0)
        self.users_table.clearSelection()

    def save_crud_user(self):
        name = self.crud_name_input.text().strip()
        # Get Role from Combo Text (NOT ID)
        # This ensures we save "Driver" to the user table, keeping PlanStaffWidget compatible.
        role = self.crud_role_input.currentText().strip()

        badge = self.crud_badge_input.text().strip()

        # Get defaults
        def_pickup = self.def_pickup_combo.currentData()
        def_dropoff = self.def_dropoff_combo.currentData()

        if not name or not role or not badge:
            QMessageBox.warning(self, "Incomplete Data", "Name, Role, and Badge are required.")
            return

        old_badge = None

        # 1. Save User Core Data
        if self.current_user_id:
            # Use cascade update: handles badge rename across all dependent tables
            success, message, old_badge = db.update_user_with_cascade(
                self.current_user_id, name, role, badge, self.source
            )
        else:
            success, message = db.add_user(name, role, badge, self.source)

        if success:
            # 2. Save Default Logistics (SSoT: user_locations where is_default=1)
            try:
                db.set_user_default_locations(badge, def_pickup, def_dropoff)
                message += "\nDefault locations updated."
            except Exception as e:
                message += f"\nWarning: Could not save locations ({e})"

            # 3. Sync Excel — targeted update instead of full refresh
            try:
                if old_badge and old_badge != badge:
                    # Badge changed: rename in Excel (handles duplicates/merge)
                    excel.rename_user_badge_in_excel(
                        self.excel_file, self.source, old_badge, badge
                    )
                    # Also update name/role in the renamed row
                    excel.update_user_info_in_excel(
                        self.excel_file, self.source, badge, name, role
                    )
                elif self.current_user_id:
                    # Editing existing user (name/role changed, badge same)
                    excel.update_user_info_in_excel(
                        self.excel_file, self.source, badge, name, role
                    )
                else:
                    # New user: full refresh adds the missing row
                    excel.refresh_excel_from_db(self.excel_file, self.source)
            except Exception as e:
                print(f"Excel sync failed: {e}")

            # 4. Refresh UI
            self._populate_role_filter() # Refresh roles in case user added a new one
            self.refresh_ui_data()       # Refresh table to show new columns
            self.populate_role_combo()
            self.users_changed.emit(self.source) # Notify Plan Staff tab

        box = QMessageBox(self)
        box.setIcon(QMessageBox.Icon.Information if success else QMessageBox.Icon.Warning)
        box.setWindowTitle("Success" if success else "Error")
        box.setText(message)
        box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
        box.exec()

    def delete_crud_user(self):
        if not self.current_user_id:
            QMessageBox.warning(self, "No Selection", "Please select a user to delete.")
            return

        badge_to_remove = self.crud_badge_input.text().strip()

        confirm = QMessageBox(self)
        confirm.setIcon(QMessageBox.Icon.Question)
        confirm.setWindowTitle("Confirm User Deletion")
        confirm.setText(f"Are you sure you want to delete user '{self.crud_name_input.text()}'?")
        yes_btn = confirm.addButton("Yes", QMessageBox.ButtonRole.YesRole)
        confirm.addButton("No", QMessageBox.ButtonRole.NoRole)
        confirm.exec()

        if confirm.clickedButton() == yes_btn:
            success_db, message_db = db.delete_user(self.current_user_id)
            final_msg = message_db

            if success_db:
                # Also clean defaults from DB (Optional but clean)
                # Note: db.delete_user only deletes from 'users'.
                # Ideally, foreign keys would handle this, or we run a manual delete on user_locations.
                # For now, leaving orphan defaults is harmless, but we could add:
                # db.set_user_default_locations(badge_to_remove, None, None)

                success_excel, msg_excel = excel.remove_user_from_excel(self.excel_file, badge_to_remove)
                if success_excel:
                    final_msg += f"\n\nRemoved from Excel: {badge_to_remove}"
                else:
                    final_msg += f"\n\nWarning: Excel sync failed ({msg_excel})"

                db.log_event(self.logged_username, self.source, "USER_DELETE", f"Deleted {badge_to_remove}")

                self.refresh_ui_data()
                self.users_changed.emit(self.source)

            QMessageBox.information(self, "Deletion Result", final_msg)

    def import_users_from_excel(self):
        # (Sin cambios en esta función)
        try:
            inserted, skipped, upserts = excel.import_excel_to_db(self.excel_file, self.source)
            db.log_event(self.logged_username, self.source, "DATA_IMPORT", f"users_inserted={inserted}; skipped={skipped}")
            QMessageBox.information(self, "Import Complete", f"Imported {inserted} new users.\nSkipped {skipped} existing.\nUpserted {upserts} schedules.")
            self.refresh_ui_data()
            self.users_changed.emit(self.source)
            self.import_done.emit(self.source)
        except ValueError as ve:
            QMessageBox.critical(self, "Invalid Excel", str(ve))

    def refresh_ui_data(self):
        self.populate_role_combo()
        self._populate_role_filter()
        self.load_users_table()
        self.clear_crud_form()
