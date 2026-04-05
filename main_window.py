# main_window.py
# Main application windows: MainWindow (RGM/Newmont profile) and AdminMainWindow.
# All widget classes have been extracted to the widgets/ package.

import logging
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QTabWidget, QApplication, QAbstractItemView,
)
from PyQt6.QtCore import Qt, pyqtSignal, QTimer
from PyQt6.QtGui import QKeySequence, QAction

import database_logic as db
from constants import Source

# Import all widgets from the widgets package
from widgets.session_widgets import SessionBarWidget
from widgets.plan_staff_widget import PlanStaffWidget
from widgets.rotation_widget import RotationHistoryWidget
from widgets.crud_widget import CrudWidget
from widgets.shift_type_widget import ShiftTypeAdminWidget
from widgets.location_widget import LocationAdminWidget
from widgets.role_widget import RoleAdminWidget
from widgets.report_settings_widget import ReportSettingsWidget
from widgets.audit_widget import AuditLogWidget

logger = logging.getLogger(__name__)


# -------------------------------------------------------------
# Main window (normal profile: RGM or Newmont)
# -------------------------------------------------------------
class MainWindow(QMainWindow):
    logout_signal = pyqtSignal()

    def __init__(
        self,
        user_role,
        excel_file,
        logged_username=None,
        can_manage_shift_types: bool = False,
        preloaded_data=None,
        is_editor: bool = True,
    ):
        super().__init__()
        self.user_role = user_role  # RGM or Newmont. Used as 'source'
        self.excel_file = excel_file
        self.logged_username = logged_username or "Unknown"
        self.can_manage_shift_types = bool(can_manage_shift_types)
        self._is_editor = is_editor

        db.setup_database()
        db.log_event(
            self.logged_username,
            self.user_role,
            "USER_LOGIN",
            f"Excel={self.excel_file}",
        )

        _mode_tag = "Editing" if self._is_editor else "View Only"
        self.setWindowTitle(
            f"Operations Manager - {self.user_role} | {self.logged_username} [{_mode_tag}]"
        )
        self.setGeometry(100, 100, 1200, 800)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Top bar
        top_layout = QHBoxLayout()
        title_label = QLabel(f"Transport & Operations Manager ({self.user_role})")
        font = title_label.font()
        font.setPointSize(20)
        font.setBold(True)
        title_label.setFont(font)

        # Session avatars bar (co-authoring indicators)
        self._session_bar = SessionBarWidget(self.logged_username)
        self._session_bar.refresh()

        logout_button = QPushButton("Log Out")
        logout_button.setFixedWidth(120)
        logout_button.setProperty("variant", "text")
        logout_button.clicked.connect(self.handle_logout)

        top_layout.addWidget(title_label)
        top_layout.addStretch()
        top_layout.addWidget(self._session_bar)
        top_layout.addWidget(logout_button)
        main_layout.addLayout(top_layout)

        # Tabs
        tabs = QTabWidget()
        main_layout.addWidget(tabs)

        # 1) Plan Staff (preview/register/reports)
        self.plan_widget = PlanStaffWidget(
            self.user_role,
            self.excel_file,
            self.logged_username,
            preloaded_data=preloaded_data
        )
        self.plan_widget.excel_path_changed.connect(self._on_excel_path_changed)
        tabs.addTab(self.plan_widget, "Plan Staff & Reports")

        # 2) Rotation History
        self.rotation_widget = RotationHistoryWidget(
            created_by=self.logged_username
        )
        tabs.addTab(self.rotation_widget, "Rotation History")
        self.plan_widget.rotation_changed.connect(self.rotation_widget.refresh_data)

        # 3) Users CRUD
        self.crud_widget = CrudWidget(
            self.user_role, self.excel_file, self.logged_username
        )
        tabs.addTab(self.crud_widget, "Users (CRUD)")

        # 4) Shift Types (only if user can manage them)
        if self.can_manage_shift_types:
            self.shift_types_widget = ShiftTypeAdminWidget(
                self.user_role, self.excel_file, self.logged_username
            )
            tabs.addTab(self.shift_types_widget, f"{self.user_role} Shift Types")
            self.shift_types_widget.types_changed.connect(
                lambda src: self.plan_widget.refresh_ui_data()
            )

        # 5) Locations admin
        self.location_widget = LocationAdminWidget(scope_source=self.user_role)
        tabs.addTab(self.location_widget, "Location")
        self.location_widget.locations_changed.connect(
            self.plan_widget.load_location_options
        )
        self.location_widget.locations_changed.connect(
            lambda: self.crud_widget._populate_location_combos(self.crud_widget.def_pickup_combo)
        )
        self.location_widget.locations_changed.connect(
            lambda: self.crud_widget._populate_location_combos(self.crud_widget.def_dropoff_combo)
        )

        # 6) Roles
        self.role_widget = RoleAdminWidget(scope_source=self.user_role)
        tabs.addTab(self.role_widget, "Roles / Dept")
        self.role_widget.roles_changed.connect(
            self.crud_widget.populate_role_combo
        )
        self.role_widget.roles_changed.connect(
            self.crud_widget._populate_role_filter
        )

        # 7) Settings Tab
        self.settings_widget = ReportSettingsWidget(
            self.logged_username, self.user_role
        )
        tabs.addTab(self.settings_widget, "Settings")

        # Hot sync
        self.crud_widget.users_changed.connect(
            lambda src: (
                self.plan_widget.refresh_ui_data() if src == self.user_role else None
            )
        )
        self.crud_widget.import_done.connect(
            lambda src: (
                self.plan_widget.refresh_ui_data() if src == self.user_role else None
            )
        )

        self._setup_undo_menu()

        # Apply read-only mode after all widgets are created
        if not self._is_editor:
            QTimer.singleShot(200, self._apply_readonly_mode)

    def _setup_undo_menu(self):
        edit_menu = self.menuBar().addMenu("Edit")

        self._act_undo = QAction("Undo", self)
        self._act_undo.setShortcut(QKeySequence.StandardKey.Undo)
        self._act_undo.setShortcutContext(Qt.ShortcutContext.WindowShortcut)
        self._act_undo.setEnabled(False)
        self._act_undo.triggered.connect(self._smart_undo)
        edit_menu.addAction(self._act_undo)

        self._act_redo = QAction("Redo", self)
        self._act_redo.setShortcut(QKeySequence.StandardKey.Redo)
        self._act_redo.setShortcutContext(Qt.ShortcutContext.WindowShortcut)
        self._act_redo.setEnabled(False)
        self._act_redo.triggered.connect(self._smart_redo)
        edit_menu.addAction(self._act_redo)

        stack = self.plan_widget._undo_stack
        stack.canUndoChanged.connect(self._act_undo.setEnabled)
        stack.canRedoChanged.connect(self._act_redo.setEnabled)
        stack.indexChanged.connect(self._refresh_undo_tooltips)

    def _refresh_undo_tooltips(self):
        stack = self.plan_widget._undo_stack
        self._act_undo.setText(
            f"Undo: {stack.undoText()}" if stack.canUndo() else "Undo"
        )
        self._act_redo.setText(
            f"Redo: {stack.redoText()}" if stack.canRedo() else "Redo"
        )

    def _smart_undo(self):
        from PyQt6.QtWidgets import QLineEdit, QTextEdit, QPlainTextEdit
        w = QApplication.focusWidget()
        if isinstance(w, (QLineEdit, QTextEdit, QPlainTextEdit)):
            w.undo()
            return
        stack = self.plan_widget._undo_stack
        if stack.canUndo():
            self.plan_widget._undo_in_progress = True
            try:
                stack.undo()
            finally:
                self.plan_widget._undo_in_progress = False

    def _smart_redo(self):
        from PyQt6.QtWidgets import QLineEdit, QTextEdit, QPlainTextEdit
        w = QApplication.focusWidget()
        if isinstance(w, (QLineEdit, QTextEdit, QPlainTextEdit)):
            w.redo()
            return
        stack = self.plan_widget._undo_stack
        if stack.canRedo():
            self.plan_widget._undo_in_progress = True
            try:
                stack.redo()
            finally:
                self.plan_widget._undo_in_progress = False

    def _sync_after_users_changed(self, src: str):
        if src == self.user_role:
            self.plan_widget.refresh_users_only()
            QApplication.processEvents()

    def handle_logout(self):
        if hasattr(self, "_session_bar"):
            self._session_bar.stop()
        self.logout_signal.emit()
        self.close()

    def _apply_readonly_mode(self):
        if self._is_editor:
            return
        from PyQt6.QtWidgets import (
            QPushButton, QLineEdit, QComboBox, QDateEdit,
            QTimeEdit, QCheckBox, QToolButton, QTableWidget,
        )
        for widget in self.findChildren(
            (QPushButton, QLineEdit, QComboBox, QDateEdit,
             QTimeEdit, QCheckBox, QToolButton)
        ):
            if widget.property("variant") == "text" and isinstance(widget, QPushButton):
                text = widget.text()
                if "Log Out" in text:
                    continue
            widget.setEnabled(False)
        for table in self.findChildren(QTableWidget):
            table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)

    def _on_excel_path_changed(self, source: str, new_path: str) -> None:
        if source != self.user_role:
            return
        self.excel_file = new_path
        if hasattr(self, "crud_widget"):
            self.crud_widget.excel_file = new_path
        if getattr(self, "can_manage_shift_types", False) and hasattr(self, "shift_types_widget"):
            self.shift_types_widget.excel_file = new_path


# -------------------------------------------------------------
# Administrator window (unified access)
# -------------------------------------------------------------
class AdminMainWindow(QMainWindow):
    logout_signal = pyqtSignal()

    def __init__(self, logged_username: str, rgm_excel: str, newmont_excel: str, is_editor: bool = True):
        super().__init__()
        self.logged_username = logged_username or "admin"
        self._is_editor = is_editor

        db.setup_database()
        db.log_event(
            self.logged_username,
            Source.ADMINISTRATOR.value,
            "USER_LOGIN",
            f"Access to admin console | RGM={rgm_excel} | Newmont={newmont_excel}",
        )

        _mode_tag = "Editing" if self._is_editor else "View Only"
        self.setWindowTitle(f"Administrator Console | {self.logged_username} [{_mode_tag}]")
        self.setGeometry(100, 100, 1400, 900)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Top bar
        top_layout = QHBoxLayout()
        title_label = QLabel("Unified Access — RGM & Newmont")
        font = title_label.font()
        font.setPointSize(20)
        font.setBold(True)
        title_label.setFont(font)

        self._session_bar = SessionBarWidget(self.logged_username)
        self._session_bar.refresh()

        logout_button = QPushButton("Log Out")
        logout_button.setFixedWidth(120)
        logout_button.setProperty("variant", "text")
        logout_button.clicked.connect(self.handle_logout)

        top_layout.addWidget(title_label)
        top_layout.addStretch()
        top_layout.addWidget(self._session_bar)
        top_layout.addWidget(logout_button)
        main_layout.addLayout(top_layout)

        # --- Dynamic tab generation from operational sources ---
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        sources_config = {
            Source.RGM.value: rgm_excel,
            Source.NEWMONT.value: newmont_excel,
        }

        self._plan_widgets = {}
        self._crud_widgets = {}
        self._type_widgets = {}

        for src_name, src_excel in sources_config.items():
            # CRUD
            crud = CrudWidget(src_name, src_excel, self.logged_username)
            self._crud_widgets[src_name] = crud
            self.tabs.addTab(crud, f"{src_name} CRUD")

            # Plan Staff
            plan = PlanStaffWidget(src_name, src_excel, self.logged_username)
            self._plan_widgets[src_name] = plan
            self.tabs.addTab(plan, f"{src_name} Plan Staff")
            plan.excel_path_changed.connect(self._on_excel_path_changed)

            # Shift Types
            types_w = ShiftTypeAdminWidget(src_name, src_excel, self.logged_username)
            self._type_widgets[src_name] = types_w

        # Backward compatibility aliases
        self.rgm_crud = self._crud_widgets[Source.RGM.value]
        self.nm_crud = self._crud_widgets[Source.NEWMONT.value]
        self.rgm_plan = self._plan_widgets[Source.RGM.value]
        self.nm_plan = self._plan_widgets[Source.NEWMONT.value]
        self.rgm_types = self._type_widgets[Source.RGM.value]
        self.nm_types = self._type_widgets[Source.NEWMONT.value]

        # 5) Rotation History (global)
        self.rotation_history = RotationHistoryWidget(created_by=None)
        self.tabs.addTab(self.rotation_history, "Rotation History")
        for plan in self._plan_widgets.values():
            plan.rotation_changed.connect(self.rotation_history.refresh_data)

        # 6) Audit Log (global)
        audit_all = AuditLogWidget(source=None)
        self.tabs.addTab(audit_all, "Audit Log")

        # 7) Shift Types tabs
        for src_name, types_w in self._type_widgets.items():
            self.tabs.addTab(types_w, f"{src_name} Shift Types")

        # 8) Locations (global admin)
        self.location_admin = LocationAdminWidget(scope_source=None)
        self.tabs.addTab(self.location_admin, "Locations")
        for plan in self._plan_widgets.values():
            self.location_admin.locations_changed.connect(
                plan.load_location_options
            )

        # 9) Settings
        settings_container = QWidget()
        settings_layout = QVBoxLayout(settings_container)
        settings_tabs = QTabWidget()
        settings_layout.addWidget(settings_tabs)
        for src_name in sources_config:
            s_widget = ReportSettingsWidget(self.logged_username, src_name)
            settings_tabs.addTab(s_widget, f"{src_name} Report Settings")
        self.tabs.addTab(settings_container, "Settings")

        # Hot sync: CRUD changes refresh plan
        for src_name in sources_config:
            crud = self._crud_widgets[src_name]
            plan = self._plan_widgets[src_name]
            crud.users_changed.connect(lambda s, p=plan: p.refresh_ui_data())
            crud.import_done.connect(lambda s, p=plan: p.refresh_ui_data())

        # Shift type changes refresh plan
        for src_name in sources_config:
            types_w = self._type_widgets[src_name]
            plan = self._plan_widgets[src_name]
            types_w.types_changed.connect(lambda s, p=plan: p.refresh_ui_data())

        self._setup_admin_undo_menu()

        if not self._is_editor:
            QTimer.singleShot(200, self._apply_readonly_mode)

    def _apply_readonly_mode(self):
        if self._is_editor:
            return
        from PyQt6.QtWidgets import (
            QPushButton, QLineEdit, QComboBox, QDateEdit,
            QTimeEdit, QCheckBox, QToolButton, QTableWidget,
        )
        for widget in self.findChildren(
            (QPushButton, QLineEdit, QComboBox, QDateEdit,
             QTimeEdit, QCheckBox, QToolButton)
        ):
            if widget.property("variant") == "text" and isinstance(widget, QPushButton):
                text = widget.text()
                if "Log Out" in text:
                    continue
            widget.setEnabled(False)
        for table in self.findChildren(QTableWidget):
            table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)

    def _setup_admin_undo_menu(self):
        edit_menu = self.menuBar().addMenu("Edit")

        self._act_undo_admin = QAction("Undo", self)
        self._act_undo_admin.setShortcut(QKeySequence.StandardKey.Undo)
        self._act_undo_admin.setShortcutContext(Qt.ShortcutContext.WindowShortcut)
        self._act_undo_admin.setEnabled(False)
        self._act_undo_admin.triggered.connect(self._admin_smart_undo)
        edit_menu.addAction(self._act_undo_admin)

        self._act_redo_admin = QAction("Redo", self)
        self._act_redo_admin.setShortcut(QKeySequence.StandardKey.Redo)
        self._act_redo_admin.setShortcutContext(Qt.ShortcutContext.WindowShortcut)
        self._act_redo_admin.setEnabled(False)
        self._act_redo_admin.triggered.connect(self._admin_smart_redo)
        edit_menu.addAction(self._act_redo_admin)

        self.tabs.currentChanged.connect(self._admin_refresh_undo_state)
        for plan in self._plan_widgets.values():
            plan._undo_stack.indexChanged.connect(self._admin_refresh_undo_state)

    def _active_plan_widget(self):
        current = self.tabs.currentWidget()
        for plan in self._plan_widgets.values():
            if current is plan:
                return plan
        return None

    def _admin_refresh_undo_state(self):
        plan = self._active_plan_widget()
        if plan:
            self._act_undo_admin.setEnabled(plan._undo_stack.canUndo())
            self._act_redo_admin.setEnabled(plan._undo_stack.canRedo())
        else:
            self._act_undo_admin.setEnabled(False)
            self._act_redo_admin.setEnabled(False)

    def _admin_smart_undo(self):
        from PyQt6.QtWidgets import QLineEdit, QTextEdit, QPlainTextEdit
        w = QApplication.focusWidget()
        if isinstance(w, (QLineEdit, QTextEdit, QPlainTextEdit)):
            w.undo(); return
        plan = self._active_plan_widget()
        if plan and plan._undo_stack.canUndo():
            plan._undo_in_progress = True
            try:
                plan._undo_stack.undo()
            finally:
                plan._undo_in_progress = False

    def _admin_smart_redo(self):
        from PyQt6.QtWidgets import QLineEdit, QTextEdit, QPlainTextEdit
        w = QApplication.focusWidget()
        if isinstance(w, (QLineEdit, QTextEdit, QPlainTextEdit)):
            w.redo(); return
        plan = self._active_plan_widget()
        if plan and plan._undo_stack.canRedo():
            plan._undo_in_progress = True
            try:
                plan._undo_stack.redo()
            finally:
                plan._undo_in_progress = False

    def handle_logout(self):
        if hasattr(self, "_session_bar"):
            self._session_bar.stop()
        self.logout_signal.emit()
        self.close()

    def _on_excel_path_changed(self, source: str, new_path: str) -> None:
        if source in self._crud_widgets:
            self._crud_widgets[source].excel_file = new_path
        if source in self._type_widgets:
            self._type_widgets[source].excel_file = new_path
  