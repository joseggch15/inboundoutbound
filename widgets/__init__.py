# widgets/__init__.py
# Widget modules extracted from main_window.py for better organization.
# Each module contains one or more related PyQt6 widget classes.

from widgets.session_widgets import UserAvatarWidget, SessionBarWidget
from widgets.common_widgets import (
    ShiftCellDelegate, ShiftInfoCard, CollapsibleGroupBox,
    DayScheduleEditor, WeekendHeader,
)
from widgets.undo_commands import UndoCellChangeCommand, UndoScheduleSnapshotCommand
from widgets.role_widget import RoleAdminWidget
from widgets.report_settings_widget import ReportSettingsWidget
from widgets.plan_staff_widget import PlanStaffWidget
from widgets.rotation_widget import RotationHistoryWidget
from widgets.crud_widget import CrudWidget
from widgets.shift_type_widget import ShiftTypeAdminWidget
from widgets.location_widget import LocationAdminWidget
from widgets.audit_widget import AuditLogWidget
