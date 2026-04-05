# services/schedule_service.py
# Schedule business logic extracted from PlanStaffWidget.
# Handles schedule validation, conflict detection, and operations management.

import logging
from datetime import date, timedelta
from typing import Optional, Dict, List, Tuple

import database_logic as db
from constants import ShiftStatus

logger = logging.getLogger(__name__)


class ScheduleService:
    """Business logic for schedule management."""

    @staticmethod
    def is_off_to_on_transition(old_status: Optional[str], new_status: str) -> bool:
        """Check if a status change is an OFF -> ON transition (requires confirmation)."""
        if not old_status:
            return False
        old = old_status.strip().upper()
        new = new_status.strip().upper()
        off_codes = ShiftStatus.off_aliases()
        on_codes = {ShiftStatus.ON.value, ShiftStatus.ON_NS.value}
        return old in off_codes and new in on_codes

    @staticmethod
    def get_shift_color_map(source: str) -> Dict[str, Dict]:
        """Get the shift type map for a source (code -> info dict)."""
        return db.get_shift_type_map(source)

    @staticmethod
    def save_schedule_range(
        badge: str, start_d: date, end_d: date,
        status: str, shift_type: Optional[str], source: str,
        in_time: Optional[str] = None, out_time: Optional[str] = None,
        remark: Optional[str] = None,
        force_new_entry_start: Optional[int] = None,
    ) -> int:
        """Save a schedule range and return the number of days written."""
        return db.upsert_schedule_range(
            badge, start_d, end_d, status, shift_type, source,
            in_time, out_time, remark, force_new_entry_start,
        )

    @staticmethod
    def clear_schedule_range(badge: str, start_d: date, end_d: date, source: str) -> int:
        """Clear schedule data for a date range."""
        return db.clear_schedule_range(badge, start_d, end_d, source)

    @staticmethod
    def get_schedule_map(badge: str, start_d: date, end_d: date, source: str) -> Dict:
        """Get schedule data for a badge in a date range."""
        return db.get_schedule_map_for_range(badge, start_d, end_d, source)

    @staticmethod
    def detect_operation_conflict(badge: str, check_date: date) -> Optional[Dict]:
        """Check if an operation already exists covering a date."""
        return db.get_operation_overlapping(badge, check_date)

    @staticmethod
    def register_operation(
        username: str, role: str, badge: str,
        start_date: date, end_date: date, created_by: str,
        entry_date=None, exit_date=None,
    ) -> None:
        """Register a rotation/operation period."""
        db.add_operation(
            username, role, badge, start_date, end_date,
            created_by, entry_date, exit_date,
        )

    @staticmethod
    def log_audit(username: str, source: str, action: str, detail: str = "") -> None:
        """Log an audit event."""
        db.log_event(username, source, action, detail)
