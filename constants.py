# constants.py
# Centralized constants and enumerations for the application.
# Replaces magic strings scattered across the codebase.

from enum import Enum


# ─── Data Sources (Companies) ────────────────────────────────────────
class Source(str, Enum):
    """Company/data source identifiers."""
    RGM = "RGM"
    NEWMONT = "Newmont"
    ADMINISTRATOR = "Administrator"

    @classmethod
    def operational_sources(cls):
        """Sources that have their own schedule data (excludes Administrator)."""
        return [cls.RGM, cls.NEWMONT]


# ─── Shift Status Codes ─────────────────────────────────────────────
class ShiftStatus(str, Enum):
    """Standard shift status codes used in schedules."""
    ON = "ON"
    ON_NS = "ON NS"
    OFF = "OFF"

    @classmethod
    def off_aliases(cls):
        """All string values that map to OFF."""
        return {"OFF", "BREAK", "KO", "LEAVE"}

    @classmethod
    def on_aliases(cls):
        """All string values that map to ON (day)."""
        return {"ON", "OK"}


class ShiftLabel(str, Enum):
    """Human-readable shift type labels."""
    DAY_SHIFT = "Day Shift"
    NIGHT_SHIFT = "Night Shift"


# ─── Shift Colors (hex) ─────────────────────────────────────────────
class ShiftColor(str, Enum):
    """Default hex colors for shift statuses."""
    GREEN_ON = "#C6EFCE"
    RED_OFF = "#FFC7CE"
    YELLOW_NS = "#FFFF99"


# ─── Excel Badge Prefixes ───────────────────────────────────────────
class BadgePrefix(str, Enum):
    """Badge prefix based on source file."""
    RGM = "ID"
    NEWMONT = "NM"


# ─── UI Constants ────────────────────────────────────────────────────
WARN_BG_HEX = "#FFFBEA"
FROZEN_COLUMN_COUNT = 3  # ROLE, NAME, BADGE
DEBOUNCE_MS = 200
WEEKEND_HEADER_YELLOW = "#FFEB3B"

AVATAR_COLORS = ["#0288D1", "#7B1FA2", "#388E3C", "#F57C00", "#D32F2F", "#00796B"]


# ─── Report Headers ─────────────────────────────────────────────────
NEWMONT_REPORT_HEADERS = [
    "#",
    "NAME",
    "FIRST NAME",
    "GID",
    "COMPANY",
    "DEPT",
    "FROM",
    "TO",
    "DATE",
    "TIME",
]

RGM_REPORT_HEADERS = [
    "NR",
    "NAME (Last, First Name)",
    "DEPARTMENT",
    "BADGE #",
    "POSITION / TITLE",
    "CREW A/B/C",
    "PICK UP LOCATION",
    "IN BOUND DATE",
    "Method Of Transport",
    "Location",
    "DEPT TIME",
    "ROSEBEL SITE OUT BOUND DATE",
]


# ─── System Shift Type Defaults ─────────────────────────────────────
SYSTEM_SHIFT_DEFAULTS = {
    Source.RGM.value: [
        {"code": "OFF",   "name": "OFF",                "color_hex": "#FFC7CE", "in_time": "00:00", "out_time": "00:00", "is_off": 1, "no_transport": 0, "apply_1d": 0},
        {"code": "ON",    "name": "ON (Day Shift)",     "color_hex": "#C6EFCE", "in_time": "07:00", "out_time": "07:00", "is_off": 0, "no_transport": 0, "apply_1d": 1},
        {"code": "ON NS", "name": "ON NS (Night Shift)","color_hex": "#FFFF99", "in_time": "07:00", "out_time": "07:00", "is_off": 0, "no_transport": 0, "apply_1d": 1},
    ],
    Source.NEWMONT.value: [
        {"code": "OFF",   "name": "OFF",                "color_hex": "#FFC7CE", "in_time": "00:00", "out_time": "00:00", "is_off": 1, "no_transport": 0, "apply_1d": 0},
        {"code": "ON",    "name": "ON (Day Shift)",     "color_hex": "#C6EFCE", "in_time": "06:00", "out_time": "12:00", "is_off": 0, "no_transport": 0, "apply_1d": 0},
        {"code": "ON NS", "name": "ON NS (Night Shift)","color_hex": "#FFFF99", "in_time": "12:00", "out_time": "06:00", "is_off": 0, "no_transport": 0, "apply_1d": 0},
    ],
}


# ─── Report Defaults by Source ───────────────────────────────────────
REPORT_DEFAULTS = {
    Source.NEWMONT.value: {
        "font_name": "Calibri",
        "font_color": "#000000",
        "header_bg_color": "#70AD47",
        "header_font_color": "#FFFFFF",
        "date_format": "dd-mmm-yy",
        "column_colors": {},
    },
    Source.RGM.value: {
        "font_name": "Arial",
        "font_color": "#000000",
        "header_bg_color": "#4472C4",
        "header_font_color": "#FFFFFF",
        "date_format": "dd/mm/yyyy",
        "column_colors": {},
    },
}


# ─── Session Management ─────────────────────────────────────────────
SESSION_TIMEOUT_SECONDS = 60


# ─── Audit Action Types ─────────────────────────────────────────────
class AuditAction(str, Enum):
    USER_LOGIN = "USER_LOGIN"
    SHIFT_MODIFICATION = "SHIFT_MODIFICATION"
    DATA_EXPORT = "DATA_EXPORT"
    DATA_IMPORT = "DATA_IMPORT"
    USER_CREATED = "USER_CREATED"
    USER_UPDATED = "USER_UPDATED"
    USER_DELETED = "USER_DELETED"
