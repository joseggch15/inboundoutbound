# services/__init__.py
# Business logic services extracted from UI widgets.
# Services coordinate between database_logic and excel_logic
# without knowing about PyQt or UI concerns.

from services.schedule_service import ScheduleService
from services.user_service import UserService
from services.report_service import ReportService
from services.auth_service import AuthService
