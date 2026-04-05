# services/report_service.py
# Report generation business logic.

import logging
from typing import Dict

import database_logic as db

logger = logging.getLogger(__name__)


class ReportService:
    """Business logic for report settings and generation."""

    @staticmethod
    def get_settings(username: str, source: str) -> Dict:
        """Get report settings for a user and source."""
        return db.get_report_settings(username, source)

    @staticmethod
    def save_settings(username: str, source: str, settings: Dict) -> None:
        """Save report settings."""
        db.save_report_settings(username, source, settings)

    @staticmethod
    def reset_settings(username: str, source: str) -> None:
        """Reset report settings to defaults."""
        db.delete_report_settings(username, source)
