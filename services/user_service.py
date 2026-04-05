# services/user_service.py
# User management business logic.

import logging
from datetime import date
from typing import Optional, Dict, List, Tuple

import database_logic as db

logger = logging.getLogger(__name__)


class UserService:
    """Business logic for user/employee management."""

    @staticmethod
    def get_users(source: str) -> List[Dict]:
        """Get all users for a source, ordered by display_order."""
        return db.get_all_users(source)

    @staticmethod
    def get_users_with_locations(source: str) -> List[Dict]:
        """Get all users with their default pickup/dropoff locations."""
        return db.get_users_with_defaults(source)

    @staticmethod
    def add_user(name: str, role: str, badge: str, source: str) -> Tuple[bool, str]:
        """Add a new employee."""
        return db.add_user(name, role, badge, source)

    @staticmethod
    def update_user(user_id: int, name: str, role: str, badge: str, source: str) -> Tuple[bool, str, Optional[str]]:
        """Update user with cascade badge rename if needed."""
        return db.update_user_with_cascade(user_id, name, role, badge, source)

    @staticmethod
    def delete_user(user_id: int) -> Tuple[bool, str]:
        """Delete an employee."""
        return db.delete_user(user_id)

    @staticmethod
    def import_users_bulk(users: list, source: str) -> int:
        """Bulk import users, skipping duplicates."""
        return db.add_users_bulk(users, source)

    @staticmethod
    def get_roles(source: str) -> List[Dict]:
        """Get all roles for a source."""
        return db.get_roles(source)

    @staticmethod
    def get_locations(source: Optional[str] = None) -> List[Dict]:
        """Get all locations, optionally filtered by source."""
        return db.get_locations(source)

    @staticmethod
    def set_default_locations(badge: str, pickup: Optional[str], dropoff: Optional[str]) -> None:
        """Set default pickup/dropoff locations for a user."""
        db.set_user_default_locations(badge, pickup, dropoff)

    @staticmethod
    def move_user(source: str, badge_a: str, badge_b: str) -> Tuple[bool, str]:
        """Swap display order of two users."""
        return db.swap_user_order(source, badge_a, badge_b)

    @staticmethod
    def get_ordered_badges(source: str) -> List[str]:
        """Get badges in display order."""
        return db.get_ordered_badges(source)
