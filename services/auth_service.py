# services/auth_service.py
# Authentication service - coordinates login and credential management.

import logging
from typing import Optional, Dict, Tuple, List

import database_logic as db

logger = logging.getLogger(__name__)


class AuthService:
    """Handles authentication and credential management."""

    @staticmethod
    def login(username: str, password: str) -> Optional[Dict]:
        """
        Authenticate a user.
        Returns user info dict on success, None on failure.
        """
        if not username or not password:
            return None
        result = db.authenticate_user(username, password)
        if result:
            logger.info("User '%s' authenticated as %s", username, result["role"])
        else:
            logger.warning("Failed login attempt for user '%s'", username)
        return result

    @staticmethod
    def create_user(username: str, password: str, role: str,
                    excel_file: str = "", can_manage: bool = False) -> Tuple[bool, str]:
        """Create a new credential."""
        return db.create_credential(username, password, role, excel_file, can_manage)

    @staticmethod
    def change_password(username: str, new_password: str) -> Tuple[bool, str]:
        """Change a user's password."""
        return db.update_credential_password(username, new_password)

    @staticmethod
    def delete_user(username: str) -> Tuple[bool, str]:
        """Delete a credential."""
        return db.delete_credential(username)

    @staticmethod
    def list_users() -> List[Dict]:
        """List all credentials (without password hashes)."""
        return db.get_all_credentials()
