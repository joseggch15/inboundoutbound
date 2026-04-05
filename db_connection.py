# db_connection.py
# Context manager for SQLite database connections.
# Eliminates the repetitive connect/cursor/commit/close boilerplate.

import sqlite3
from contextlib import contextmanager
from typing import Optional


# DB_FILE is set by database_logic at import time.
# We import it lazily to avoid circular imports.
_db_file: Optional[str] = None


def set_db_file(path: str) -> None:
    """Called once by database_logic to register the DB path."""
    global _db_file
    _db_file = path


def get_db_file() -> str:
    if _db_file is None:
        raise RuntimeError("DB file path not initialized. Call set_db_file() first.")
    return _db_file


@contextmanager
def get_connection(row_factory: bool = False):
    """
    Context manager that yields a sqlite3 Connection.

    Usage:
        with get_connection() as conn:
            cur = conn.cursor()
            cur.execute(...)

        # conn.commit() is called automatically on success.
        # conn.rollback() is called on exception.
        # conn.close() is always called.
    """
    conn = sqlite3.connect(get_db_file())
    if row_factory:
        conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


@contextmanager
def get_cursor(row_factory: bool = False):
    """
    Context manager that yields a sqlite3 Cursor (convenience wrapper).

    Usage:
        with get_cursor(row_factory=True) as cur:
            cur.execute("SELECT ...")
            rows = cur.fetchall()
    """
    with get_connection(row_factory=row_factory) as conn:
        yield conn.cursor()


def ensure_cursor(external_cursor: Optional[sqlite3.Cursor] = None):
    """
    For functions that support external cursor (transaction participation).

    Usage:
        with ensure_cursor(external_cursor) as (cur, should_commit):
            cur.execute(...)
            if should_commit:
                cur.connection.commit()
    """
    @contextmanager
    def _wrapper():
        if external_cursor is not None:
            yield external_cursor, False
        else:
            conn = sqlite3.connect(get_db_file())
            cur = conn.cursor()
            try:
                yield cur, True
                conn.commit()
            except Exception:
                conn.rollback()
                raise
            finally:
                conn.close()

    return _wrapper()
