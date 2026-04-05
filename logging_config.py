# logging_config.py
# Centralized logging configuration for the application.
# Replaces silent `except Exception: pass` with proper logging.

import logging
import sys
from pathlib import Path


def setup_logging(level: int = logging.INFO) -> None:
    """
    Configure application-wide logging.

    Logs go to:
      1. Console (stderr) — for development
      2. app.log file — for production diagnostics
    """
    log_dir = Path(__file__).resolve().parent
    log_file = log_dir / "app.log"

    root_logger = logging.getLogger()

    # Avoid adding handlers multiple times if called again
    if root_logger.handlers:
        return

    root_logger.setLevel(level)

    formatter = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # Console handler
    console_handler = logging.StreamHandler(sys.stderr)
    console_handler.setLevel(logging.WARNING)
    console_handler.setFormatter(formatter)
    root_logger.addHandler(console_handler)

    # File handler
    try:
        file_handler = logging.FileHandler(str(log_file), encoding="utf-8")
        file_handler.setLevel(level)
        file_handler.setFormatter(formatter)
        root_logger.addHandler(file_handler)
    except (OSError, PermissionError):
        # If we can't write to file, just use console
        console_handler.setLevel(level)


def get_logger(name: str) -> logging.Logger:
    """Get a named logger for a module."""
    return logging.getLogger(name)
