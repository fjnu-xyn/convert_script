import logging
import os
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Optional

_BASE_DIR = Path(__file__).parent.resolve()
_LOG_DIR = _BASE_DIR / "logs"
_LOG_DIR.mkdir(exist_ok=True)

_LOGGER_NAME = "converter_app"


def _ensure_configured() -> logging.Logger:
    """Configure root app logger once with stdout and rotating file handlers."""
    logger = logging.getLogger(_LOGGER_NAME)
    if logger.handlers:
        return logger

    level_name = os.getenv("LOG_LEVEL", "INFO").upper()
    level = getattr(logging, level_name, logging.INFO)
    logger.setLevel(level)

    fmt = "%(asctime)s %(levelname)s [%(name)s] %(message)s"
    datefmt = "%Y-%m-%d %H:%M:%S"
    formatter = logging.Formatter(fmt=fmt, datefmt=datefmt)

    # Console handler -> stdout (so Streamlit redirect_stdout can capture logs)
    sh = logging.StreamHandler(stream=sys.stdout)
    sh.setLevel(level)
    sh.setFormatter(formatter)
    logger.addHandler(sh)

    # Rotating file handler
    fh = RotatingFileHandler(_LOG_DIR / "app.log", maxBytes=2 * 1024 * 1024, backupCount=3, encoding="utf-8")
    fh.setLevel(level)
    fh.setFormatter(formatter)
    logger.addHandler(fh)

    # Do not propagate to root to avoid duplicate logging
    logger.propagate = False
    return logger


def get_logger(name: Optional[str] = None) -> logging.Logger:
    """Get a child logger under the app logger namespace.

    Example: get_logger("excel_to_word_converter") -> logger named
    "converter_app.excel_to_word_converter".
    """
    _ensure_configured()
    if name:
        return logging.getLogger(f"{_LOGGER_NAME}.{name}")
    return logging.getLogger(_LOGGER_NAME)
