"""Logging centralisé pour Sophie."""

import logging
import os
from logging.handlers import RotatingFileHandler

PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOG_DIR = os.path.join(PROJECT_DIR, "logs")
os.makedirs(LOG_DIR, exist_ok=True)

_formatter = logging.Formatter(
    "[%(asctime)s] %(levelname)s — %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)


def get_logger(name: str = "sophie") -> logging.Logger:
    """Retourne un logger configuré avec console + fichier rotatif."""
    logger = logging.getLogger(name)

    if logger.handlers:
        return logger

    logger.setLevel(logging.DEBUG)

    # Console
    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    console.setFormatter(_formatter)
    logger.addHandler(console)

    # Fichier rotatif (5 MB, 3 backups)
    file_handler = RotatingFileHandler(
        os.path.join(LOG_DIR, "sophie.log"),
        maxBytes=5 * 1024 * 1024,
        backupCount=3,
        encoding="utf-8",
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(_formatter)
    logger.addHandler(file_handler)

    return logger
