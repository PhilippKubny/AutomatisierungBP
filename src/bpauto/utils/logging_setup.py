"""Utility to provide a shared logger configuration for the project."""

from __future__ import annotations

import logging
import sys
from typing import Final

_LOGGER_NAME: Final = "bpauto"


def setup_logger(level: int = logging.INFO) -> logging.Logger:
    """Return the shared bpauto logger configured for console output."""

    logger = logging.getLogger(_LOGGER_NAME)
    logger.setLevel(level)
    logger.propagate = False

    if not any(isinstance(handler, logging.StreamHandler) for handler in logger.handlers):
        handler = logging.StreamHandler(sys.stdout)
        formatter = logging.Formatter(
            "[%(asctime)s] %(levelname)s - %(message)s",
            "%H:%M:%S",
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)

    return logger
