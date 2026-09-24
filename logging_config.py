"""
Centralised logging setup.

Call `configure_logging()` once at startup (in main.py).
All other modules should do:
    import logging
    logger = logging.getLogger(__name__)
"""

import logging
import sys


def configure_logging(debug: bool = False) -> None:
    """Configure root logger with a structured format."""

    log_level = logging.DEBUG if debug else logging.INFO

    fmt = (
        "%(asctime)s | %(levelname)-8s | %(name)s:%(lineno)d | %(message)s"
    )
    date_fmt = "%Y-%m-%d %H:%M:%S"

    handler = logging.StreamHandler(sys.stdout)
    handler.setLevel(log_level)
    handler.setFormatter(logging.Formatter(fmt, datefmt=date_fmt))

    root_logger = logging.getLogger()
    root_logger.setLevel(log_level)
    # Avoid duplicate handlers if called more than once (e.g. in tests)
    root_logger.handlers.clear()
    root_logger.addHandler(handler)

    # Silence noisy third-party loggers
    logging.getLogger("uvicorn.access").setLevel(logging.WARNING)
