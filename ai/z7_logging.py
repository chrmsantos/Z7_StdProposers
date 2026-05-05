import logging
from logging.handlers import RotatingFileHandler
import os
from datetime import datetime
from pathlib import Path


def get_runtime_dir() -> Path:
    user_profile = os.environ.get("USERPROFILE")
    if not user_profile:
        return Path.cwd()

    runtime_dir = Path(user_profile) / "AppData" / "Local" / "Z7" / "Tmp" / "StdProposers"
    runtime_dir.mkdir(parents=True, exist_ok=True)
    return runtime_dir


def get_logs_dir() -> Path:
    logs_dir = get_runtime_dir() / "logs"
    logs_dir.mkdir(parents=True, exist_ok=True)
    return logs_dir


def build_log_path(component: str) -> Path:
    safe_component = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in component)
    return get_logs_dir() / f"{safe_component}.log"


def configure_component_logger(component: str, level: int = logging.INFO) -> logging.Logger:
    logger = logging.getLogger(f"z7.{component}")
    if logger.handlers:
        return logger

    logger.setLevel(level)
    logger.propagate = False

    formatter = logging.Formatter(
        fmt="%(asctime)s.%(msecs)03d | %(levelname)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    log_path = build_log_path(component)
    file_handler = RotatingFileHandler(log_path, maxBytes=2 * 1024 * 1024, backupCount=3, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    return logger


def log_exception(logger: logging.Logger, context: str, exc: Exception) -> None:
    logger.exception("%s: %s", context, exc)
