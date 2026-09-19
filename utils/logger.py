"""utils/logger.py — Shared logger. Usage:  log = get_logger(__name__)
RULE: log metadata only (sizes, counts, timings, ids). Never raw text, keys or DATABASE_URL."""
import logging
import os

os.makedirs("logs", exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    handlers=[logging.FileHandler("logs/app.log", encoding="utf-8"), logging.StreamHandler()],
)
# Silence noisy third-party loggers
for noisy in ("httpx", "httpcore", "openai", "sqlalchemy.engine"):
    logging.getLogger(noisy).setLevel(logging.WARNING)


def get_logger(name: str) -> logging.Logger:
    return logging.getLogger(name)
