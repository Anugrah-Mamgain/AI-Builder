"""db/database.py — Supabase (shared PostgreSQL) with automatic local SQLite fallback.  Owner: P2"""
import os

from dotenv import load_dotenv
from sqlalchemy import create_engine, text
from sqlalchemy.orm import declarative_base, sessionmaker

import config as cfg
from utils.logger import get_logger

load_dotenv()
log = get_logger(__name__)
Base = declarative_base()


def _make_engine():
    url = os.getenv("DATABASE_URL", "").strip()
    if url:
        if url.startswith("postgres://"):            # some providers use the old prefix
            url = url.replace("postgres://", "postgresql://", 1)
        try:
            eng = create_engine(url, pool_pre_ping=True, pool_size=5, max_overflow=5,
                                connect_args={"connect_timeout": 5})
            with eng.connect() as conn:
                conn.execute(text("SELECT 1"))        # fail fast if unreachable
            log.info("Connected to shared PostgreSQL (Supabase)")
            return eng, "shared"
        except Exception as e:
            log.warning(f"PostgreSQL unavailable ({type(e).__name__}); using local SQLite")
    os.makedirs(os.path.dirname(cfg.SQLITE_PATH), exist_ok=True)
    eng = create_engine(f"sqlite:///{cfg.SQLITE_PATH}", connect_args={"check_same_thread": False})
    log.info("Using local SQLite")
    return eng, "local"


engine, DB_MODE = _make_engine()
SessionLocal = sessionmaker(bind=engine, expire_on_commit=False)


def init_db():
    from db import models  # noqa: F401  (registers tables)
    Base.metadata.create_all(bind=engine)
