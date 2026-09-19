"""Run from project root:  python -m scripts.check_db"""
from sqlalchemy import inspect, text

from db.database import DB_MODE, engine, init_db

init_db()
with engine.connect() as conn:
    conn.execute(text("SELECT 1"))
print(f"DB mode : {DB_MODE}  (expected: shared)")
print(f"Tables  : {sorted(inspect(engine).get_table_names())}")
