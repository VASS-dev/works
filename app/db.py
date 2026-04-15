from __future__ import annotations

from pathlib import Path

from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, Session

from app import config
from app.models import Base

DEFAULT_DB_PATH = Path(__file__).resolve().parent.parent / "data" / "worksplanning.db"


def make_engine(db_path: Path | str | None = None, echo: bool = False):
    if db_path is not None:
        path = Path(db_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        url = f"sqlite:///{path}"
    else:
        url = config.DB_URL
        if url.startswith("sqlite:///"):
            raw = url[len("sqlite:///"):]
            if raw.startswith("/"):
                p = Path(raw)
            else:
                p = Path(raw)
            p.parent.mkdir(parents=True, exist_ok=True)
    engine = create_engine(url, echo=echo, future=True)
    return engine


def init_db(db_path: Path | str | None = None, drop_first: bool = False):
    engine = make_engine(db_path)
    if drop_first:
        Base.metadata.drop_all(engine)
    Base.metadata.create_all(engine)
    _apply_lightweight_migrations(engine)
    return engine


def _apply_lightweight_migrations(engine) -> None:
    """Добавляет недостающие колонки. Замена Alembic на время Фазы 0-1.

    Для Postgres используем Inspector, для SQLite — PRAGMA (быстрее).
    """
    wanted: dict[str, list[tuple[str, str]]] = {
        "plan_entry": [
            ("status", "VARCHAR(16) NOT NULL DEFAULT 'draft'"),
            ("approved_at", "TIMESTAMP"),
            ("approved_by", "VARCHAR(64)"),
        ],
        "contract": [
            ("contract_no", "VARCHAR(128)"),
            ("contract_date", "DATE"),
        ],
    }
    is_sqlite = engine.dialect.name == "sqlite"
    if is_sqlite:
        with engine.begin() as conn:
            for table, cols in wanted.items():
                existing = {
                    r[1] for r in conn.exec_driver_sql(f"PRAGMA table_info({table})").fetchall()
                }
                for name, ddl in cols:
                    if name not in existing:
                        conn.exec_driver_sql(f"ALTER TABLE {table} ADD COLUMN {name} {ddl}")
    else:
        from sqlalchemy import inspect
        insp = inspect(engine)
        with engine.begin() as conn:
            for table, cols in wanted.items():
                if not insp.has_table(table):
                    continue
                existing = {c["name"] for c in insp.get_columns(table)}
                for name, ddl in cols:
                    if name not in existing:
                        conn.exec_driver_sql(f"ALTER TABLE {table} ADD COLUMN {name} {ddl}")


def make_session_factory(engine) -> sessionmaker[Session]:
    return sessionmaker(bind=engine, autoflush=False, expire_on_commit=False, future=True)
