"""alembic の 0001_initial_schema が models と食い違っていないか（SQLite で upgrade → 比較 → downgrade）。"""
from pathlib import Path

from alembic import command
from alembic.autogenerate import compare_metadata
from alembic.config import Config
from alembic.migration import MigrationContext
from alembic.script import ScriptDirectory
from sqlalchemy import create_engine, inspect

from app.db import Base

API_DIR = Path(__file__).resolve().parents[1]

ALL_TABLES = {
    "requests", "work_logs", "batch_logs", "settings", "dept_approvers", "approver_map_rows", "workers", "orders",
    "company_calendar", "sessions", "google_credentials", "idempotency_keys", "jobs", "sheet_sync_state",
    "sync_runs", "audit_log",
}


def _config(url: str) -> Config:
    cfg = Config(str(API_DIR / "alembic.ini"))
    cfg.set_main_option("script_location", str(API_DIR / "alembic"))
    cfg.set_main_option("sqlalchemy.url", url)
    cfg.attributes["configure_logger"] = False
    return cfg


def test_single_initial_revision():
    script = ScriptDirectory.from_config(_config("sqlite://"))
    revisions = list(script.walk_revisions())
    assert [r.revision for r in revisions] == ["0001"]
    assert Path(revisions[0].path).name == "0001_initial_schema.py"


def test_upgrade_head_matches_models_and_downgrades(tmp_path):
    url = f"sqlite:///{(tmp_path / 'm.db').as_posix()}"
    cfg = _config(url)
    command.upgrade(cfg, "head")

    engine = create_engine(url)
    with engine.connect() as conn:
        diff = compare_metadata(MigrationContext.configure(conn), Base.metadata)
        tables = set(inspect(conn).get_table_names())
        sync_cols = {c["name"] for c in inspect(conn).get_columns("requests")}
    engine.dispose()
    assert diff == []
    assert ALL_TABLES <= tables
    assert set(Base.metadata.tables) == ALL_TABLES
    assert {"row_version", "sheet_row", "cell_kinds", "old_pdf_file_id_pending", "target_date_text"} <= sync_cols

    command.downgrade(cfg, "base")
    engine = create_engine(url)
    with engine.connect() as conn:
        assert set(inspect(conn).get_table_names()) <= {"alembic_version"}
    engine.dispose()


def test_a_tables_have_row_sync_columns():
    for name in ("requests", "work_logs", "batch_logs"):
        cols = Base.metadata.tables[name].c
        assert cols.row_version.nullable is False and cols.row_version.server_default.arg == "1"
        assert cols.sheet_row.nullable is True
        assert cols.cell_kinds.nullable is False


def test_text_columns_are_not_null_default_empty():
    for table in Base.metadata.tables.values():
        for col in table.c:
            if col.type.__class__.__name__ == "Text" and not col.primary_key and col.name not in (
                    "refresh_token_encrypted",):
                assert col.nullable is False, f"{table.name}.{col.name}"
                assert col.server_default is not None, f"{table.name}.{col.name}"
