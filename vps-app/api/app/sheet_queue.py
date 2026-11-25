"""行の反映の印付け（決定 A・実装仕様 6-2 の 1）。

W・M・D の業務関数は、A 区分の表（requests・work_logs・batch_logs）の行を書いたら、同じトランザクションで
mark_dirty(db, table_name, pk) を呼ぶ。row_version を +1 し、sheet_sync_state.desired_version をそれに合わせる
（新しい行は sheet_sync_state を作り、seq＝作った順を振る）。worker の行の反映のレーン（担当 D の row_sync）が、
desired_version > synced_version の行を、DB の今の値でシートへ書く。

- 新しい work_logs・batch_logs の行は、先に db.flush() して id を得てから呼ぶ（pk は str(id)）。
- commit は呼び出し側（rpc.py など）。
"""
from __future__ import annotations

from typing import Any

from sqlalchemy import func, select
from sqlalchemy.orm import Session

from app import timeutil
from app.models import SHEET_SYNC_SEQ, SYNCED_TABLES, SheetSyncState


def _model(table_name: str):
    model = SYNCED_TABLES.get(table_name)
    if model is None:
        raise ValueError(f"行の反映の対象ではない表です: {table_name}")
    return model


def _pk_value(table_name: str, pk: Any) -> Any:
    if table_name == "requests":
        return str(pk)
    try:
        return int(pk)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{table_name} の pk は数です: {pk!r}") from exc


def _load(db: Session, table_name: str, pk: Any):
    row = db.get(_model(table_name), _pk_value(table_name, pk))
    if row is None:
        raise ValueError(f"{table_name} に行がありません: {pk}")
    return row


def next_seq(db: Session) -> int:
    """新しい行の作った順の番号。PostgreSQL は連番、SQLite は今の最大 + 1。"""
    if db.get_bind().dialect.name == "postgresql":
        return int(db.execute(SHEET_SYNC_SEQ.next_value()).scalar_one())
    current = db.scalar(select(func.max(SheetSyncState.seq))) or 0
    return int(current) + 1


def mark_dirty(db: Session, table_name: str, pk: str) -> int:
    """行を「シートへ書く」印にする。新しい row_version を返す（同じトランザクション。commit しない）。"""
    row = _load(db, table_name, pk)
    row.row_version = int(row.row_version or 0) + 1
    db.flush()
    key = str(_pk_value(table_name, pk))
    state = db.get(SheetSyncState, (table_name, key))
    now = timeutil.now_utc()
    if state is None:
        state = SheetSyncState(table_name=table_name, pk=key, desired_version=row.row_version, synced_version=0,
                               synced_values={}, state="ok", attempts=0, next_try_at=now, last_error="",
                               first_dirty_at=now, seq=next_seq(db))
        db.add(state)
    else:
        was_clean = state.desired_version <= state.synced_version
        state.desired_version = row.row_version
        if was_clean or state.first_dirty_at is None:
            state.first_dirty_at = now
            state.next_try_at = now
            state.attempts = 0
            if state.state == "retry":
                state.state = "ok"
    db.flush()
    return row.row_version


def current_version(db: Session, table_name: str, pk: str) -> int:
    """その行の今の row_version（GAS の仕事の requires を作るため。担当 M）。"""
    return int(_load(db, table_name, pk).row_version)


def requires_for(db: Session, *rows: tuple[str, str]) -> list[dict]:
    """[(table, pk), ...] から requires の形 [{"table", "pk", "version"}] を作る小道具。"""
    return [{"table": t, "pk": str(p), "version": current_version(db, t, p)} for t, p in rows]


def is_synced(db: Session, table_name: str, pk: str, version: int) -> bool:
    """その版まで行の反映が済んだか（jobs の requires の確かめ）。"""
    state = db.get(SheetSyncState, (table_name, str(pk)))
    return state is not None and int(state.synced_version) >= int(version)
