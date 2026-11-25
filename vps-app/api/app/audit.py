"""書き込みの記録（実装仕様 4-1）。rpc.py が書き込みの関数を呼んだとき、業務と同じトランザクションで残す。

失敗した呼び出しは、業務の書き込みを rollback した後で、別のトランザクションで ok=False として残す。
audit_log は消さない（worker の purge も触らない）。
"""
from __future__ import annotations

import json
from collections.abc import Iterable

from sqlalchemy.orm import Session

from app import timeutil
from app.models import AuditLog


def record(db: Session, user_email: str, fn: str, request_ids: Iterable[str], ok: bool, error: str = "") -> AuditLog:
    """1 行足す（commit は呼び出し側）。"""
    ids = [str(i) for i in request_ids if i is not None and str(i) != ""]
    row = AuditLog(
        at=timeutil.now_utc(),
        user_email=(user_email or "").strip().lower()[:256],
        fn=fn[:64],
        request_ids=json.dumps(ids, ensure_ascii=False),
        ok=bool(ok),
        error=(error or "")[:2000],
    )
    db.add(row)
    return row
