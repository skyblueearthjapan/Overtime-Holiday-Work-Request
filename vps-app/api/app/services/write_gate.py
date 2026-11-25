"""シートへ書く前の関所（データ安全 高-3）。

GAS の移転モードと VPS の書き込みを連動させる。行の反映（row_sync）・毎晩の照合（reconcile）・マスタ同期
（master_sync）は、シートへ書く前に check() を通す。

- production: 最後に合格した gas_ping（sync_runs kind='gas_ping'、ok=True）が SHEET_WRITE_PING_MAX_MINUTES 分以内で、
  その返事の migrated（migrationMode）が true のときだけ書ける。GAS が移転モードでない（戻した後・切り替えの前）なら
  書かない。GAS の側でプロパティが読めないときも migrated:false が返るので、書かない側に倒れる。
- staging: 同じ時間内に、ステージングの GAS（verify_gas_ping がステージングの scriptId・ID を確かめて合格にした
  ping）に届いていれば書ける（写しの GAS の移転モードは問わない。ステージングでの試しを止めないため）。

止めた理由は呼ぶ側が sync_runs に残し、ops_getStatus の sheetWriteGate に出る（管理者の欄）。
"""
from __future__ import annotations

from datetime import datetime, timedelta
from typing import Any

from sqlalchemy import select
from sqlalchemy.orm import Session

from app import models, timeutil

GAS_PING_KIND = "gas_ping"


def _max_age(settings: Any) -> timedelta:
    return timedelta(minutes=int(getattr(settings, "sheet_write_ping_max_minutes", 10) or 10))


def last_ok_ping(db: Session) -> models.SyncRun | None:
    R = models.SyncRun
    return db.scalars(select(R).where(R.kind == GAS_PING_KIND, R.ok.is_(True))
                      .order_by(R.started_at.desc(), R.id.desc()).limit(1)).first()


def evaluate(db: Session, settings: Any, now: datetime | None = None) -> tuple[bool, str]:
    """シートへ書いてよいか（理由つき）。試験では conftest がこの関数を差し替えない（check を差し替える）。"""
    now = now or timeutil.now_utc()
    row = last_ok_ping(db)
    env = "staging" if getattr(settings, "is_staging", False) else "production"
    if row is None:
        return False, f"合格した gas_ping がまだありません（{env}）。シートへの書き込みを止めています"
    at = row.finished_at or row.started_at
    limit = _max_age(settings)
    if now - at > limit:
        return False, (f"合格した gas_ping が {int(limit.total_seconds() // 60)} 分以上ありません（最後 {at.isoformat()}）。"
                       "シートへの書き込みを止めています")
    if env == "production":
        migrated = (row.summary_json or {}).get("migrationMode")
        if migrated is not True:
            return False, ("本番の GAS が移転モードではありません（migrated が true ではない）。GAS と VPS が同じシートを"
                           "書き合わないよう、シートへの書き込みを止めています")
    return True, ""


def check(db: Session, settings: Any, now: datetime | None = None) -> tuple[bool, str]:
    """呼ぶ側はこれを使う（試験の土台が差し替える）。"""
    return evaluate(db, settings, now)


def status(db: Session, settings: Any, now: datetime | None = None) -> dict:
    """管理者の欄に出す形。"""
    ok, reason = evaluate(db, settings, now)
    row = last_ok_ping(db)
    return {"allowed": ok, "blocked": not ok, "reason": reason,
            "lastOkPingAt": (row.finished_at or row.started_at).isoformat() if row else "",
            "migrated": (row.summary_json or {}).get("migrationMode") if row else None}
