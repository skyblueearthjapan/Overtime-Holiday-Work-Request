"""管理者の「運用」欄の中身（実装仕様 3-10・4-2 の ops_getStatus。担当 M）。

get_ops_status(db) は JSON にできる dict を返す（日時は JST の 'yyyy-MM-dd HH:mm:ss'）。W の ops_getStatus がそのまま返す。
- rowSync: 行の反映の待ち件数・最も古い待ちの経過秒・5 分を超えた遅れの件数・止まっている行（blocked）
- gasJobs: GAS の仕事の状態ごとの件数、失敗の一覧（id・申請 ID・action・理由・回数）、前の仕事の失敗で止まっている仕事、
  待ち（pending）で理由のあるもの（行の反映待ち・gas_ping 待ちなど）
- gasPing: 最後の gas_ping の時刻・結果・合わない点
- lastRuns: 朝バッチ・手動の PDF 一括・7:05 の確かめ・マスタ同期・共有タブ取り込み・毎晩の照合の最後の結果
- reconcileFixedRows: 最後の照合で直した行の数（照合の summary に数があれば）
- exhausted: 再試行を使い切った仕事の記録（新しい順に最大 20）
- pdfMode: Settings の PDF_MODE（表示だけ。8-7）
"""
from __future__ import annotations

from datetime import datetime, timedelta
from typing import Any

from sqlalchemy import func, select
from sqlalchemy.orm import Session

from app import jobs, models, timeutil
from app.config import get_settings
from app.services import gas_jobs, write_gate

ROW_SYNC_DELAY_ALERT = timedelta(minutes=5)
LAST_RUN_KINDS = ("morning_batch", "manual_batch_pdfs", "morning_batch_check", "sync_masters", "sync_sheet_config",
                  "sheet_reconcile")
FIXED_ROW_KEYS = ("fixedRows", "fixed_rows", "fixed", "repairedRows", "repaired")


def _t(value: datetime | None) -> str:
    return timeutil.to_jst_text(value, "yyyy-MM-dd HH:mm:ss") if value else ""


def _plain(value: Any) -> Any:
    if isinstance(value, datetime):
        return _t(value)
    if isinstance(value, dict):
        return {k: _plain(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [_plain(v) for v in value]
    return value


def _row_sync(db: Session, now: datetime) -> dict:
    S = models.SheetSyncState
    dirty = S.desired_version > S.synced_version
    pending = int(db.scalar(select(func.count()).select_from(S).where(dirty)) or 0)
    oldest = db.scalar(select(func.min(S.first_dirty_at)).where(dirty))
    oldest_seconds = int((now - oldest).total_seconds()) if oldest else 0
    delayed = int(db.scalar(select(func.count()).select_from(S).where(
        dirty, S.first_dirty_at < now - ROW_SYNC_DELAY_ALERT)) or 0)
    blocked = [{"table": r.table_name, "pk": r.pk, "error": r.last_error or "", "attempts": r.attempts,
                "since": _t(r.first_dirty_at)}
               for r in db.scalars(select(S).where(S.state == "blocked").order_by(S.first_dirty_at).limit(100))]
    return {"pending": pending, "oldestPendingSeconds": oldest_seconds, "delayed": delayed, "blocked": blocked}


def _job_row(j: dict) -> dict:
    return {"id": j["id"], "requestId": j["requestId"], "action": j["kind"], "reason": j["error"],
            "attempts": j["attempts"], "state": j["state"], "createdAt": _t(j["createdAt"]),
            "runAfter": _t(j["runAfter"])}


def _gas_jobs(db: Session) -> dict:
    J = models.Job
    waiting = [
        {"id": r.id, "requestId": r.request_id, "action": r.kind, "reason": r.last_error, "attempts": r.attempts,
         "runAfter": _t(r.run_after)}
        for r in db.scalars(select(J).where(J.lane == "gas", J.state == "pending", J.last_error != "")
                            .order_by(J.id).limit(100))
    ]
    return {
        "counts": jobs.counts_by_state(db, lane="gas"),
        "failed": [_job_row(j) for j in jobs.list_failed(db, lane="gas")],
        "blocked": [_job_row(j) for j in jobs.list_blocked(db)],
        "waiting": waiting,
    }


def _last_run(db: Session, kind: str) -> dict | None:
    R = models.SyncRun
    row = db.scalars(select(R).where(R.kind.in_((kind, "schedule:" + kind)))
                     .order_by(R.started_at.desc(), R.id.desc()).limit(1)).first()
    if row is None:
        return None
    return {"kind": row.kind, "startedAt": _t(row.started_at), "finishedAt": _t(row.finished_at), "ok": bool(row.ok),
            "summary": _plain(row.summary_json or {})}


def _gas_ping(db: Session, now: datetime) -> dict:
    row = gas_jobs.last_gas_ping(db)
    ok, reason = gas_jobs.gas_ping_ok(db, now)
    if row is None:
        return {"at": "", "ok": False, "fresh": False, "problems": [reason], "lanePaused": True}
    summary = row.summary_json or {}
    return {"at": _t(row.finished_at or row.started_at), "ok": bool(row.ok), "fresh": ok,
            "problems": list(summary.get("problems") or []), "lanePaused": not ok, "reason": reason,
            "pdfMode": summary.get("pdfMode", ""), "version": summary.get("version")}


def _fixed_rows(run: dict | None) -> int | None:
    if not run:
        return None
    summary = run.get("summary") or {}
    for key in FIXED_ROW_KEYS:
        value = summary.get(key)
        if isinstance(value, bool):
            continue
        if isinstance(value, int):
            return value
        if isinstance(value, list):
            return len(value)
    return None


def _exhausted(db: Session) -> list[dict]:
    R = models.SyncRun
    rows = db.scalars(select(R).where(R.kind == gas_jobs.EXHAUSTED_KIND).order_by(R.started_at.desc(), R.id.desc())
                      .limit(20))
    return [{"at": _t(r.started_at), **_plain(r.summary_json or {})} for r in rows]


def get_ops_status(db: Session) -> dict:
    now = timeutil.now_utc()
    last_runs = {kind: _last_run(db, kind) for kind in LAST_RUN_KINDS}
    pdf_mode = db.get(models.Setting, "PDF_MODE")
    return {
        "now": _t(now),
        "rowSync": _row_sync(db, now),
        "gasJobs": _gas_jobs(db),
        "gasPing": _gas_ping(db, now),
        # シートへ書く前の関所（データ安全 高-3）。blocked なら行の反映・照合・マスタ同期は書かない
        "sheetWriteGate": write_gate.status(db, get_settings(), now),
        "lastRuns": last_runs,
        "reconcileFixedRows": _fixed_rows(last_runs.get("sheet_reconcile")),
        "exhausted": _exhausted(db),
        "pdfMode": pdf_mode.value if pdf_mode is not None else "",
    }
