"""GAS の後処理の仕事を積む口（実装仕様 7-4・10-2）。

担当 M の services.gas_jobs.enqueue_* を 10-2 の名前と引数で呼ぶ。M のモジュールがまだ無いブランチでは、
同じ kind 名（7-2 の action 名）で jobs.enqueue に直接積む（M が登録するハンドラーが後で拾う）。
どちらの場合も、呼ぶ前に行を書いて sheet_queue.mark_dirty を済ませておく（requires はその版を指す）。

管理者の関数（ops）で M・D の仕事を呼ぶ所は service() を使う（モジュールが無ければ「まだ用意されていません」の
RpcError）。
"""
from __future__ import annotations

import importlib
import time
from types import ModuleType
from typing import Any

from sqlalchemy.orm import Session

from app import jobs, sheet_queue
from app.constants import PRIORITY_SEBASTIAN
from app.models import Job
from app.rpc import RpcError

from . import store

KIND_AFTER_APPROVE = "internal.afterApprove"
KIND_AFTER_COMPLETE = "internal.afterComplete"
KIND_REGENERATE = "internal.regenerate"
KIND_LEDGER_UPSERT = "internal.ledgerUpsert"
KIND_AFTER_DELETE = "internal.afterDelete"


def _optional(name: str) -> ModuleType | None:
    try:
        return importlib.import_module(f"app.services.{name}")
    except ModuleNotFoundError as exc:
        if exc.name in (f"app.services.{name}",):
            return None
        raise


def service(name: str) -> ModuleType:
    """担当 D・M のモジュール（無ければ RpcError）。"""
    module = _optional(name)
    if module is None:
        raise RpcError(f"この機能はまだ用意されていません。（services.{name}）")
    return module


def _gas_jobs() -> ModuleType | None:
    return _optional("gas_jobs")


# ===== M が無いときの積み方（kind 名で直接） =====


def _requires(db: Session, request_id: str, with_worklog: bool = False) -> list[dict]:
    rows = [("requests", request_id)]
    if with_worklog:
        wl = store.first_worklog(db, request_id)
        if wl is not None:
            rows.append(("work_logs", str(wl.id)))
    return sheet_queue.requires_for(db, *rows)


def _status(db: Session, request_id: str) -> str:
    from app.models import Request

    req = db.get(Request, request_id)
    return req.status if req is not None else ""


def _fallback(db: Session, kind: str, request_id: str, args: dict, *, with_worklog: bool = False,
              priority: int = 0, batch_id: str = "") -> int:
    return jobs.enqueue(db, kind, args, lane="gas", request_id=request_id, priority=priority,
                        requires=_requires(db, request_id, with_worklog),
                        guard={"status": _status(db, request_id)}, batch_id=batch_id)


# ===== 10-2 の関数 =====


def after_approve(db: Session, request_id: str) -> int:
    m = _gas_jobs()
    if m is not None:
        return m.enqueue_after_approve(db, request_id)
    return _fallback(db, KIND_AFTER_APPROVE, request_id, {})


def after_complete(db: Session, request_id: str, *, source: str, status_at_enqueue: str,
                   wl_override: dict | None = None, priority: int = 0) -> int:
    m = _gas_jobs()
    if m is not None:
        return m.enqueue_after_complete(db, request_id, source=source, status_at_enqueue=status_at_enqueue,
                                        wl_override=wl_override, priority=priority)
    args: dict[str, Any] = {"source": source, "statusAtEnqueue": status_at_enqueue}
    if wl_override is not None:
        args["wlOverride"] = wl_override
    return _fallback(db, KIND_AFTER_COMPLETE, request_id, args, with_worklog=True, priority=priority)


def regenerate(db: Session, request_id: str, *, reason: str, old_pdf_file_id: str = "", batch_id: str = "") -> int:
    m = _gas_jobs()
    if m is not None:
        return m.enqueue_regenerate(db, request_id, reason=reason, old_pdf_file_id=old_pdf_file_id,
                                    batch_id=batch_id)
    return _fallback(db, KIND_REGENERATE, request_id, {"reason": reason, "oldPdfFileId": old_pdf_file_id},
                     batch_id=batch_id)


def ledger_upsert(db: Session, request_id: str, *, batch_id: str = "") -> int:
    m = _gas_jobs()
    if m is not None:
        return m.enqueue_ledger_upsert(db, request_id, batch_id=batch_id)
    return _fallback(db, KIND_LEDGER_UPSERT, request_id, {}, batch_id=batch_id)


def after_delete(db: Session, request_id: str, *, request_type: str, target_date: str, old_pdf_file_id: str) -> int:
    m = _gas_jobs()
    if m is not None:
        return m.enqueue_after_delete(db, request_id, request_type=request_type, target_date=target_date,
                                      old_pdf_file_id=old_pdf_file_id)
    return _fallback(db, KIND_AFTER_DELETE, request_id, {"oldPdfFileId": old_pdf_file_id,
                                                         "requestType": request_type, "targetDate": target_date})


SEBASTIAN_PRIORITY = PRIORITY_SEBASTIAN


def wait_for_result(session_factory, job_id: int, timeout_seconds: float) -> dict | None:
    """M の gas_jobs.wait_for_result（無ければ jobs の行を見て待つ）。None は待ち切れ。"""
    m = _gas_jobs()
    if m is not None:
        return m.wait_for_result(session_factory, job_id, timeout_seconds)
    import json

    deadline = time.monotonic() + max(0.0, float(timeout_seconds))
    while True:
        with session_factory() as s:
            job = s.get(Job, int(job_id))
            if job is not None and job.state in ("done", "failed", "skipped"):
                try:
                    result = json.loads(job.result_json or "{}")
                except ValueError:
                    result = {}
                if not isinstance(result, dict):
                    result = {}
                if job.state != "done":
                    result.setdefault("ok", False)
                    result.setdefault("error", job.last_error or job.state)
                return result
        if time.monotonic() >= deadline:
            return None
        time.sleep(0.5)
