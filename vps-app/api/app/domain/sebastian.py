"""Sebastian の入口 POST /api/sebastian（実装仕様 4-4。src/api.gs の doPost と同じ JSON の取り決め）。

- 経路は main.py（担当 F）。ここは handle(db, raw_body) -> dict だけ。返事はいつも HTTP 200。
- 本文は Content-Type に関係なく生の文字を json.loads（空なら '{}'）。トークンは hmac.compare_digest。
- 返事の文字は api.gs の文字どおり（末尾に ': <requestId>' が付く文、{error, action} で ok が無い形、
  JSON の誤り・トークン未設定・不一致の 3 つは action も無い）。'invalid JSON: ' の後ろは Python の例外文
  （受け入れる差）。
- completeOvertime: 開始＝対象日 17:20 JST、終了＝今（manualEndTime は無視）、休憩は画面と同じ時刻窓方式（9-1 #23）、
  updatedBy='sebastian-agent'。internal.afterComplete（source=sebastian、優先度 10、wlOverride なし）。
- markHolidayStart・markHolidayDone: 画面の 9・10 と同じ計算。本人確認はしない（今と同じ）。
- 返す status は書く前に読んだ値、workerName は生の値、targetDate はシートの型どおり。
- pdf: status が approved でなければ null。approved なら、積んだ仕事を確定（commit）してから結果を待つ。
"""
from __future__ import annotations

import hmac
import json
import logging
from datetime import datetime
from typing import Any

from sqlalchemy.orm import Session, sessionmaker

from app import timeutil
from app.config import Settings, get_settings
from app.jsutil import js_string
from app.models import Request
from app.rpc import RpcError

from . import gasjobs, store, worklog
from .jsdate import iso_utc_ms, js_truthy

log = logging.getLogger(__name__)
UPDATED_BY = "sebastian-agent"
PDF_TIMEOUT_ERROR = "PDF作成の結果待ちが時間切れになりました（作成は続いています）"
# 想定外の例外の返事（内部の例外の文は返さない。詳細はログだけ。セキュリティ確認 中-1）
INTERNAL_ERROR = "internal error"
# 確定の後に PDF の結果待ちで例外が出たときの pdf（保存は済んでいるので ok:true のまま返す。正しさ 低-19）
PDF_WAIT_ERROR = "PDF作成の結果を確かめられませんでした（作成は続いている可能性があります）"


def target_date_value(req: Request, settings: Settings) -> Any:
    """シートの targetDate の型どおり: 日付のセルなら DB_SHEET_TIMEZONE の 0 時を UTC にした ISO、文字ならその文字。"""
    kind = (req.cell_kinds or {}).get("targetDate")
    if kind == "string" or req.target_date is None:
        return store.target_raw_text(req)
    midnight = datetime(req.target_date.year, req.target_date.month, req.target_date.day, tzinfo=settings.db_sheet_tz)
    return iso_utc_ms(midnight)


def pdf_from_result(result: dict | None) -> dict | None:
    """GAS の内部用の入口の返事（または仕事の結果）を、今の Sebastian の pdf の形にする。"""
    if result is None:
        return {"error": PDF_TIMEOUT_ERROR}
    if not isinstance(result, dict):
        return {"error": "unknown result"}
    pdf = result.get("pdf") if isinstance(result.get("pdf"), dict) else None
    if pdf is None and "status" in result and ("pdfFileId" in result or "error" in result):
        pdf = result
    if pdf is not None:
        status = pdf.get("status")
        if status == "created":
            return {"ok": True, "pdfFileId": pdf.get("pdfFileId", ""), "pdfName": pdf.get("pdfName", ""),
                    "folderId": pdf.get("folderId", "")}
        if status == "already":
            return {"already": True, "pdfFileId": pdf.get("pdfFileId", "")}
        if status == "error":
            return {"error": js_string(pdf.get("error") or result.get("message") or "PDF error")}
        if status == "skipped":
            return None
    if result.get("already"):
        return {"already": True, "pdfFileId": result.get("pdfFileId", "")}
    if result.get("ok") is True and "pdfFileId" in result:
        return {"ok": True, "pdfFileId": result.get("pdfFileId", ""), "pdfName": result.get("pdfName", ""),
                "folderId": result.get("folderId", "")}
    if result.get("ok") is False or "error" in result:
        return {"error": js_string(result.get("message") or result.get("error") or "PDF error")}
    return None


def _wait_pdf(db: Session, job_id: int, settings: Settings) -> dict | None:
    """WorkLogs は確定済み。ここでの例外は外へ出さず pdf:{error} にする（{error, action} を返すと
    Sebastian が再送して二重の完了になる恐れがある）。"""
    try:
        factory = sessionmaker(bind=db.get_bind(), expire_on_commit=False)
        result = gasjobs.wait_for_result(factory, job_id, settings.sebastian_pdf_wait_seconds)
        return pdf_from_result(result)
    except Exception:  # noqa: BLE001 - 詳細はログだけ
        log.exception("sebastian: waiting for pdf job %s failed", job_id)
        return {"error": PDF_WAIT_ERROR}


def _complete_overtime(db: Session, request_id: str, now: datetime, settings: Settings) -> dict:
    done = worklog.core_overtime_done(db, request_id, None, now, msgs=worklog.SEBASTIAN, email=None,
                                      updated_by=UPDATED_BY, source="sebastian",
                                      priority=gasjobs.SEBASTIAN_PRIORITY, send_override=False)
    out = {
        "ok": True, "requestId": request_id, "workerName": done.req.worker_name,
        "targetDate": target_date_value(done.req, settings), "actualMinutes": done.actual,
        "breakMinutes": done.brk, "netMinutes": done.net, "status": done.status, "pdf": None,
    }
    db.commit()
    if done.status == "approved":
        out["pdf"] = _wait_pdf(db, done.job_id, settings)
    return out


def _holiday_start(db: Session, request_id: str, manual: str, now: datetime, settings: Settings) -> dict:
    req, start = worklog.core_holiday_start(db, request_id, manual, now, msgs=worklog.SEBASTIAN, email=None,
                                            updated_by=UPDATED_BY)
    out = {"ok": True, "requestId": request_id, "workerName": req.worker_name,
           "targetDate": target_date_value(req, settings), "actualStartAt": store.jst_text(start)}
    db.commit()
    return out


def _holiday_done(db: Session, request_id: str, manual: str, now: datetime, settings: Settings) -> dict:
    done = worklog.core_holiday_done(db, request_id, manual, now, msgs=worklog.SEBASTIAN, email=None,
                                     updated_by=UPDATED_BY, source="sebastian",
                                     priority=gasjobs.SEBASTIAN_PRIORITY, send_override=False)
    out = {
        "ok": True, "requestId": request_id, "workerName": done.req.worker_name,
        "targetDate": target_date_value(done.req, settings),
        "actualStartAt": store.jst_text(done.start), "actualEndAt": store.jst_text(done.end),
        "actualMinutes": done.actual, "breakMinutes": done.brk, "netMinutes": done.net,
        "status": done.status, "pdf": None,
    }
    db.commit()
    if done.status == "approved":
        out["pdf"] = _wait_pdf(db, done.job_id, settings)
    return out


def handle(db: Session, raw_body: bytes, settings: Settings | None = None) -> dict:
    s = settings or get_settings()
    try:
        text = (raw_body or b"").decode("utf-8") or "{}"
        body = json.loads(text)
    except (ValueError, UnicodeDecodeError) as exc:
        return {"error": "invalid JSON: " + str(exc)}

    expected = s.sebastian_api_token or ""
    if not expected:
        return {"error": "server misconfigured: SEBASTIAN_API_TOKEN not set"}
    token = body.get("token") if isinstance(body, dict) else None
    if not isinstance(token, str) or not hmac.compare_digest(token.encode("utf-8"), expected.encode("utf-8")):
        return {"error": "unauthorized"}

    action = js_string(body.get("action")) if js_truthy(body.get("action")) else ""
    now = timeutil.now_utc()
    try:
        if action == "ping":
            return {"ok": True, "pong": iso_utc_ms(now)}
        if action in ("completeOvertime", "markHolidayStart", "markHolidayDone"):
            if not js_truthy(body.get("requestId")):
                raise RpcError("requestId is required")
            request_id = js_string(body.get("requestId"))
            if action == "completeOvertime":
                return _complete_overtime(db, request_id, now, s)
            if action == "markHolidayStart":
                manual = js_string(body.get("manualStartTime")) if js_truthy(body.get("manualStartTime")) else ""
                return _holiday_start(db, request_id, manual, now, s)
            manual = js_string(body.get("manualEndTime")) if js_truthy(body.get("manualEndTime")) else ""
            return _holiday_done(db, request_id, manual, now, s)
        raise RpcError("unknown action: " + action)
    except RpcError as exc:
        db.rollback()
        return {"error": exc.message, "action": action}
    except Exception:  # noqa: BLE001 - GAS と同じく {error, action}。内部の例外の文は返さない（詳細はログ）
        db.rollback()
        log.exception("sebastian %s failed", action)
        return {"error": INTERNAL_ERROR, "action": action}
