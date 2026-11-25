"""申請の作成・編集・削除（4-3 の 7・11・12）。formSubmit.gs:678-837、worklog.gs:540-700 の写し。

- 作成（api_submitRequest）: 入力の誤りは {ok:false, error} を 200 で返す。作業員コードと予定時間の誤りは 400。
  Requests に 1 行、WorkLogs に下書きの行（updatedAt は日時のセル＝cell_kinds 'serial'、updatedBy='formSubmit'）。
  積むのは行の反映だけ。二重押しは rpc の冪等キーが止める（同じ人・同じ日の申請は止めない＝今と同じ）。
- 編集: PDF があれば 3 列を空にして internal.regenerate（reason=edit）。無ければ何も積まない（Q8）。
- 削除: canceled、PDF の 3 列を空、WorkLogs の 5 列を空。internal.afterDelete（PDF の有無に関係なく）。
"""
from __future__ import annotations

import math
import uuid
from typing import Any

from app import timeutil
from app.constants import HD_HOURS, OT_HOURS
from app.jsutil import js_round, js_string, normalize_gas
from app.models import Request, WorkLog
from app import sheet_queue
from app.rpc import RpcContext, RpcError

from . import gasjobs, store
from .fiscal import fy_start_0316
from .jsdate import YMD_RE, js_number, js_truthy, strict_date, valid_ymd_text
from .masters import lookup_order_info, lookup_worker_email


def _field(data: dict, key: str, default_empty: bool = False) -> str:
    value = data.get(key)
    if default_empty and not js_truthy(value):
        value = ""
    return normalize_gas(value)


MAX_PLANNED_MINUTES = 2**31 - 1
MAX_CELL_CHARS = 50000  # Google スプレッドシートの 1 セルの上限


def planned_minutes_overtime(h_str: Any) -> int:
    n = js_number(normalize_gas(h_str))
    if math.isnan(n) or math.isinf(n) or n <= 0 or not math.isfinite(n * 60) or n * 60 >= MAX_PLANNED_MINUTES:
        # 列（整数）に入らない大きさも不正にする（GAS は Infinity を書く。受け入れる差＝仕様 8-2。正しさ 低-10）
        raise RpcError(f"残業の予定時間が不正です: {js_string(h_str)}")
    return js_round(n * 60)


def planned_minutes_holiday(label: Any) -> int:
    s = normalize_gas(label)
    if s == "半日":
        return 240
    if s == "1日":
        return 480
    raise RpcError(f"休日の予定時間が不正です: {js_string(label)}")


def parse_worker_choice(label: str) -> tuple[str, str]:
    import re

    s = normalize_gas(label)
    if not s:
        return "", ""
    pieces = re.split(r"\s+", s)
    return pieces[0] or "", " ".join(pieces[1:]) or ""


def fiscal_check_overtime(target: str, today: str, fy_start: str) -> str:
    """isWithinCurrentFiscalYear_（誤りの文、合っていれば ''）。"""
    if not target or not YMD_RE.match(target):
        return f"作業実施日の形式が不正です（yyyy-MM-dd）: {target}"
    if not valid_ymd_text(target):
        return f"作業実施日が無効です: {target}"
    if target > today:
        return f"未来日は申請できません: {target}"
    if target < fy_start:
        return f"年度開始日（{fy_start}）より前の日付は申請できません。"
    return ""


def holiday_date_check(target: str, retro: bool, today: str, fy_start: str) -> str:
    """isHolidayDateValid_。"""
    if not target or not YMD_RE.match(target):
        return f"作業実施日の形式が不正です（yyyy-MM-dd）: {target}"
    if not valid_ymd_text(target):
        return f"作業実施日が無効です: {target}"
    if target < fy_start:
        return f"年度開始日（{fy_start}）より前の日付は申請できません。"
    if retro and target > today:
        return f"遡及申請では未来日を選択できません: {target}"
    return ""


def enrich_work_nos(db, w1: str, w2: str, w3: str) -> dict:
    result: dict[str, Any] = {"errors": []}
    for suffix, work_no in (("1", w1), ("2", w2), ("3", w3)):
        result["orderNo" + suffix] = work_no
        result["customer" + suffix] = ""
        result["product" + suffix] = ""
        if not work_no:
            continue
        info = lookup_order_info(db, work_no)
        if info:
            result["customer" + suffix] = info["dest"] or info["customer"] or ""
            result["product" + suffix] = info["product"] or ""
        else:
            result["errors"].append("工番マスタ未登録: " + work_no)
    return result


def api_submitRequest(ctx: RpcContext, data=None):
    if not isinstance(data, dict):
        data = {}
    request_type = _field(data, "requestType")
    dept = _field(data, "dept")
    worker_label = _field(data, "workerLabel")
    target_date = _field(data, "targetDate")
    work_no1 = _field(data, "workNo1", True)
    work_no2 = _field(data, "workNo2", True)
    work_no3 = _field(data, "workNo3", True)
    reason = _field(data, "reason", True)
    reason_detail = _field(data, "reasonDetail", True)
    work_content = _field(data, "workContent", True)
    hours = _field(data, "hours", True)
    retro_raw = data.get("isRetroactive")
    is_retro = retro_raw is True or (retro_raw is not None and js_string(retro_raw).lower() == "true")

    if request_type not in ("overtime", "holiday"):
        return {"ok": False, "error": "申請種別が不正です: " + request_type}
    if not dept:
        return {"ok": False, "error": "部署が指定されていません。"}
    if not worker_label:
        return {"ok": False, "error": "作業員が指定されていません。"}
    if not reason:
        return {"ok": False, "error": "理由を選択してください。"}
    if "その他" in reason and not reason_detail:
        return {"ok": False, "error": "理由が「その他」の場合、補足理由が必須です。"}
    if not hours:
        return {"ok": False, "error": "予定時間を選択してください。"}

    now = ctx.now
    today = timeutil.to_jst_text(now, "yyyy-MM-dd")
    fy_start = fy_start_0316(timeutil.to_jst(now).date()).isoformat()
    if request_type == "overtime" and not is_retro:
        target_date = today
    else:
        if not target_date:
            return {"ok": False, "error": "作業実施日を選択してください。"}
        if not YMD_RE.match(target_date):
            return {"ok": False, "error": "作業実施日の形式が不正です: " + target_date}
        if not valid_ymd_text(target_date):
            return {"ok": False, "error": "作業実施日が無効です: " + target_date}

    if request_type == "overtime":
        err = fiscal_check_overtime(target_date, today, fy_start)
    else:
        err = holiday_date_check(target_date, is_retro, today, fy_start)
    if err:
        return {"ok": False, "error": err}

    worker_code, worker_name = parse_worker_choice(worker_label)
    if not worker_code:
        raise RpcError("作業員コードが取得できません。")
    worker_email = lookup_worker_email(ctx.db, worker_code)

    approved_minutes = (planned_minutes_overtime(hours) if request_type == "overtime"
                        else planned_minutes_holiday(hours))
    enriched = enrich_work_nos(ctx.db, work_no1, work_no2, work_no3)

    # Google スプレッドシートのセルは 50,000 字まで。GAS も appendRow が例外で失敗するので、書く前に止める
    # （上限は GAS が失敗する所だけ。文字の中身は変えない。セキュリティ 低-2）
    for label, text in (("部署", dept), ("作業員", worker_name), ("作業員コード", worker_code),
                        ("業務NO.1", work_no1), ("業務NO.2", work_no2), ("業務NO.3", work_no3), ("理由", reason),
                        ("補足理由", reason_detail), ("業務内容", work_content)):
        if len(text) > MAX_CELL_CHARS:
            raise RpcError(f"{label}が長すぎます（{MAX_CELL_CHARS}字まで）。")

    request_id = str(uuid.uuid4())
    parsed = strict_date(target_date)
    req = Request(
        request_id=request_id, request_type=request_type, status="submitted", dept=dept,
        worker_code=worker_code, worker_name=worker_name, worker_email=worker_email,
        target_date=parsed, target_date_text="" if parsed is not None else target_date,
        submitted_at=now, approved_at=None, approved_by="", approved_minutes=approved_minutes,
        reason=reason, reason_detail=reason_detail, work_content=work_content, job_id1="",
        work_no1=work_no1, order_no1=enriched["orderNo1"], customer1=enriched["customer1"],
        product1=enriched["product1"],
        work_no2=work_no2, order_no2=enriched["orderNo2"], customer2=enriched["customer2"],
        product2=enriched["product2"],
        work_no3=work_no3, order_no3=enriched["orderNo3"], customer3=enriched["customer3"],
        product3=enriched["product3"],
        approved_by2="", approved_at2=None, hr_mail_sent_at="", pdf_generated_at=None, pdf_file_id="",
        pdf_folder_id="", export_error=" / ".join(enriched["errors"]) if enriched["errors"] else "",
    )
    ctx.db.add(req)
    ctx.db.flush()
    sheet_queue.mark_dirty(ctx.db, "requests", request_id)

    # ensureWorkLogRow_: requestId・updatedAt（日時のセル）・updatedBy='formSubmit' の下書きの行
    if store.first_worklog(ctx.db, request_id) is None:
        wl = WorkLog(request_id=request_id, updated_at=store.jst_text(now), updated_by="formSubmit",
                     cell_kinds={"updatedAt": "serial"})
        ctx.db.add(wl)
        ctx.db.flush()
        sheet_queue.mark_dirty(ctx.db, "work_logs", str(wl.id))

    return {"ok": True, "success": True, "requestId": request_id}


def api_editRequest(ctx: RpcContext, request_id=None, patch=None):
    if not isinstance(patch, dict):
        patch = {}
    hours = normalize_gas(patch.get("hours") if js_truthy(patch.get("hours")) else "")
    target_date = normalize_gas(patch.get("targetDate") if js_truthy(patch.get("targetDate")) else "")

    db = ctx.db
    req = store.lock_request(db, request_id)
    if req is None:
        raise RpcError("申請が見つかりません。")
    status = normalize_gas(req.status)
    if status == "canceled":
        raise RpcError("キャンセル済みの申請は編集できません。")
    if store.worklog_view(db, req.request_id).actualEndAt:
        raise RpcError("作業完了済みの申請は編集できません。")

    request_type = normalize_gas(req.request_type)
    if request_type == "overtime":
        if not hours:
            raise RpcError("予定時間を指定してください。")
        if hours not in OT_HOURS:
            raise RpcError("予定時間が不正です: " + hours)
        new_minutes = planned_minutes_overtime(hours)
    else:
        if not hours:
            raise RpcError("予定時間を指定してください。")
        if hours not in HD_HOURS:
            raise RpcError("予定時間が不正です: " + hours)
        new_minutes = planned_minutes_holiday(hours)
        if target_date:
            if not YMD_RE.match(target_date):
                raise RpcError("作業実施日の形式が不正です: " + target_date)
            if not valid_ymd_text(target_date):
                raise RpcError("作業実施日が無効です: " + target_date)

    req.approved_minutes = new_minutes
    new_target_date = ""
    if request_type == "holiday" and target_date:
        parsed = strict_date(target_date)
        req.target_date = parsed
        req.target_date_text = "" if parsed is not None else target_date
        new_target_date = target_date

    req.modified_at = ctx.now
    req.modified_by = ctx.user.email or "unknown"
    store.set_cell_kind(req, "modifiedAt", "string")

    had_pdf = js_truthy(req.pdf_file_id)
    old_pdf = store.clear_pdf_columns(req) if had_pdf else ""
    store.mark_request(db, req)
    if had_pdf:
        gasjobs.regenerate(db, req.request_id, reason="edit", old_pdf_file_id=old_pdf)

    out: dict[str, Any] = {"ok": True, "newStatus": req.status, "newMinutes": new_minutes}
    if new_target_date:
        out["newTargetDate"] = new_target_date
    return out


def api_deleteRequest(ctx: RpcContext, request_id=None):
    db = ctx.db
    req = store.lock_request(db, request_id)
    if req is None:
        raise RpcError("申請が見つかりません。")
    if normalize_gas(req.status) == "canceled":
        raise RpcError("既にキャンセル済みです。")

    request_type = normalize_gas(req.request_type)
    target = store.target_ymd(req) or store.target_raw_text(req)
    req.status = "canceled"
    old_pdf = store.clear_pdf_columns(req) if js_truthy(req.pdf_file_id) else ""
    store.mark_request(db, req)

    if store.first_worklog(db, req.request_id) is not None:
        store.update_worklog(db, req.request_id, {
            "actualStartAt": "", "actualEndAt": "", "actualMinutes": "", "breakMinutes": "", "netMinutes": "",
            "updatedAt": store.jst_text(ctx.now), "updatedBy": ctx.user.email or "delete",
        })

    gasjobs.after_delete(db, req.request_id, request_type=request_type, target_date=target,
                         old_pdf_file_id=old_pdf)
    return {"ok": True}


HANDLERS = {
    "api_submitRequest": api_submitRequest,
    "api_editRequest": api_editRequest,
    "api_deleteRequest": api_deleteRequest,
}
