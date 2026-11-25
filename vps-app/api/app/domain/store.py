"""業務の表の読み書きの小道具（GAS の getRequestById_・findWorkLogRowNo_・buildWorkLogsMapByRequestId_・
buildJoinedRecords_ の写し）。

- シートの並び（GAS が上から読む順）は sheet_row の順。まだシートに書いていない新しい行（sheet_row が NULL）は末尾。
- 行ロック（実装仕様 5-5）は lock_request の 1 か所。PostgreSQL は SELECT … FOR UPDATE、SQLite は何もしない。
  一括の処理は ID を並べ替えてからロックする（lock_requests）。
- WorkLogs は requestId が重なることがある。書くときは上の行（findWorkLogRowNo_）、集計・表示は下の行が勝つ
  （buildWorkLogsMapByRequestId_ は上から読んで Map.set するため）。
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from typing import Any

from sqlalchemy import func, or_, select
from sqlalchemy.orm import Session

from app import sheet_queue, timeutil
from app.jsutil import normalize_gas
from app.models import Request, WorkLog

from .jsdate import js_new_date

DT_FMT = "yyyy-MM-dd HH:mm:ss"


def jst_text(moment: datetime | None, pattern: str = DT_FMT) -> str:
    return timeutil.to_jst_text(moment, pattern) if moment is not None else ""


# ===== Requests =====


def _request_order():
    return (Request.sheet_row.is_(None), Request.sheet_row, Request.submitted_at, Request.request_id)


def requests_in_sheet_order(db: Session) -> list[Request]:
    return list(db.scalars(select(Request).order_by(*_request_order())))


def request_count(db: Session) -> int:
    return int(db.scalar(select(func.count()).select_from(Request)) or 0)


def _like_any(text: str) -> str:
    r"""LIKE '%text%' の形（% _ \ を \ で逃がす。escape="\\" と一緒に使う）。"""
    return "%" + text.replace("\\", "\\\\").replace("%", r"\%").replace("_", r"\_") + "%"


def _gas_id_arg(request_id: Any) -> str | None:
    """GAS は normalize_(行の ID) === 引数 で探す。normalize_ の結果と同じ形の引数しか当たらない。"""
    if not isinstance(request_id, str) or not request_id or normalize_gas(request_id) != request_id:
        return None
    return request_id


def find_request_id(db: Session, request_id: Any) -> str | None:
    """引数に当たる申請の DB の ID（GAS と同じく行の ID の前後の空白を無視する。無ければ None）。

    同じ ID の行があればそれ。無ければ、前後に空白などが付いた行（normalize_ すると同じ）をシートの順で探す。
    """
    rid = _gas_id_arg(request_id)
    if rid is None:
        return None
    if db.scalar(select(Request.request_id).where(Request.request_id == rid)) is not None:
        return rid
    rows = db.scalars(select(Request).where(Request.request_id.like(_like_any(rid), escape="\\"))
                      .order_by(*_request_order())).all()
    for row in rows:
        if normalize_gas(row.request_id) == rid:
            return row.request_id
    return None


def lock_request(db: Session, request_id: Any) -> Request | None:
    """申請の行を FOR UPDATE で読む（無ければ None）。GAS は normalize_(行の ID) === 引数 で探す。"""
    db_id = find_request_id(db, request_id)
    if db_id is None:
        return None
    stmt = select(Request).where(Request.request_id == db_id).with_for_update()
    return db.execute(stmt.execution_options(populate_existing=True)).scalar_one_or_none()


def lock_requests(db: Session, request_ids: list[str]) -> dict[str, Request]:
    """一括の処理: ID を並べ替えてから 1 件ずつロックする（デッドロックを防ぐ）。"""
    out: dict[str, Request] = {}
    for rid in sorted({r for r in request_ids if isinstance(r, str) and r}):
        row = lock_request(db, rid)
        if row is not None:
            out[rid] = row
    return out


def target_ymd(req: Request) -> str | None:
    """fmtDate_(new Date(targetDate)) の値（読めなければ None）。"""
    if req.target_date is not None:
        return req.target_date.isoformat()
    text = req.target_date_text or ""
    if not text:
        return None
    moment = js_new_date(text)
    return timeutil.to_jst_text(moment, "yyyy-MM-dd") if moment is not None else None


def target_date_of(req: Request) -> date | None:
    ymd = target_ymd(req)
    return date.fromisoformat(ymd) if ymd else None


def target_raw_text(req: Request) -> str:
    """シートの targetDate が文字のときの値（DB の日付から戻す）。"""
    if req.target_date_text:
        return req.target_date_text
    return req.target_date.isoformat() if req.target_date is not None else ""


def set_cell_kind(row: Any, column: str, kind: str) -> None:
    kinds = dict(row.cell_kinds or {})
    kinds[column] = kind
    row.cell_kinds = kinds


def clear_pdf_columns(req: Request) -> str:
    """clearExistingPdfForRegeneration_ の DB 側: 3 列を空にして、前の ID を old_pdf_file_id_pending に残す。"""
    old = normalize_gas(req.pdf_file_id)
    req.pdf_generated_at = None
    req.pdf_file_id = ""
    req.pdf_folder_id = ""
    if old:
        req.old_pdf_file_id_pending = old
    return old


def mark_request(db: Session, req: Request) -> int:
    return sheet_queue.mark_dirty(db, "requests", req.request_id)


# ===== WorkLogs =====


def _wl_order():
    return (WorkLog.sheet_row.is_(None), WorkLog.sheet_row, WorkLog.id)


def worklogs_for(db: Session, request_id: Any) -> list[WorkLog]:
    """normalize_(行の requestId) === 引数 の WorkLogs の行（シートの順）。前後の空白は GAS と同じく無視する。"""
    rid = normalize_gas(request_id)
    if not rid:
        return []
    stmt = (select(WorkLog)
            .where(or_(WorkLog.request_id == rid, WorkLog.request_id.like(_like_any(rid), escape="\\")))
            .order_by(*_wl_order()))
    return [wl for wl in db.scalars(stmt) if normalize_gas(wl.request_id) == rid]


def first_worklog(db: Session, request_id: str) -> WorkLog | None:
    """findWorkLogRowNo_: 上から探して最初の行。"""
    rows = worklogs_for(db, request_id)
    return rows[0] if rows else None


def worklogs_in_sheet_order(db: Session) -> list[WorkLog]:
    return list(db.scalars(select(WorkLog).order_by(*_wl_order())))


@dataclass
class WlView:
    """buildWorkLogsMapByRequestId_ の 1 件（文字と数）。"""

    actualStartAt: str = ""
    actualEndAt: str = ""
    actualMinutes: int = 0
    breakMinutes: int = 0
    netMinutes: int = 0


EMPTY_WL = WlView()


def worklog_map(db: Session) -> dict[str, WlView]:
    out: dict[str, WlView] = {}
    for wl in worklogs_in_sheet_order(db):
        rid = normalize_gas(wl.request_id)
        if not rid:
            continue
        out[rid] = WlView(
            actualStartAt=wl.actual_start_at or "",
            actualEndAt=wl.actual_end_at or "",
            actualMinutes=int(wl.actual_minutes or 0),
            breakMinutes=int(wl.break_minutes or 0),
            netMinutes=int(wl.net_minutes or 0),
        )
    return out


def worklog_view(db: Session, request_id: str) -> WlView:
    """wlMap.get(requestId) || {}（下の行が勝つ）。"""
    rows = worklogs_for(db, request_id)
    if not rows:
        return EMPTY_WL
    wl = rows[-1]
    return WlView(wl.actual_start_at or "", wl.actual_end_at or "", int(wl.actual_minutes or 0),
                  int(wl.break_minutes or 0), int(wl.net_minutes or 0))


def update_worklog(db: Session, request_id: str, patch: dict[str, Any]) -> WorkLog:
    """updateWorkLog_: 上の行を書き換える（無ければ requestId だけの行を作る）。

    patch のキーは GAS の見出し（actualStartAt・actualEndAt・actualMinutes・breakMinutes・netMinutes・updatedAt・
    updatedBy）。日時の 3 列は文字（表示形式 @）で書くので、updatedAt の型は 'string'。分の '' は NULL。
    """
    wl = first_worklog(db, request_id)
    if wl is None:
        wl = WorkLog(request_id=request_id)
        db.add(wl)
    columns = {
        "actualStartAt": "actual_start_at", "actualEndAt": "actual_end_at", "actualMinutes": "actual_minutes",
        "breakMinutes": "break_minutes", "netMinutes": "net_minutes", "updatedAt": "updated_at",
        "updatedBy": "updated_by",
    }
    for key, value in patch.items():
        attr = columns[key]
        if attr.endswith("_minutes"):
            value = None if value == "" or value is None else int(value)
        setattr(wl, attr, value)
    if "updatedAt" in patch:
        set_cell_kind(wl, "updatedAt", "string")
    db.flush()
    sheet_queue.mark_dirty(db, "work_logs", str(wl.id))
    return wl


def parse_wl_datetime(raw: str) -> datetime | None:
    """new Date(String(raw).replace(' ', 'T') + '+09:00')（WorkLogs の JST の文字）。"""
    if not raw:
        return None
    text = str(raw).replace(" ", "T", 1) + "+09:00"
    return js_new_date(text)


# ===== buildJoinedRecords_ =====


@dataclass
class Joined:
    requestId: str
    status: str
    requestType: str
    dept: str
    workerCode: str
    workerName: str
    targetDate: date
    approvedMinutes: int
    actualStartAt: str
    actualEndAt: str
    actualMinutes: int
    breakMinutes: int
    netMinutes: int
    pdfFileId: str
    pdfGeneratedAt: datetime | None
    submittedAt: datetime | None


def joined_records(db: Session, date_from: date, date_to: date) -> list[Joined]:
    """status が空・canceled を除き、対象日が [from, to) の申請（submitted も含む）に WorkLogs を結ぶ。"""
    wl_map = worklog_map(db)
    out: list[Joined] = []
    for req in requests_in_sheet_order(db):
        status = normalize_gas(req.status)
        if not status or status == "canceled":
            continue
        td = target_date_of(req)
        if td is None or td < date_from or td >= date_to:
            continue
        rid = normalize_gas(req.request_id)
        wl = wl_map.get(rid, EMPTY_WL)
        out.append(Joined(
            requestId=rid,
            status=status,
            requestType=normalize_gas(req.request_type),
            dept=normalize_gas(req.dept),
            workerCode=normalize_gas(req.worker_code),
            workerName=normalize_gas(req.worker_name),
            targetDate=td,
            approvedMinutes=int(req.approved_minutes or 0),
            actualStartAt=wl.actualStartAt or "",
            actualEndAt=wl.actualEndAt or "",
            actualMinutes=wl.actualMinutes,
            breakMinutes=wl.breakMinutes,
            netMinutes=wl.netMinutes,
            pdfFileId=normalize_gas(req.pdf_file_id),
            pdfGeneratedAt=req.pdf_generated_at,
            submittedAt=req.submitted_at,
        ))
    return out
