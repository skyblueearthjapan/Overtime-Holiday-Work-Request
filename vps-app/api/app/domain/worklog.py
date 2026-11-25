"""作業実績（4-3 の 8・9・10・26）。worklog.gs:143-342・adminApi.gs:432-518 の写し。

- 残業の開始は対象日 17:20 JST 固定。終了は手入力 'HH:mm'（合わなければ黙って今の時刻＝gap-1 G13）か今。
- 休日の開始は手入力か今。完了は WorkLogs の開始から。終了が開始より前でも 0 分。
- 実働 = max(0, js_round(ミリ秒差 / 60000))、休憩 = 時刻窓方式、net = max(0, 実働 - 休憩)。
- 書く前に申請の行を FOR UPDATE でロックする（gap-1 G3。承認と完了が同時でも、どちらかが必ず PDF の仕事を積む）。
- 完了（画面・総務・Sebastian）は internal.afterComplete を積む（status に関係なく。台帳はいつも書く）。
  開始は何も積まない。
- 画面の関数と Sebastian は同じ計算（core_*）を使い、エラー文だけが違う（Messages）。
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Any

from sqlalchemy.orm import Session

from app.constants import OVERTIME_START_HH_MM
from app.jsutil import js_round, js_string, normalize_gas
from app.models import Request
from app.rpc import RpcContext, RpcError

from . import gasjobs, store
from .breaks import calc_break_by_clock_time
from .jsdate import HM_RE, js_truthy, jst_at, truncate_ms

DT = "yyyy-MM-dd HH:mm:ss"


@dataclass(frozen=True)
class Messages:
    not_found: str
    not_overtime: str
    not_holiday: str
    canceled: str
    no_start: str
    with_id: bool = False

    def fmt(self, text: str, request_id: str) -> str:
        return text + request_id if self.with_id else text


SCREEN = Messages(
    not_found="申請が見つかりません。",
    not_overtime="残業申請ではありません。",
    not_holiday="休日申請ではありません。",
    canceled="キャンセル済みです。",
    no_start="休日開始が記録されていません（開始ボタンを押してください）。",
)
SEBASTIAN = Messages(
    not_found="申請が見つかりません: ",
    not_overtime="残業申請ではありません: ",
    not_holiday="休日出勤申請ではありません: ",
    canceled="キャンセル済みの申請です: ",
    no_start="休日開始が記録されていません（先にstartを実行してください）: ",
    with_id=True,
)


def _target_ymd_or_raise(req: Request) -> str:
    ymd = store.target_ymd(req)
    if ymd is None:
        raise RpcError("作業実施日が無効です: " + store.target_raw_text(req))
    return ymd


def _manual_time(ymd_getter, manual: Any, now: datetime) -> datetime:
    """手入力 'HH:mm' があれば対象日のその時刻、無ければ今。時・分が範囲外なら 400（受け入れる差。8-2）。"""
    if js_truthy(manual) and HM_RE.match(js_string(manual)):
        text = js_string(manual)
        moment = jst_at(ymd_getter(), text)
        if moment is None:
            raise RpcError("時刻が不正です: " + text)
        return moment
    return now


def compute_minutes(start: datetime, end: datetime) -> tuple[int, int, int]:
    delta_ms = (truncate_ms(end) - truncate_ms(start)) // timedelta(milliseconds=1)
    actual = max(0, js_round(delta_ms / 60000))
    brk = calc_break_by_clock_time(start, end)
    return actual, brk, max(0, actual - brk)


def _load(db: Session, request_id: Any, msgs: Messages) -> Request:
    req = store.lock_request(db, request_id)
    if req is None:
        raise RpcError(msgs.fmt(msgs.not_found, js_string(request_id)))
    return req


def _check(req: Request, expected_type: str, msgs: Messages) -> None:
    rid = req.request_id
    if normalize_gas(req.request_type) != expected_type:
        raise RpcError(msgs.fmt(msgs.not_overtime if expected_type == "overtime" else msgs.not_holiday, rid))
    if normalize_gas(req.status) == "canceled":
        raise RpcError(msgs.fmt(msgs.canceled, rid))


def assert_self(req: Request, email: str) -> None:
    """assertSelf_: workerEmail が空でなく、ログインのメールと大文字小文字を区別して違えば拒む（Q24 の癖のまま）。"""
    worker_email = normalize_gas(req.worker_email)
    if worker_email and worker_email != email:
        raise RpcError("本人のみ操作可能です。")


@dataclass
class Completed:
    req: Request
    status: str
    start: datetime
    end: datetime
    actual: int
    brk: int
    net: int
    job_id: int

    @property
    def wl_override(self) -> dict:
        return {
            "actualStartAt": store.jst_text(self.start), "actualEndAt": store.jst_text(self.end),
            "actualMinutes": self.actual, "breakMinutes": self.brk, "netMinutes": self.net,
        }


def _write_complete(db: Session, req: Request, start: datetime, end: datetime, now: datetime, updated_by: str,
                    source: str, priority: int = 0, send_override: bool = True) -> Completed:
    actual, brk, net = compute_minutes(start, end)
    store.update_worklog(db, req.request_id, {
        "actualStartAt": store.jst_text(start), "actualEndAt": store.jst_text(end), "actualMinutes": actual,
        "breakMinutes": brk, "netMinutes": net, "updatedAt": store.jst_text(now), "updatedBy": updated_by,
    })
    status = normalize_gas(req.status)
    done = Completed(req, status, start, end, actual, brk, net, 0)
    done.job_id = gasjobs.after_complete(db, req.request_id, source=source, status_at_enqueue=status,
                                         wl_override=done.wl_override if send_override else None,
                                         priority=priority)
    return done


# ===== 共通の計算（画面と Sebastian） =====


def core_overtime_done(db: Session, request_id: Any, manual_end: Any, now: datetime, *, msgs: Messages,
                       email: str | None, updated_by: str, source: str, priority: int = 0,
                       send_override: bool = True) -> Completed:
    req = _load(db, request_id, msgs)
    _check(req, "overtime", msgs)
    if email is not None:
        assert_self(req, email)
    ymd = _target_ymd_or_raise(req)
    start = jst_at(ymd, "%02d:%02d" % OVERTIME_START_HH_MM)
    assert start is not None
    end = _manual_time(lambda: ymd, manual_end, now)
    return _write_complete(db, req, start, end, now, updated_by, source, priority, send_override)


def core_holiday_start(db: Session, request_id: Any, manual_start: Any, now: datetime, *, msgs: Messages,
                       email: str | None, updated_by: str) -> tuple[Request, datetime]:
    req = _load(db, request_id, msgs)
    _check(req, "holiday", msgs)
    if email is not None:
        assert_self(req, email)
    start = _manual_time(lambda: _target_ymd_or_raise(req), manual_start, now)
    store.update_worklog(db, req.request_id, {
        "actualStartAt": store.jst_text(start), "updatedAt": store.jst_text(now), "updatedBy": updated_by,
    })
    return req, start


def core_holiday_done(db: Session, request_id: Any, manual_end: Any, now: datetime, *, msgs: Messages,
                      email: str | None, updated_by: str, source: str, priority: int = 0,
                      send_override: bool = True) -> Completed:
    req = _load(db, request_id, msgs)
    _check(req, "holiday", msgs)
    if email is not None:
        assert_self(req, email)
    wl = store.first_worklog(db, req.request_id)
    if wl is None:
        raise RpcError(msgs.fmt(msgs.no_start, req.request_id))
    start = store.parse_wl_datetime(wl.actual_start_at or "")
    if start is None:
        raise RpcError(msgs.fmt(msgs.no_start, req.request_id))
    end = _manual_time(lambda: _target_ymd_or_raise(req), manual_end, now)
    return _write_complete(db, req, start, end, now, updated_by, source, priority, send_override)


# ===== 画面の関数 =====


def api_markOvertimeDone(ctx: RpcContext, request_id=None, manual_end_time=None):
    done = core_overtime_done(ctx.db, request_id, manual_end_time, ctx.now, msgs=SCREEN, email=ctx.user.email,
                              updated_by=ctx.user.email or "unknown", source="screen")
    return {"ok": True, "requestId": done.req.request_id, "actualMinutes": done.actual,
            "breakMinutes": done.brk, "netMinutes": done.net, "pdf": None}


def api_markHolidayStart(ctx: RpcContext, request_id=None, manual_start_time=None):
    req, start = core_holiday_start(ctx.db, request_id, manual_start_time, ctx.now, msgs=SCREEN,
                                    email=ctx.user.email, updated_by=ctx.user.email or "unknown")
    return {"ok": True, "requestId": req.request_id, "actualStartAt": store.jst_text(start)}


def api_markHolidayDone(ctx: RpcContext, request_id=None, manual_end_time=None):
    done = core_holiday_done(ctx.db, request_id, manual_end_time, ctx.now, msgs=SCREEN, email=ctx.user.email,
                             updated_by=ctx.user.email or "unknown", source="screen")
    return {"ok": True, "requestId": done.req.request_id, "actualMinutes": done.actual,
            "breakMinutes": done.brk, "netMinutes": done.net, "pdf": None}


_HM_ANY = re.compile(r"([0-9]{2}:[0-9]{2})")


def api_adminManualComplete(ctx: RpcContext, request_id=None, end_time_str=None, start_time_str=None):
    db = ctx.db
    req = store.lock_request(db, request_id)
    if req is None:
        raise RpcError("申請が見つかりません。")
    if normalize_gas(req.status) == "canceled":
        raise RpcError("キャンセル済みの申請には入力できません。")
    ymd = _target_ymd_or_raise(req)

    if js_truthy(start_time_str):
        start_str = js_string(start_time_str)
    elif normalize_gas(req.request_type) == "overtime":
        start_str = "17:20"
    else:
        start_str = ""
        wl = store.worklog_view(db, req.request_id)
        if wl.actualStartAt:
            m = _HM_ANY.search(str(wl.actualStartAt))
            start_str = m.group(1) if m else ""
        if not start_str:
            raise RpcError("休日出勤の開始時刻が不明です。開始時刻を指定してください。")

    if not js_truthy(end_time_str):
        raise RpcError("終了時刻を指定してください。")
    end_str = js_string(end_time_str)

    start = jst_at(ymd, start_str)
    end = jst_at(ymd, end_str)
    if start is None:
        raise RpcError("開始時刻が不正です: " + start_str)
    if end is None:
        raise RpcError("終了時刻が不正です: " + end_str)
    if end <= start:
        raise RpcError("終了時刻は開始時刻より後にしてください。")

    done = _write_complete(db, req, start, end, ctx.now, ctx.user.email or "admin-manual", "admin")
    return {"ok": True, "actualMinutes": done.actual, "breakMinutes": done.brk, "netMinutes": done.net}


HANDLERS = {
    "api_markOvertimeDone": api_markOvertimeDone,
    "api_markHolidayStart": api_markHolidayStart,
    "api_markHolidayDone": api_markHolidayDone,
    "api_adminManualComplete": api_adminManualComplete,
}
