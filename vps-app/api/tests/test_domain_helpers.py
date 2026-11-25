"""担当 W の試験の小道具（試験は入っていない）。申請・実績・マスタ・承認者を DB に置き、業務関数を直接呼ぶ。"""
from __future__ import annotations

import itertools
import json
from datetime import date, datetime

from sqlalchemy import select

from app import timeutil
from app.auth import CurrentUser
from app.models import (
    CompanyCalendar,
    DeptApprover,
    Job,
    Order,
    Request,
    Setting,
    SheetSyncState,
    Worker,
    WorkLog,
)
from app.rpc import RpcContext

ADMIN = "owner@lineworks-local.info"  # BOOTSTRAP_ADMINS
APPROVER = "boss@lineworks-local.info"
USER = "tanaka@lineworks-local.info"

_row = itertools.count(2)
_uid = itertools.count(1)


def jst(y, m, d, h=0, mi=0, s=0, ms=0) -> datetime:
    return datetime(y, m, d, h, mi, s, ms * 1000, tzinfo=timeutil.JST)


def ctx(db, email=USER, settings=None, now=None) -> RpcContext:
    from conftest import make_settings

    return RpcContext(db=db, user=CurrentUser(email=email, name=email.split("@")[0]),
                      settings=settings or make_settings(), now=(now or jst(2026, 11, 26, 10)).astimezone(timeutil.UTC))


def add_request(db, request_id=None, *, request_type="overtime", status="submitted", dept="製造1課",
                worker_code="A001", worker_name="山田太郎", worker_email="", target=date(2026, 11, 26),
                target_text="", submitted_at=None, approved_minutes=60, pdf_file_id="", approved_by2="",
                approved_at=None, approved_by="", sheet_row="auto", cell_kinds=None, reason="急ぎ作業を要するため",
                **extra) -> Request:
    rid = request_id or f"00000000-0000-4000-8000-{next(_uid):012d}"
    req = Request(request_id=rid, request_type=request_type, status=status, dept=dept, worker_code=worker_code,
                  worker_name=worker_name, worker_email=worker_email, target_date=target,
                  target_date_text=target_text, submitted_at=submitted_at or jst(2026, 11, 26, 7),
                  approved_minutes=approved_minutes, pdf_file_id=pdf_file_id, approved_by2=approved_by2,
                  approved_at=approved_at, approved_by=approved_by,
                  sheet_row=next(_row) if sheet_row == "auto" else sheet_row, cell_kinds=cell_kinds or {},
                  reason=reason, **extra)
    if pdf_file_id:
        req.pdf_folder_id = req.pdf_folder_id or "folder-1"
        req.pdf_generated_at = req.pdf_generated_at or jst(2026, 11, 26, 20)
    db.add(req)
    db.commit()
    return req


def add_wl(db, request_id, *, start="", end="", actual=None, brk=None, net=None, sheet_row="auto",
           updated_by="formSubmit") -> WorkLog:
    wl = WorkLog(request_id=request_id, actual_start_at=start, actual_end_at=end, actual_minutes=actual,
                 break_minutes=brk, net_minutes=net, updated_by=updated_by,
                 sheet_row=next(_row) if sheet_row == "auto" else sheet_row)
    db.add(wl)
    db.commit()
    return wl


def add_worker(db, sort_order, code, name, dept, active="", google=None):
    raw = {"作業員コード": code, "氏名": name, "部署": dept, "在籍フラグ": active}
    if google is not None:
        raw["Googleアカウント（任意）"] = google
    db.add(Worker(sort_order=sort_order, worker_code=code, worker_name=name, dept=dept, active_flag=active,
                  google_account=google or "", raw=raw))
    db.commit()


def add_order(db, sort_order, code, customer="", dest="", product=""):
    db.add(Order(sort_order=sort_order, code=code, customer=customer, dest=dest, product=product))
    db.commit()


def add_calendar(db, sort_order, date_text, kind="休日", weekday="", note=""):
    db.add(CompanyCalendar(sort_order=sort_order, date_text=date_text, kind=kind, weekday=weekday, note=note))
    db.commit()


def set_dept_approvers(db, rows):
    """rows: [(dept, 'a@x,b@x', name)]（DeptApprovers の 2 行目から）。"""
    for i, (dept, emails, name) in enumerate(rows, start=2):
        db.add(DeptApprover(sheet_row=i, dept=dept, emails=emails, name=name))
    db.commit()


def make_admin(db, *emails):
    row = db.get(Setting, "ADMIN_EMAILS")
    if row is None:
        db.add(Setting(key="ADMIN_EMAILS", value=",".join(emails), sheet_row=3))
    else:
        row.value = ",".join(emails)
    db.commit()


def jobs_of(db) -> list[dict]:
    rows = db.scalars(select(Job).order_by(Job.id)).all()
    return [{"kind": j.kind, "requestId": j.request_id, "args": json.loads(j.args_json), "priority": j.priority,
             "batchId": j.batch_id, "requires": json.loads(j.requires_json), "guard": json.loads(j.guard_json)}
            for j in rows]


def dirty(db) -> set[tuple[str, str]]:
    return {(s.table_name, s.pk) for s in db.scalars(select(SheetSyncState))
            if s.desired_version > s.synced_version}


def wls(db, request_id) -> list[WorkLog]:
    return list(db.scalars(select(WorkLog).where(WorkLog.request_id == request_id).order_by(WorkLog.id)))
