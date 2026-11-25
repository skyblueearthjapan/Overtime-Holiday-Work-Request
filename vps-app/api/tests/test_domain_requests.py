"""申請の作成・編集・削除（4-3 の 7・11・12。formSubmit.gs:678-837・worklog.gs:540-700）。"""
from __future__ import annotations

from datetime import date

import pytest
from sqlalchemy import select

from app.models import Request
from app.rpc import RpcError
from app.domain import requests as rq
from test_domain_helpers import (
    USER,
    add_order,
    add_request,
    add_wl,
    add_worker,
    ctx,
    dirty,
    jobs_of,
    jst,
    wls,
)

NOW = jst(2026, 11, 26, 10)


def base(**kw):
    data = {"requestType": "overtime", "dept": "製造1課", "workerLabel": "A001 山田 太郎", "targetDate": "",
            "reason": "急ぎ作業を要するため", "hours": "1.5"}
    data.update(kw)
    return data


def submit(db, data, now=NOW):
    return rq.api_submitRequest(ctx(db, now=now), data)


# ===== 作成 =====


def test_submit_overtime_happy_path_writes_request_and_draft_worklog(db):
    add_worker(db, 1, "A001", "山田 太郎", "製造1課", google="Yamada@lineworks-local.info")
    add_order(db, 1, "LW001", customer="受注先A", dest="納入先A", product="品名A")
    add_order(db, 2, "LW002", customer="受注先B", dest="", product="品名B")
    res = submit(db, base(workNo1="ＬＷ００１", workNo2="LW002", workNo3="LW999", reasonDetail=" 補足 ",
                          workContent="作業", targetDate="2020-01-01"))
    assert res["ok"] is True and res["success"] is True and len(res["requestId"]) == 36
    req = db.get(Request, res["requestId"])
    assert req.status == "submitted" and req.request_type == "overtime"
    assert req.target_date == date(2026, 11, 26)  # 通常の残業は送られた日付を捨てて今日（JST）
    assert req.worker_code == "A001" and req.worker_name == "山田 太郎"
    assert req.worker_email == "Yamada@lineworks-local.info"
    assert req.approved_minutes == 90
    assert (req.work_no1, req.order_no1, req.customer1, req.product1) == ("LW001", "LW001", "納入先A", "品名A")
    assert (req.customer2, req.product2) == ("受注先B", "品名B")  # 納入先が空なら受注先
    assert req.export_error == "工番マスタ未登録: LW999"
    assert req.reason_detail == "補足" and req.submitted_at is not None
    wl = wls(db, req.request_id)
    assert len(wl) == 1 and wl[0].updated_by == "formSubmit" and wl[0].updated_at == "2026-11-26 10:00:00"
    assert wl[0].cell_kinds == {"updatedAt": "serial"} and wl[0].actual_start_at == ""
    assert dirty(db) == {("requests", req.request_id), ("work_logs", str(wl[0].id))}
    assert jobs_of(db) == []  # 申請は行の反映だけ


def test_submit_export_error_joins_with_slash(db):
    res = submit(db, base(workNo1="X1", workNo2="X2"))
    assert db.get(Request, res["requestId"]).export_error == "工番マスタ未登録: X1 / 工番マスタ未登録: X2"


@pytest.mark.parametrize("data,error", [
    (base(requestType="other"), "申請種別が不正です: other"),
    (base(requestType=None), "申請種別が不正です: "),
    (base(dept=" "), "部署が指定されていません。"),
    (base(workerLabel=""), "作業員が指定されていません。"),
    (base(reason=""), "理由を選択してください。"),
    (base(reason="その他（→補足理由必須）"), "理由が「その他」の場合、補足理由が必須です。"),
    (base(reason="これはその他です"), "理由が「その他」の場合、補足理由が必須です。"),
    (base(hours=""), "予定時間を選択してください。"),
    (base(isRetroactive=True, targetDate=""), "作業実施日を選択してください。"),
    (base(isRetroactive="TRUE", targetDate="2026/11/20"), "作業実施日の形式が不正です: 2026/11/20"),
    (base(isRetroactive="true", targetDate="2026-13-01"), "作業実施日が無効です: 2026-13-01"),
    (base(isRetroactive=True, targetDate="2026-11-27"), "未来日は申請できません: 2026-11-27"),
    (base(isRetroactive=True, targetDate="2026-03-15"), "年度開始日（2026-03-16）より前の日付は申請できません。"),
    (base(requestType="holiday", hours="1日", targetDate="2026-03-15"),
     "年度開始日（2026-03-16）より前の日付は申請できません。"),
    (base(requestType="holiday", hours="1日", targetDate="2026-11-27", isRetroactive=True),
     "遡及申請では未来日を選択できません: 2026-11-27"),
    (base(requestType="holiday", hours="1日", targetDate=""), "作業実施日を選択してください。"),
])
def test_submit_input_errors_are_returned_not_raised(db, data, error):
    assert submit(db, data) == {"ok": False, "error": error}
    assert db.scalars(select(Request)).all() == []


@pytest.mark.parametrize("data,error", [
    (base(hours="0"), "残業の予定時間が不正です: 0"),
    (base(hours="abc"), "残業の予定時間が不正です: abc"),
    (base(requestType="holiday", hours="2日", targetDate="2026-11-28"), "休日の予定時間が不正です: 2日"),
])
def test_submit_planned_time_errors_raise(db, data, error):
    with pytest.raises(RpcError) as e:
        submit(db, data)
    assert e.value.message == error


def test_submit_does_not_check_hours_against_list_q13(db):
    res = submit(db, base(hours="100"))
    assert db.get(Request, res["requestId"]).approved_minutes == 6000
    res = submit(db, base(hours="0.25"))
    assert db.get(Request, res["requestId"]).approved_minutes == 15


def test_submit_holiday_future_without_upper_limit_q26(db):
    res = submit(db, base(requestType="holiday", hours="半日", targetDate="2027-12-31"))
    req = db.get(Request, res["requestId"])
    assert req.target_date == date(2027, 12, 31) and req.approved_minutes == 240


def test_submit_feb_30_is_accepted_like_v8_and_kept_as_text(db):
    res = submit(db, base(requestType="holiday", hours="1日", targetDate="2027-02-30"))
    req = db.get(Request, res["requestId"])
    assert req.target_date is None and req.target_date_text == "2027-02-30"


def test_submit_same_person_same_day_twice_is_allowed(db):
    assert submit(db, base())["ok"] and submit(db, base())["ok"]
    assert len(db.scalars(select(Request)).all()) == 2


def test_submit_without_google_account_column_has_empty_email(db):
    add_worker(db, 1, "A001", "山田", "製造1課")
    res = submit(db, base())
    assert db.get(Request, res["requestId"]).worker_email == ""


# ===== 編集 =====


def edit(db, rid, patch, email=USER):
    return rq.api_editRequest(ctx(db, email=email, now=NOW), rid, patch)


def test_edit_overtime_hours(db):
    req = add_request(db, approved_minutes=60)
    res = edit(db, req.request_id, {"hours": "2.0", "targetDate": "2026-12-01"})
    assert res == {"ok": True, "newStatus": "submitted", "newMinutes": 120}  # newTargetDate は出さない
    db.refresh(req)
    assert req.approved_minutes == 120 and req.target_date == date(2026, 11, 26)
    assert req.modified_by == USER and req.modified_at is not None
    assert req.cell_kinds.get("modifiedAt") == "string"
    assert ("requests", req.request_id) in dirty(db)
    assert jobs_of(db) == []  # PDF が無ければ何も積まない（Q8）


def test_edit_holiday_date_and_pdf_regenerate(db):
    req = add_request(db, request_type="holiday", status="approved", pdf_file_id="pdf-old")
    res = edit(db, req.request_id, {"hours": "1日", "targetDate": "2026-03-01"})  # 年度・未来は見ない（Q7）
    assert res == {"ok": True, "newStatus": "approved", "newMinutes": 480, "newTargetDate": "2026-03-01"}
    db.refresh(req)
    assert req.target_date == date(2026, 3, 1)
    assert (req.pdf_file_id, req.pdf_folder_id, req.pdf_generated_at) == ("", "", None)
    assert req.old_pdf_file_id_pending == "pdf-old"
    jobs = jobs_of(db)
    assert [(j["kind"], j["args"]) for j in jobs] == [
        ("internal.regenerate", {"reason": "edit", "oldPdfFileId": "pdf-old"})]
    assert jobs[0]["requires"][0]["table"] == "requests" and jobs[0]["requires"][0]["version"] == req.row_version


@pytest.mark.parametrize("setup,patch,error", [
    ("missing", {"hours": "1.0"}, "申請が見つかりません。"),
    ("canceled", {"hours": "1.0"}, "キャンセル済みの申請は編集できません。"),
    ("done", {"hours": "1.0"}, "作業完了済みの申請は編集できません。"),
    ("ot", {}, "予定時間を指定してください。"),
    ("ot", {"hours": "1"}, "予定時間が不正です: 1"),
    ("ot", {"hours": "4.5"}, "予定時間が不正です: 4.5"),
    ("hol", {"hours": "1.0"}, "予定時間が不正です: 1.0"),
    ("hol", {"hours": ""}, "予定時間を指定してください。"),
    ("hol", {"hours": "半日", "targetDate": "2026/12/01"}, "作業実施日の形式が不正です: 2026/12/01"),
    ("hol", {"hours": "半日", "targetDate": "2026-12-32"}, "作業実施日が無効です: 2026-12-32"),
])
def test_edit_errors(db, setup, patch, error):
    rid = "missing-id"
    if setup == "canceled":
        rid = add_request(db, status="canceled").request_id
    elif setup == "done":
        rid = add_request(db).request_id
        add_wl(db, rid, start="2026-11-26 17:20:00", end="2026-11-26 19:00:00", actual=100, brk=0, net=100)
    elif setup == "ot":
        rid = add_request(db).request_id
    elif setup == "hol":
        rid = add_request(db, request_type="holiday").request_id
    with pytest.raises(RpcError) as e:
        edit(db, rid, patch)
    assert e.value.message == error


# ===== 削除 =====


def test_delete_clears_pdf_and_worklog_and_enqueues_after_delete(db):
    req = add_request(db, request_type="holiday", status="approved", pdf_file_id="pdf-1", approved_by2=USER)
    wl = add_wl(db, req.request_id, start="2026-11-26 08:00:00", end="2026-11-26 17:00:00", actual=540, brk=55,
                net=485)
    assert rq.api_deleteRequest(ctx(db, now=NOW), req.request_id) == {"ok": True}
    db.refresh(req)
    db.refresh(wl)
    assert req.status == "canceled" and req.pdf_file_id == "" and req.old_pdf_file_id_pending == "pdf-1"
    assert (wl.actual_start_at, wl.actual_end_at, wl.actual_minutes, wl.break_minutes, wl.net_minutes) == (
        "", "", None, None, None)
    assert wl.updated_by == USER and wl.updated_at == "2026-11-26 10:00:00"
    assert wl.cell_kinds["updatedAt"] == "string"
    assert {("requests", req.request_id), ("work_logs", str(wl.id))} <= dirty(db)
    assert [(j["kind"], j["args"]) for j in jobs_of(db)] == [("internal.afterDelete", {
        "oldPdfFileId": "pdf-1", "requestType": "holiday", "targetDate": "2026-11-26"})]


def test_delete_without_pdf_or_worklog_still_enqueues(db):
    req = add_request(db)
    rq.api_deleteRequest(ctx(db), req.request_id)
    assert wls(db, req.request_id) == []  # WorkLogs の行は作らない
    assert [(j["kind"], j["args"]["oldPdfFileId"]) for j in jobs_of(db)] == [("internal.afterDelete", "")]


def test_delete_twice(db):
    req = add_request(db)
    rq.api_deleteRequest(ctx(db), req.request_id)
    with pytest.raises(RpcError) as e:
        rq.api_deleteRequest(ctx(db), req.request_id)
    assert e.value.message == "既にキャンセル済みです。"
    with pytest.raises(RpcError) as e:
        rq.api_deleteRequest(ctx(db), "nope")
    assert e.value.message == "申請が見つかりません。"
