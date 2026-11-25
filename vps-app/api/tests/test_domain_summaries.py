"""集計（4-3 の 18・22〜25・27〜29。adminApi.gs:36-376・staffSummaryApi.gs）。数は固定の申請から手で求めた値。"""
from __future__ import annotations

from datetime import date

import pytest

from app.domain import summaries
from app.rpc import RpcError
from test_domain_helpers import ADMIN, add_request, add_wl, add_worker, ctx, jst

NOW = jst(2026, 11, 26, 10)


def rec(db, dept, name, day, net, *, status="approved", rtype="overtime", code="", pdf="", minutes=60,
        submitted=None, brk=0):
    r = add_request(db, dept=dept, worker_name=name, worker_code=code, target=day, status=status,
                    request_type=rtype, pdf_file_id=pdf, approved_minutes=minutes,
                    submitted_at=submitted or jst(day.year, day.month, day.day, 7))
    add_wl(db, r.request_id, start=f"{day.isoformat()} 17:20:00", end=f"{day.isoformat()} 20:00:00",
           actual=net + brk, brk=brk, net=net)
    return r


@pytest.fixture
def month_fixture(db):
    rec(db, "製造1課", "山田", date(2026, 11, 2), 1500)
    rec(db, "製造1課", "山田", date(2026, 11, 8), 1000, status="submitted", rtype="holiday", minutes=480)
    rec(db, "製造1課", "佐藤", date(2026, 11, 3), 1800, pdf="pdf-1")
    rec(db, "製造2課", "鈴木", date(2026, 11, 4), 100, pdf="pdf-2")
    rec(db, "品質管理", "高橋", date(2026, 11, 30), 3600)
    rec(db, "製造1課", "山田", date(2026, 11, 5), 999, status="canceled")
    rec(db, "製造1課", "山田", date(2026, 10, 31), 999)
    rec(db, "製造1課", "山田", date(2026, 12, 1), 999)


def test_monthly_summary_numbers(db, month_fixture):
    res = summaries.api_adminMonthlySummary(ctx(db, email=ADMIN, now=NOW), "2026-11", "ALL")
    assert res["yearMonth"] == "2026-11" and res["deptFilter"] == "ALL"
    assert res["monthInfo"] == {"daysInMonth": 30, "dayOfMonth": 26, "isCurrentMonth": True}
    assert res["kpi"] == {"totalPeople": 4, "over40": 2, "over60": 1, "projectedOver40": 2, "projectedOver60": 1,
                          "pdfMissingTotal": 2, "totalNetMinutes": 8000}
    assert [(p["workerName"], p["totalNet"], p["projectedTotal"]) for p in res["people"]] == [
        ("高橋", 3600, 4154), ("山田", 2500, 2885), ("佐藤", 1800, 2077), ("鈴木", 100, 115)]
    yamada = res["people"][1]
    assert yamada == {
        "dept": "製造1課", "workerName": "山田", "overtimeNet": 1500, "holidayNet": 1000, "totalNet": 2500,
        "approvedTotal": 540, "pdfMissing": 1, "pacePerDay": 2500 / 26, "projectedTotal": 2885,
        "projectedOver40": True, "projectedOver60": False, "remain40": 0, "remain60": 1100,
    }
    assert [w["workerName"] for w in res["watch"]] == ["高橋", "山田", "佐藤"]  # 1800 ちょうども入る（>=）
    assert set(res["watch"][0]) == {"dept", "workerName", "totalNet", "projectedTotal", "remain40", "remain60",
                                    "projectedOver40", "projectedOver60", "pdfMissing"}
    assert res["charts"]["dept"] == {"labels": ["製造1課", "製造2課", "品質管理"], "values": [4300, 100, 3600]}
    assert res["charts"]["people"] == {"labels": ["高橋", "山田", "佐藤", "鈴木"], "values": [3600, 2500, 1800, 100],
                                       "projected": [4154, 2885, 2077, 115], "limit40": 2400, "limit60": 3600}


def test_monthly_summary_dept_filter_past_and_future_months(db, month_fixture):
    res = summaries.api_adminMonthlySummary(ctx(db, email=ADMIN, now=NOW), "2026-11", "製造1課")
    assert res["deptFilter"] == "製造1課" and [p["workerName"] for p in res["people"]] == ["山田", "佐藤"]
    past = summaries.api_adminMonthlySummary(ctx(db, email=ADMIN, now=NOW), "2026-10", None)
    assert past["deptFilter"] == "ALL" and past["monthInfo"] == {"daysInMonth": 31, "dayOfMonth": 31,
                                                                  "isCurrentMonth": False}
    future = summaries.api_adminMonthlySummary(ctx(db, email=ADMIN, now=NOW), "2026-12", "ALL")
    assert future["monthInfo"]["dayOfMonth"] == 31  # 未来月も月末まで経過（Q15）
    assert future["people"][0]["pacePerDay"] == pytest.approx(999 / 31)


@pytest.mark.parametrize("ym", ["2026-13", "2026", "", None, "abc-11", "2026-0"])
def test_monthly_summary_bad_month(db, ym):
    with pytest.raises(RpcError) as e:
        summaries.api_adminMonthlySummary(ctx(db, email=ADMIN), ym, "ALL")
    assert e.value.message == "yearMonthは yyyy-mm で指定してください"


def test_daily_detail_returns_strings_and_retro_flags(db):
    a = rec(db, "製造1課", "山田", date(2026, 11, 20), 120, submitted=jst(2026, 11, 22, 6, 30))
    b = rec(db, "品質管理", "高橋", date(2026, 11, 20), 60, pdf="p1")
    c = rec(db, "製造1課", "佐藤", date(2026, 11, 20), 0)
    rec(db, "製造1課", "山田", date(2026, 11, 21), 50)
    res = summaries.api_adminDailyDetail(ctx(db, email=ADMIN, now=NOW), "2026-11-20", "ALL")
    assert res["date"] == "2026-11-20" and res["deptFilter"] == "ALL" and res["total"] == 3
    assert res["pdfMissing"] == 1  # approved かつ net > 0 かつ PDF 無し
    assert [i["requestId"] for i in res["items"]] == [c.request_id, a.request_id, b.request_id]  # 製造1課佐藤 < 製造1課山田 < 品質
    item = next(i for i in res["items"] if i["requestId"] == b.request_id)
    assert item == {
        "requestId": b.request_id, "status": "approved", "requestType": "overtime", "dept": "品質管理",
        "workerCode": "", "workerName": "高橋", "targetDate": "2026-11-20", "approvedMinutes": 60,
        "actualStartAt": "2026-11-20 17:20:00", "actualEndAt": "2026-11-20 20:00:00", "actualMinutes": 60,
        "breakMinutes": 0, "netMinutes": 60, "pdfFileId": "p1", "pdfGeneratedAt": "2026-11-26 20:00:00",
        "isRetroactive": False, "daysAgo": 6,
    }
    assert next(i for i in res["items"] if i["requestId"] == a.request_id)["isRetroactive"] is True
    only = summaries.api_adminDailyDetail(ctx(db, email=ADMIN, now=NOW), "2026-11-20", "品質管理")
    assert only["total"] == 1 and only["deptFilter"] == "品質管理"
    with pytest.raises(RpcError) as e:
        summaries.api_adminDailyDetail(ctx(db, email=ADMIN), "x", "ALL")
    assert e.value.message == "日付が不正です（yyyy-mm-dd）"


def test_fiscal_trend_uses_april_first(db):
    rec(db, "A", "a", date(2026, 3, 31), 10)
    rec(db, "A", "a", date(2026, 4, 1), 20)
    rec(db, "B", "b", date(2026, 4, 30), 5, status="submitted")
    rec(db, "A", "a", date(2027, 3, 31), 40)
    rec(db, "A", "a", date(2027, 4, 1), 80)
    res = summaries.api_adminFiscalTrend(ctx(db, email=ADMIN, now=NOW), "2026-11-26", "ALL")
    assert res["fiscalYearStart"] == "2026-04-01" and res["fiscalYearEnd"] == "2027-04-01"
    assert res["labels"][0] == "2026-04" and res["labels"][-1] == "2027-03" and len(res["labels"]) == 12
    assert res["values"][0] == 25 and res["values"][-1] == 40 and sum(res["values"]) == 65
    assert summaries.api_adminFiscalTrend(ctx(db, email=ADMIN), "2026-11-26", "B")["values"][0] == 5
    assert summaries.api_adminFiscalTrend(ctx(db, email=ADMIN), "2027-03-31", "ALL")["fiscalYearStart"] == "2026-04-01"
    with pytest.raises(RpcError) as e:
        summaries.api_adminFiscalTrend(ctx(db, email=ADMIN), "", "ALL")
    assert e.value.message == "baseDateYmdが不正です（yyyy-mm-dd）"


def test_special_clause_count(db):
    months = [date(2026, m, 10) for m in (4, 5, 6, 7, 8, 9)]
    for d in months:
        rec(db, "A", "six", d, 3600)
    for d in months[:5]:
        rec(db, "A", "five", d, 1800)
        rec(db, "A", "five", d.replace(day=11), 1800, status="submitted")  # submitted も数える
    for d in months[:4]:
        rec(db, "B", "four", d, 3600)
    rec(db, "B", "four", date(2026, 3, 31), 3600)  # 前の年度
    rec(db, "B", "few", date(2026, 4, 10), 3599)
    res = summaries.api_adminSpecialClauseCount(ctx(db, email=ADMIN, now=NOW), "2026-11-26", "ALL")
    assert res["fiscalYear"] == "2026-04-01 ~ 2027-03-31"
    assert res["list"] == [
        {"dept": "A", "workerName": "six", "count60": 6, "isDanger": True, "isWarn": True},
        {"dept": "A", "workerName": "five", "count60": 5, "isDanger": False, "isWarn": True},
        {"dept": "B", "workerName": "four", "count60": 4, "isDanger": False, "isWarn": False},
    ]
    assert [x["workerName"] for x in summaries.api_adminSpecialClauseCount(
        ctx(db, email=ADMIN), "2026-11-26", "B")["list"]] == ["four"]
    assert summaries.api_adminSpecialClauseCount(ctx(db, email=ADMIN), "2026-03-31", "ALL")["fiscalYear"] == (
        "2025-04-01 ~ 2026-03-31")
    with pytest.raises(RpcError) as e:
        summaries.api_adminSpecialClauseCount(ctx(db, email=ADMIN), "?", "ALL")
    assert e.value.message == "基準日が不正です"


def test_pending_watch(db):
    a = add_request(db, submitted_at=jst(2026, 11, 24, 10), target=date(2026, 11, 20))  # ちょうど 48 時間
    b = add_request(db, submitted_at=jst(2026, 11, 24, 10, 0, 1), target=date(2026, 11, 26),
                    request_type="holiday", dept="B", worker_name="b")
    add_request(db, status="approved")
    res = summaries.api_adminPendingWatch(ctx(db, email=ADMIN, now=NOW))
    assert res == [
        {"requestType": "overtime", "dept": "製造1課", "workerName": "山田太郎", "targetDateLabel": "11/20(金)",
         "isPastDate": True, "submittedAt": "2026-11-24 10:00:00", "hoursPending": 48, "isOver48h": True,
         "targetDate": "2026-11-20", "isRetroactive": True, "daysAgo": 6},
        {"requestType": "holiday", "dept": "B", "workerName": "b", "targetDateLabel": "11/26(木)",
         "isPastDate": False, "submittedAt": "2026-11-24 10:00:01", "hoursPending": 47, "isOver48h": False,
         "targetDate": "2026-11-26", "isRetroactive": False, "daysAgo": 0},
    ]
    assert "requestId" not in res[0]  # G18: 足さない
    _ = (a, b)


def test_dept_options(db):
    add_worker(db, 1, "A1", "x", "製造2課")
    add_worker(db, 2, "A2", "y", "品質管理")
    add_worker(db, 3, "A3", "z", "製造1課")
    assert summaries.api_adminDeptOptions(ctx(db, email=ADMIN)) == ["ALL", "品質管理", "製造1課", "製造2課"]


# ===== スタッフ集計 =====


@pytest.fixture
def staff_fixture(db):
    # 締め月 2026-11 = 11/16〜12/15。年度（3/16 始まり）= 2026-03-16〜2027-03-16
    rec(db, "製造1課", "山田", date(2026, 11, 16), 1200, code="A001")
    rec(db, "製造1課", "山田", date(2026, 12, 15), 1300, code="A001", rtype="holiday", pdf="p")
    rec(db, "製造1課", "山田", date(2026, 12, 1), 0, code="A001", minutes=90)  # 承認済みだが実績なし → 見込み
    rec(db, "製造1課", "山田", date(2026, 11, 20), 500, code="A001", status="submitted", minutes=120)
    rec(db, "製造1課", "山田", date(2026, 11, 15), 2400, code="A001")  # 前の締め月（10 月分）
    rec(db, "製造1課", "山田", date(2026, 3, 15), 9999, code="A001")  # 前の年度
    rec(db, "製造2課", "鈴木", date(2026, 11, 30), 300)  # コードなし → 氏名で
    rec(db, "製造2課", "鈴木", date(2026, 12, 16), 300)  # 翌締め月


def test_closing_summary_personal(db, staff_fixture):
    res = summaries.api_getClosingSummary(ctx(db, now=NOW), "2026-11", "", "A001 山田")
    assert res["mode"] == "personal"
    assert res["meta"] == {"yearMonth": "2026-11", "monthLabel": "2026年 11月分", "periodFrom": "11/16",
                           "periodTo": "12/15", "fyLabel": "2026年度", "limit40": 2400, "limit60": 3600}
    assert res["worker"] == {"workerCode": "A001", "workerName": "山田", "dept": "製造1課"}
    assert res["kpi"] == {"otNet": 1200, "holNet": 1300, "totalNet": 2500, "pendMin": 210, "pendCount": 2}
    assert [d["date"] for d in res["detail"]] == ["11/16", "11/20", "12/1", "12/15"]
    assert res["detail"][-1] == {"date": "12/15", "type": "holiday", "approvedMinutes": 60, "netMinutes": 1300,
                                 "breakMinutes": 0, "status": "approved", "pdfFileId": "p"}
    assert res["annualOver"] == 2  # 10 月分 2400（>=）と 11 月分 2500


def test_closing_summary_list_and_default_month(db, staff_fixture):
    res = summaries.api_getClosingSummary(ctx(db, now=jst(2026, 12, 15, 23, 30)), "", "ALL", "")
    assert res["mode"] == "list" and res["dept"] == "" and res["meta"]["yearMonth"] == "2026-11"
    assert res["people"] == [
        {"workerCode": "A001", "workerName": "山田", "dept": "製造1課", "otNet": 1200, "holNet": 1300,
         "totalNet": 2500, "pendMin": 210, "annualOver": 2},
        {"workerCode": "", "workerName": "鈴木", "dept": "製造2課", "otNet": 300, "holNet": 0, "totalNet": 300,
         "pendMin": 0, "annualOver": 0},
    ]
    nxt = summaries.api_getClosingSummary(ctx(db, now=jst(2026, 12, 16, 0, 30)), None, "製造2課", None)
    assert nxt["meta"]["yearMonth"] == "2026-12" and nxt["dept"] == "製造2課"
    assert [p["workerName"] for p in nxt["people"]] == ["鈴木"]


def test_closing_summary_by_name_and_bad_month(db, staff_fixture):
    res = summaries.api_getClosingSummary(ctx(db, now=NOW), "2026-11", "製造2課", "X999 鈴木")
    assert res["kpi"]["otNet"] == 300 and res["worker"] == {"workerCode": "X999", "workerName": "鈴木",
                                                            "dept": "製造2課"}
    empty = summaries.api_getClosingSummary(ctx(db, now=NOW), "2026-11", "", "Z9 誰か")
    assert empty["detail"] == [] and empty["worker"] == {"workerCode": "Z9", "workerName": "誰か", "dept": ""}
    with pytest.raises(RpcError) as e:
        summaries.api_getClosingSummary(ctx(db), "2026-13", "", "")
    assert e.value.message == "yearMonth は yyyy-MM で指定してください: 2026-13"
