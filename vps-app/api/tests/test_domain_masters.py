"""マスタを読む関数（4-3 の 1〜5。masters.gs・deptAuth.gs・formMap.gs:222-470）。"""
from __future__ import annotations

from datetime import date

from app.domain import masters
from test_domain_helpers import USER, add_calendar, add_order, add_worker, ctx, jst


def test_dept_list_and_workers_by_dept(db):
    add_worker(db, 1, "B002", "佐藤", "製造1課")
    add_worker(db, 2, "A001", "山田", "製造1課", active="1")
    add_worker(db, 3, "A010", "退職", "製造1課", active="0")  # 在籍フラグが 0 → 除く
    add_worker(db, 4, "C001", "鈴木", "Ｚ部", active="TRUE")
    add_worker(db, 5, "", "名無し", "空部署")  # コードが空 → 除く
    add_worker(db, 6, "D001", "高橋", "品質管理", active="true")
    assert masters.api_getDeptList(ctx(db)) == ["Z部", "品質管理", "製造1課"]  # UTF-16 の符号順
    assert masters.api_getWorkersByDept(ctx(db), "製造1課") == ["A001 山田", "B002 佐藤"]
    assert masters.api_getWorkersByDept(ctx(db), "") == []
    assert masters.api_getWorkersByDept(ctx(db)) == []
    assert masters.api_getWorkersByDept(ctx(db), "無い部署") == []
    assert masters.api_getWorkersByDept(ctx(db), " 製造1課") == []  # 部署は正規化せずに引く


def test_worker_info_is_null_without_google_column(db):
    add_worker(db, 1, "A001", "山田", "製造1課")
    assert masters.api_getWorkerInfo(ctx(db)) is None


def test_worker_info_case_sensitive_match(db):
    add_worker(db, 1, "A001", "山田", "製造1課", google="Tanaka@lineworks-local.info")
    add_worker(db, 2, "A002", "田中", "製造2課", google=USER)
    assert masters.api_getWorkerInfo(ctx(db)) == {"workerCode": "A002", "workerName": "田中", "dept": "製造2課",
                                                  "email": USER}
    assert masters.api_getWorkerInfo(ctx(db, email="nobody@lineworks-local.info")) is None


def test_order_list_dedupes_and_sorts(db):
    add_order(db, 1, "LW200", customer="c", dest="d2", product="p2")
    add_order(db, 2, "LW100", dest="d1", product="p1")
    add_order(db, 3, "LW200", dest="dup", product="dup")
    add_order(db, 4, "", dest="x")
    add_order(db, 5, "ＡＢ1", product="全角")
    assert masters.api_getOrderList(ctx(db)) == [
        {"code": "AB1", "product": "全角", "dest": ""},
        {"code": "LW100", "product": "p1", "dest": "d1"},
        {"code": "LW200", "product": "p2", "dest": "d2"},
    ]


# ===== 休日の候補日 =====

THU = jst(2026, 11, 26, 6, 30)  # 木曜の朝（UTC では前日 21:30）


def test_candidates_from_company_calendar(db):
    add_calendar(db, 1, "日付", "区分", "曜日", "備考")  # 見出しのような行は飛ばす
    add_calendar(db, 2, "2026-11-22", "休日", "日")  # 今日より前
    add_calendar(db, 3, "2026-11-28", "出勤土曜", "土")
    add_calendar(db, 4, "2026-11-29", "休日", "")
    add_calendar(db, 5, "2026-12-05", "休日", "土", "棚卸")
    add_calendar(db, 6, "2026-11-26", "祝日", "木", "創立記念日")
    add_calendar(db, 7, "2026-12-04", "休日", "金")  # 来週金曜（範囲の端）
    add_calendar(db, 8, "2026-12-06", "休日", "日")  # 範囲外
    res = masters.api_getHolidayCandidateDates(ctx(db, now=THU))
    assert res == [
        {"date": "2026-11-26", "label": "11/26(木) 創立記念日", "type": "holiday", "holidayName": "創立記念日"},
        {"date": "2026-11-29", "label": "11/29(日)", "type": "weekend"},
        {"date": "2026-12-04", "label": "12/4(金)", "type": "weekend"},
    ]


def test_candidates_fall_back_when_calendar_is_empty(db):
    res = masters.api_getHolidayCandidateDates(ctx(db, now=THU))
    assert res == [
        {"date": "2026-11-28", "label": "11/28(土)", "type": "weekend"},
        {"date": "2026-11-29", "label": "11/29(日)", "type": "weekend"},
    ]


def test_candidates_fall_back_includes_computed_holidays(db):
    # 2026-11-02(月) の週: 11/3 文化の日（火）。来週金曜は 11/13
    res = masters.legacy_candidates(date(2026, 11, 2))
    assert res == [
        {"date": "2026-11-03", "label": "11/3(火) 文化の日", "type": "holiday"},
        {"date": "2026-11-07", "label": "11/7(土)", "type": "weekend"},
        {"date": "2026-11-08", "label": "11/8(日)", "type": "weekend"},
    ]


def test_known_holidays_substitute_and_weekend_merge(db):
    # 2026-05-03（日）憲法記念日 → 振替は 5/4 だが みどりの日 が後で上書き、5/6 は今の GAS では作らない
    hol = {h["date"]: h["title"] for h in masters.known_japanese_holidays(date(2026, 4, 27), date(2026, 5, 8))}
    assert hol == {"2026-04-29": "昭和の日", "2026-05-03": "憲法記念日", "2026-05-04": "みどりの日",
                   "2026-05-05": "こどもの日"}
    res = masters.legacy_candidates(date(2026, 4, 27))
    assert {"date": "2026-05-03", "label": "5/3(日)", "type": "weekend", "holidayName": "憲法記念日"} in res
    # 春分・秋分・ハッピーマンデー
    assert masters.known_japanese_holidays(date(2026, 3, 16), date(2026, 3, 27)) == [
        {"date": "2026-03-20", "title": "春分の日"}]
    assert {h["date"] for h in masters.known_japanese_holidays(date(2026, 9, 14), date(2026, 9, 25))} == {
        "2026-09-21", "2026-09-23"}
    assert masters.known_japanese_holidays(date(2027, 1, 4), date(2027, 1, 15)) == [
        {"date": "2027-01-11", "title": "成人の日"}]


def test_candidates_bad_calendar_date_falls_back(db):
    add_calendar(db, 1, "2026-11-2９x", "休日")  # 形が壊れていれば GAS の catch と同じく従来方式
    res = masters.api_getHolidayCandidateDates(ctx(db, now=THU))
    assert [c["date"] for c in res] == ["2026-11-28", "2026-11-29"]
