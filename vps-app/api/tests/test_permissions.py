"""権限（実装仕様 8-6）: 2 キーの和・空でも全員にしない・DeptApprovers があればそれだけ・ApproverMap の 1 行目見出し。"""
from __future__ import annotations

from app import permissions
from app.models import ApproverMapRow, DeptApprover, Setting


def add_settings(db, **kv):
    for i, (k, v) in enumerate(kv.items()):
        db.add(Setting(key=k, value=v, sheet_row=3 + i))
    db.commit()


def add_dept(db, *rows):
    for i, (dept, emails, name) in enumerate(rows):
        db.add(DeptApprover(sheet_row=2 + i, dept=dept, emails=emails, name=name))
    db.commit()


def add_approver_map(db, *rows):
    for i, cells in enumerate(rows):
        db.add(ApproverMapRow(sheet_row=1 + i, cells=list(cells)))
    db.commit()


AM_HEADER = ["部署", "承認者メール", "role(approver/admin)", "有効フラグ"]


def test_admin_is_union_of_two_keys_and_bootstrap(db, settings):
    add_settings(db, ADMIN_EMAILS=" A@lineworks-local.info , b@lineworks-local.info",
                 GENERAL_AFFAIRS_EMAILS="ga@lineworks-local.info")
    for email in ("a@lineworks-local.info", "b@lineworks-local.info", "ga@lineworks-local.info",
                  "owner@lineworks-local.info", "A@lineworks-local.info"):
        assert permissions.is_admin(db, email, settings), email
    assert not permissions.is_admin(db, "someone@lineworks-local.info", settings)
    assert not permissions.is_admin(db, "", settings)


def test_empty_admin_list_never_makes_everyone_admin(db, settings):
    add_settings(db, ADMIN_EMAILS="", GENERAL_AFFAIRS_EMAILS="")
    assert not permissions.is_admin(db, "someone@lineworks-local.info", settings)
    assert permissions.admin_emails(db, settings) == ["owner@lineworks-local.info"]


def test_admin_from_approver_map_header_in_first_row(db, settings):
    add_approver_map(db, AM_HEADER,
                     ["製造1課", "boss@lineworks-local.info", "admin", True],
                     ["製造1課", "off@lineworks-local.info", "admin", False],
                     ["製造1課", "off2@lineworks-local.info", "admin", "FALSE"],
                     ["製造1課", "appr@lineworks-local.info", "approver", True])
    assert permissions.is_admin(db, "boss@lineworks-local.info", settings)
    assert not permissions.is_admin(db, "off@lineworks-local.info", settings)
    assert not permissions.is_admin(db, "off2@lineworks-local.info", settings)
    assert not permissions.is_admin(db, "appr@lineworks-local.info", settings)


def test_approver_map_today_has_description_row_so_it_does_not_work(db, settings):
    """今の ApproverMap は 1 行目が説明・2 行目が見出し。GAS は 1 行目を見出しとして読むので効かない（今と同じ）。"""
    add_approver_map(db, ["ApproverMap の説明"], AM_HEADER, ["製造1課", "boss@lineworks-local.info", "admin", True])
    assert not permissions.is_admin(db, "boss@lineworks-local.info", settings)
    assert permissions.dept_approver_map(db) == {}


def test_dept_approvers_with_rows_are_used_exclusively(db, settings):
    add_dept(db, ("製造１課", "X@lineworks-local.info, y@lineworks-local.info", "田中"),
             ("製造1課", "z@lineworks-local.info", ""),
             ("", "ignored@lineworks-local.info", ""),
             ("総務", "x@lineworks-local.info", "田中"))
    add_approver_map(db, AM_HEADER, ["組立", "am@lineworks-local.info", "approver", True])
    assert permissions.dept_approver_map(db) == {
        "製造1課": ["x@lineworks-local.info", "y@lineworks-local.info", "z@lineworks-local.info"],
        "総務": ["x@lineworks-local.info"],
    }
    assert permissions.can_approve_dept(db, "X@lineworks-local.info", "製造1課", settings)
    assert permissions.can_approve_dept(db, "z@lineworks-local.info", "製造１課", settings)
    assert not permissions.can_approve_dept(db, "z@lineworks-local.info", "総務", settings)
    assert not permissions.can_approve_dept(db, "am@lineworks-local.info", "組立", settings)  # ApproverMap は見ない
    assert permissions.can_approve_dept(db, "owner@lineworks-local.info", "どこでも", settings)  # 管理者は全部署
    assert permissions.is_any_dept_approver(db, "y@lineworks-local.info", settings)
    assert not permissions.is_any_dept_approver(db, "am@lineworks-local.info", settings)
    assert not permissions.is_any_dept_approver(db, "", settings)


def test_falls_back_to_approver_map_without_dept_rows(db, settings):
    add_approver_map(db, AM_HEADER,
                     ["組立", " AM@lineworks-local.info ", "approver", True],
                     ["組立", "off@lineworks-local.info", "approver", False],
                     ["", "nodept@lineworks-local.info", "approver", True])
    assert permissions.dept_approver_map(db) == {"組立": ["am@lineworks-local.info"]}
    assert permissions.can_approve_dept(db, "am@lineworks-local.info", "組立", settings)
    assert permissions.is_any_dept_approver(db, "am@lineworks-local.info", settings)


def test_approver_map_without_needed_columns_is_empty(db):
    add_approver_map(db, ["dept", "mail"], ["組立", "a@lineworks-local.info"])
    assert permissions.dept_approver_map(db) == {}


def test_approver_profiles(db):
    add_dept(db, ("製造1課", "x@lineworks-local.info,y@lineworks-local.info", ""),
             ("総務", "x@lineworks-local.info", "田中太郎"),
             ("組立", "", "空"),
             ("製造1課", "x@lineworks-local.info", ""))
    assert permissions.approver_profiles(db) == [
        {"email": "x@lineworks-local.info", "name": "田中太郎", "depts": ["製造1課", "総務"]},
        {"email": "y@lineworks-local.info", "name": "y@lineworks-local.info", "depts": ["製造1課"]},
    ]
    assert permissions.approver_profiles.__doc__


def test_no_profiles_without_dept_approvers(db):
    add_approver_map(db, AM_HEADER, ["組立", "a@lineworks-local.info", "approver", True])
    assert permissions.approver_profiles(db) == []
