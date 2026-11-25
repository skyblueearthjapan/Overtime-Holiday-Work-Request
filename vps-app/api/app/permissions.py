"""権限（実装仕様 8-6）。GAS の isAdmin_・getDeptApproverMap_・canApproveDept_・getApproverProfiles_ を写す。

- 管理者 = Settings の ADMIN_EMAILS と GENERAL_AFFAIRS_EMAILS（キーの前後の空白を取る）の和（カンマ区切り、
  小文字で比べる）＋ BOOTSTRAP_ADMINS。一覧が空でも全員を管理者にしない（罠 35）。
  ApproverMap の role=admin（1 行目を見出しとして読む）も GAS と同じに見る。
- 部署の承認者 = DeptApprovers にデータ行があればそれだけ（A 部署・B メールのカンマ区切りを小文字に）。
  無ければ ApproverMap（1 行目見出し: 部署・承認者メール・有効フラグ）。管理者は全部署。
- 表は config_sync（担当 D）が共有タブから取り込んだもの（settings・dept_approvers・approver_map_rows）。
"""
from __future__ import annotations

from typing import Any

from sqlalchemy import select
from sqlalchemy.orm import Session

from app.config import Settings, get_settings
from app.jsutil import js_string, normalize_gas
from app.models import ApproverMapRow, DeptApprover, Setting

ADMIN_SETTING_KEYS = ("ADMIN_EMAILS", "GENERAL_AFFAIRS_EMAILS")


def settings_map(db: Session) -> dict[str, str]:
    """Settings の key → value（config_sync が同じキーは下の行を残している）。"""
    return {row.key: row.value for row in db.scalars(select(Setting))}


def _split_lower(raw: str) -> list[str]:
    return [s.strip().lower() for s in (raw or "").split(",") if s.strip()]


def admin_emails(db: Session, settings: Settings | None = None) -> list[str]:
    """Settings の 2 つのキーの和 ＋ BOOTSTRAP_ADMINS（小文字、重なりなし）。"""
    s = settings or get_settings()
    values = settings_map(db)
    combined = ",".join(js_string(values.get(k, "")) for k in ADMIN_SETTING_KEYS)
    out: list[str] = []
    for email in _split_lower(combined) + list(s.bootstrap_admin_list):
        if email not in out:
            out.append(email)
    return out


def _approver_map_values(db: Session) -> list[list[Any]]:
    rows = db.scalars(select(ApproverMapRow).order_by(ApproverMapRow.sheet_row)).all()
    return [list(r.cells or []) for r in rows]


def _cell(row: list[Any], idx: int) -> Any:
    """JS の row[idx]（idx が -1 や範囲外なら undefined）。"""
    if idx < 0 or idx >= len(row):
        return None
    return row[idx]


def _enabled(row: list[Any], idx: int) -> bool:
    """GAS: enabled = idx>=0 ? row[idx] : true。false か 'false'（大文字小文字を問わず）なら無効。"""
    if idx < 0:
        return True
    value = _cell(row, idx)
    if value is False:
        return False
    return js_string(value).lower() != "false" if value is not None else True


def is_admin(db: Session, email: str, settings: Settings | None = None) -> bool:
    """総務部（管理者）か。メールが空なら偽。"""
    if not email:
        return False
    if email.lower() in admin_emails(db, settings):
        return True
    values = _approver_map_values(db)
    if not values:
        return False
    header = [normalize_gas(h) for h in values[0]]
    idx_mail = header.index("承認者メール") if "承認者メール" in header else -1
    idx_role = header.index("role(approver/admin)") if "role(approver/admin)" in header else -1
    idx_enabled = header.index("有効フラグ") if "有効フラグ" in header else -1
    for row in values[1:]:
        if not _enabled(row, idx_enabled):
            continue
        # GAS は mail === email（大文字小文字を区別する）と role === 'admin'
        if normalize_gas(_cell(row, idx_mail)) == email and normalize_gas(_cell(row, idx_role)) == "admin":
            return True
    return False


def _dept_rows(db: Session) -> list[DeptApprover]:
    return list(db.scalars(select(DeptApprover).where(DeptApprover.sheet_row >= 2).order_by(DeptApprover.sheet_row)))


def dept_approver_map(db: Session) -> dict[str, list[str]]:
    """部署 → 承認者のメール（小文字、重なりなし、出てきた順）。"""
    result: dict[str, list[str]] = {}
    rows = _dept_rows(db)
    if rows:
        for row in rows:
            dept = normalize_gas(row.dept)
            raw = normalize_gas(row.emails)
            if not dept:
                continue
            bucket = result.setdefault(dept, [])
            for email in _split_lower(raw):
                if email not in bucket:
                    bucket.append(email)
        return result

    values = _approver_map_values(db)
    if not values:
        return {}
    header = [normalize_gas(h) for h in values[0]]
    idx_dept = header.index("部署") if "部署" in header else -1
    idx_mail = header.index("承認者メール") if "承認者メール" in header else -1
    idx_enabled = header.index("有効フラグ") if "有効フラグ" in header else -1
    if idx_dept < 0 or idx_mail < 0:
        return {}
    for row in values[1:]:
        if not _enabled(row, idx_enabled):
            continue
        dept = normalize_gas(_cell(row, idx_dept))
        mail = normalize_gas(_cell(row, idx_mail)).lower()
        if not dept or not mail:
            continue
        bucket = result.setdefault(dept, [])
        if mail not in bucket:
            bucket.append(mail)
    return result


def can_approve_dept(db: Session, email: str, dept: str, settings: Settings | None = None) -> bool:
    """その部署を承認できるか（管理者は全部署）。"""
    if is_admin(db, email, settings):
        return True
    if not email:
        return False
    return email.lower() in dept_approver_map(db).get(normalize_gas(dept), [])


def is_any_dept_approver(db: Session, email: str, settings: Settings | None = None) -> bool:
    """どこかの部署の承認者か（承認者画面の入口。管理者は別に判定する）。"""
    if not email:
        return False
    target = email.lower()
    return any(target in emails for emails in dept_approver_map(db).values())


def approver_profiles(db: Session) -> list[dict]:
    """DeptApprovers の承認者一覧 [{email, name, depts}]（getApproverProfiles_ と同じ。ApproverMap は見ない）。"""
    profiles: dict[str, dict] = {}
    for row in _dept_rows(db):
        dept = normalize_gas(row.dept)
        raw = normalize_gas(row.emails)
        name = normalize_gas(row.name)
        if not dept or not raw:
            continue
        for email in _split_lower(raw):
            profile = profiles.setdefault(email, {"email": email, "name": name or email, "depts": []})
            if dept not in profile["depts"]:
                profile["depts"].append(dept)
            if name and name != email:
                profile["name"] = name
    return list(profiles.values())
