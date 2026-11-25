"""安全装置（実装仕様 9 章。最重要）。書き込み先を許可リストで決め、環境に合わない先へは書かせない。

- 書いてよいスプレッドシート: production は PROD_SPREADSHEET_ID だけ、staging は STAGING_SPREADSHEET_ID だけ。
- 書いてよいタブ: A 区分（Requests・WorkLogs・BatchLogs）は row_sync・reconcile だけ、
  C 区分（作業員マスタ・工番マスタ・社内カレンダーマスタ・業務NOマスタ）は master_sync だけ。
- どの環境も書かない: マスタ元・蓄積台帳（両方の環境）・PDF フォルダ・ほかのスプレッドシート、B・D・E 区分のタブ。
- 使ってよい書き方: updateCells・values.batchUpdate・values.batchClear（C 区分の余りの行だけ）・
  appendDimension（ROWS、A 区分の末尾だけ）。行・列・タブの挿入・削除・並べ替え・作り直しは使わない。
- Drive: api・worker は呼ばない（このモジュールに Drive の許可は無い）。

確かめる時: 起動時（verify_startup。api と worker）と、書き込みの直前
（services/sheets_api.SheetsApi の書く関数の中で assert_writable と assert_write_requests）。
GAS の内部用の入口は verify_gas_ping で照合する（合わなければ GAS のレーンを止める＝担当 M）。
"""
from __future__ import annotations

import logging
from collections.abc import Iterable, Mapping
from typing import Any

from app import config as _config
from app.config import Settings, get_settings

log = logging.getLogger(__name__)


class UnsafeWriteError(RuntimeError):
    """許可していない書き込み先・書き方。処理を止める（握りつぶさない。再試行しない）。"""


# 区分ごとのタブ（実装仕様 6-1）
A_TABS = frozenset({"Requests", "WorkLogs", "BatchLogs"})
C_TABS = frozenset({"作業員マスタ", "工番マスタ", "社内カレンダーマスタ", "業務NOマスタ"})

# 書く目的 → 書いてよいタブ
PURPOSE_TABS: dict[str, frozenset[str]] = {
    "row_sync": A_TABS,
    "reconcile": A_TABS,
    "master_sync": C_TABS,
}

# spreadsheets.batchUpdate で使ってよい要求の種類
ALLOWED_BATCH_REQUESTS = frozenset({"updateCells", "appendDimension"})
# 使ってよい書き方（SheetsApi の関数ごとの名前）
WRITE_METHODS = frozenset({"updateCells", "values.batchUpdate", "values.batchClear", "appendDimension"})


def _settings(settings: Settings | None) -> Settings:
    return settings or get_settings()


def never_writable_ids(settings: Settings | None = None) -> frozenset[str]:
    """どの環境からも書かない ID（マスタ元・台帳・PDF フォルダ。両方の環境の分）。"""
    s = _settings(settings)
    ids = {
        _config.MASTER_SOURCE_SPREADSHEET_ID_FIXED,
        _config.PROD_LEDGER_SPREADSHEET_ID_FIXED,
        _config.PROD_PDF_FOLDER_ID_FIXED,
        _config.PROD_GAS_SCRIPT_ID_FIXED,
        _config.STAGING_LEDGER_SPREADSHEET_ID_KNOWN,
        _config.STAGING_PDF_FOLDER_ID_KNOWN,
        _config.STAGING_GAS_SCRIPT_ID_KNOWN,
        s.prod_ledger_spreadsheet_id,
        s.staging_ledger_spreadsheet_id,
        s.prod_pdf_folder_id,
        s.staging_pdf_folder_id,
    }
    return frozenset(i for i in ids if i)


def writable_spreadsheet_id(settings: Settings | None = None) -> str:
    """この環境で書いてよいただ 1 つのスプレッドシート（DB）。"""
    return _settings(settings).db_spreadsheet_id


def assert_writable(spreadsheet_id: str, tab: str, purpose: str, settings: Settings | None = None) -> str:
    """書き込みの直前に呼ぶ。スプレッドシート・タブ・目的が許可リストに合わなければ UnsafeWriteError。"""
    s = _settings(settings)
    sid = (spreadsheet_id or "").strip()
    if not sid:
        raise UnsafeWriteError("書き込み先のスプレッドシート ID が空です")
    if sid in never_writable_ids(s):
        raise UnsafeWriteError("マスタ元・蓄積台帳・PDF フォルダへは書き込めません")
    if s.is_staging and sid in _config.PRODUCTION_IDS:
        raise UnsafeWriteError("ステージングから本番のスプレッドシートへは書き込めません")
    if not s.is_staging and sid in _config.KNOWN_STAGING_IDS:
        raise UnsafeWriteError("本番からステージングのスプレッドシートへは書き込めません")
    if sid != writable_spreadsheet_id(s):
        raise UnsafeWriteError("書き込みを許可していないスプレッドシートです")
    allowed_tabs = PURPOSE_TABS.get(purpose)
    if allowed_tabs is None:
        raise UnsafeWriteError(f"知らない書き込みの目的です: {purpose}")
    if tab not in allowed_tabs:
        raise UnsafeWriteError(f"タブ「{tab}」へは {purpose} から書き込めません")
    return sid


# 戻しの道具（tools/restore_tabs_from_copy.py）が控えから写し戻してよいタブ。
# B 区分の 3 タブは、道具の引数でタブの名前をはっきり書いたときだけ（protected_listed=True）。
RESTORE_PROTECTED_TABS = frozenset({"Settings", "DeptApprovers", "ApproverMap"})
RESTORE_TABS = A_TABS | C_TABS


def assert_restore_target(spreadsheet_id: str, tab: str, settings: Settings | None = None, *,
                          protected_listed: bool = False) -> str:
    """戻しの道具の書き込み先の確かめ。この環境の DB スプレッドシートの、A・C 区分（と名前を書いた B 区分の 3 タブ）
    だけ。api・worker はこの関数を使わない（PURPOSE_TABS にも入れない）。"""
    s = _settings(settings)
    sid = (spreadsheet_id or "").strip()
    if sid in never_writable_ids(s):
        raise UnsafeWriteError("マスタ元・蓄積台帳・PDF フォルダへは書き込めません")
    if s.is_staging and sid in _config.PRODUCTION_IDS:
        raise UnsafeWriteError("ステージングから本番のスプレッドシートへは書き込めません")
    if not s.is_staging and sid in _config.KNOWN_STAGING_IDS:
        raise UnsafeWriteError("本番からステージングのスプレッドシートへは書き込めません")
    if not sid or sid != writable_spreadsheet_id(s):
        raise UnsafeWriteError("戻し先はこの環境の DB スプレッドシートだけです")
    if tab in RESTORE_TABS:
        return sid
    if tab in RESTORE_PROTECTED_TABS:
        if not protected_listed:
            raise UnsafeWriteError(f"タブ「{tab}」は --tabs に名前を書いたときだけ戻せます")
        return sid
    raise UnsafeWriteError(f"タブ「{tab}」は戻しの道具で書けません")


def assert_write_method(method: str, tab: str) -> None:
    """書き方の許可リスト。values.batchClear は C 区分だけ、appendDimension は A 区分だけ。"""
    if method not in WRITE_METHODS:
        raise UnsafeWriteError(f"使わない書き方です: {method}")
    if method == "values.batchClear" and tab not in C_TABS:
        raise UnsafeWriteError("値の消去はマスタのタブの余りの行だけに使えます")
    if method == "appendDimension" and tab not in A_TABS:
        raise UnsafeWriteError("行の追加は Requests・WorkLogs・BatchLogs の末尾だけに使えます")


def assert_write_requests(requests: Iterable[Mapping[str, Any]], tab: str, sheet_id: int) -> None:
    """spreadsheets.batchUpdate の要求を確かめる。updateCells と appendDimension（ROWS）だけ、すべて tab のシート。"""
    items = list(requests)
    if not items:
        raise UnsafeWriteError("書き込みの要求が空です")
    for req in items:
        if not isinstance(req, Mapping) or len(req) != 1:
            raise UnsafeWriteError("書き込みの要求の形が正しくありません")
        kind, body = next(iter(req.items()))
        if kind not in ALLOWED_BATCH_REQUESTS:
            raise UnsafeWriteError(f"使わない書き方です: {kind}")
        assert_write_method(kind, tab)
        if not isinstance(body, Mapping):
            raise UnsafeWriteError("書き込みの要求の形が正しくありません")
        if kind == "appendDimension":
            if body.get("dimension") != "ROWS":
                raise UnsafeWriteError("列は足せません（行だけ）")
            target = body.get("sheetId")
        else:
            area = body.get("range") or body.get("start") or {}
            target = area.get("sheetId") if isinstance(area, Mapping) else None
        if target is None or int(target) != int(sheet_id):
            raise UnsafeWriteError(f"要求の sheetId がタブ「{tab}」と違います")


def a1_sheet_name(a1: str) -> str:
    """A1 表記の範囲からタブ名を取り出す（'タブ'!A1:B2 → タブ。! が無ければ範囲全体がタブ名）。"""
    text = (a1 or "").strip()
    if text.startswith("'"):
        out, i = [], 1
        while i < len(text):
            ch = text[i]
            if ch == "'":
                if i + 1 < len(text) and text[i + 1] == "'":
                    out.append("'")
                    i += 2
                    continue
                return "".join(out)
            out.append(ch)
            i += 1
        raise UnsafeWriteError(f"範囲の書き方が正しくありません: {a1}")
    return text.split("!", 1)[0]


def assert_ranges_in_tab(ranges: Iterable[str], tab: str) -> None:
    for a1 in ranges:
        if a1_sheet_name(a1) != tab:
            raise UnsafeWriteError(f"範囲「{a1}」はタブ「{tab}」の外です")


def verify_startup(settings: Settings | None = None) -> None:
    """起動時の確認（api と worker）。書き込み先が許可リストを通らなければ UnsafeWriteError で起動しない。"""
    s = _settings(settings)
    target = writable_spreadsheet_id(s)
    for purpose, tabs in PURPOSE_TABS.items():
        for tab in tabs:
            assert_writable(target, tab, purpose, s)
    log.info("safety: APP_ENV=%s、書き込み先は DB スプレッドシートの A・C 区分のタブだけ", s.app_env)


# ===== GAS の内部用の入口の照合（実装仕様 9-1 の表の最後の行） =====


def verify_gas_ping(ping: Mapping[str, Any] | None, settings: Settings | None = None) -> list[str]:
    """internal.ping の返事を、この環境の許可リストと照合する。合わない点の一覧（空なら合格）。"""
    s = _settings(settings)
    problems: list[str] = []
    if not isinstance(ping, Mapping):
        return ["internal.ping の返事が JSON のオブジェクトではありません"]
    if ping.get("ok") is not True:
        problems.append(f"internal.ping が失敗しました: {ping.get('error') or ping.get('message') or '理由不明'}")
        return problems
    gas_settings = ping.get("settings") if isinstance(ping.get("settings"), Mapping) else {}
    sheet_id = str(ping.get("spreadsheetId") or "")
    script_id = str(ping.get("scriptId") or "")
    template = str(gas_settings.get("TEMPLATE_SSID") or "").strip()
    pdf_root = str(gas_settings.get("PDF_ROOT_FOLDER_ID") or "").strip()
    ledger = str(gas_settings.get("ACCUMULATION_SS_ID") or "").strip()
    tz = str(ping.get("spreadsheetTimeZone") or "")

    if s.is_staging:
        expected = {
            "spreadsheetId": (sheet_id, s.staging_spreadsheet_id),
            "TEMPLATE_SSID": (template, s.staging_spreadsheet_id),
            "PDF_ROOT_FOLDER_ID": (pdf_root, s.staging_pdf_folder_id),
            "ACCUMULATION_SS_ID": (ledger, s.staging_ledger_spreadsheet_id),
        }
        if script_id == _config.PROD_GAS_SCRIPT_ID_FIXED or not script_id:
            problems.append("scriptId が本番の GAS です（ステージングは GAS の写しを呼ぶ）")
        for name, (actual, _want) in expected.items():
            if actual in _config.PRODUCTION_IDS:
                problems.append(f"{name} が本番の ID を指しています")
    else:
        expected = {
            "spreadsheetId": (sheet_id, s.prod_spreadsheet_id),
            "TEMPLATE_SSID": (template, s.prod_spreadsheet_id),
            "PDF_ROOT_FOLDER_ID": (pdf_root, s.prod_pdf_folder_id),
            "ACCUMULATION_SS_ID": (ledger, s.prod_ledger_spreadsheet_id),
        }
        if script_id != _config.PROD_GAS_SCRIPT_ID_FIXED:
            problems.append("scriptId が本番の GAS と違います")
        for name, (actual, _want) in expected.items():
            if actual in _config.KNOWN_STAGING_IDS:
                problems.append(f"{name} がステージングの ID を指しています")
    for name, (actual, want) in expected.items():
        if actual != want:
            problems.append(f"{name} が {s.app_env} の値と違います")
    # Apps Script の getSpreadsheetTimeZone() は、シートの時間帯が GMT（Sheets API では "Etc/GMT"）のとき
    # 空の文字を返す（2026-11-26 ステージングで確認）。GMT の系統を設定しているときだけ、空を同じ値として扱う。
    # シートの本当の時間帯は、行の反映（row_sync）が Sheets API の properties.timeZone で別に確かめる。
    gmt_family = {"Etc/GMT", "GMT", "Etc/UTC", "UTC", "Etc/GMT+0", "Etc/GMT-0", "Etc/Greenwich", "Etc/Universal"}
    if tz == "" and s.db_sheet_timezone in gmt_family:
        tz = s.db_sheet_timezone
    if tz != s.db_sheet_timezone:
        problems.append(f"spreadsheetTimeZone（{tz or '空'}）が DB_SHEET_TIMEZONE（{s.db_sheet_timezone}）と違います")
    # 同じ指摘を重ねない（並びは保つ）
    return list(dict.fromkeys(problems))
