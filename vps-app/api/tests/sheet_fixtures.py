"""担当 D の試験で使う、本番 DB スプレッドシートとマスタ元の形をした偽のブック。

見出しの文字・並び・見出しの行・値の型と表示形式の形は本番の控え（2026-11-26）に合わせた。
値はすべて作り物（個人のデータは入れない）。
"""
from __future__ import annotations

from app.config import MASTER_SOURCE_SPREADSHEET_ID_FIXED, PROD_SPREADSHEET_ID_FIXED
from conftest import TEST_ENV
from fake_sheets import DATE_FMT, DATETIME_FMT, TEXT_FMT, FakeSheets, serial

STAGING_SID = TEST_ENV["STAGING_SPREADSHEET_ID"]

# 本番の Requests の見出し（39 列。末尾に後から足された列が並ぶ）
REQUEST_HEADERS = [
    "requestId", "requestType(overtime/holiday)", "status(submitted/approved/canceled)", "dept", "workerCode",
    "workerName", "workerEmail", "targetDate", "submittedAt", "approvedAt", "approvedBy", "approvedMinutes", "reason",
    "workContent", "jobId1", "jobId2", "jobId3", "workNo1", "workNo2", "workNo3", "orderNo1", "orderNo2", "orderNo3",
    "customer1", "customer2", "customer3", "product1", "product2", "product3", "hrMailSentAt", "pdfGeneratedAt",
    "pdfFileId", "pdfFolderId", "exportError", "reasonDetail", "approvedBy2", "approvedAt2", "modifiedAt",
    "modifiedBy",
]
WORKLOG_NOTE = "※実績。残業はactualEndAtのみ入力（開始は17:20固定扱いで計算）。"
WORKLOG_HEADERS = ["requestId", "actualStartAt", "actualEndAt", "actualMinutes", "breakMinutes", "netMinutes",
                   "updatedAt", "updatedBy"]
BATCH_HEADERS = ["実行日時", "バッチ名", "対象日", "成功", "スキップ", "失敗", "エラー詳細"]

R1 = "11111111-1111-4111-8111-111111111111"
R2 = "22222222-2222-4222-8222-222222222222"
R3 = "33333333-3333-4333-8333-333333333333"
R4 = "44444444-4444-4444-8444-444444444444"
ORPHAN = "99999999-9999-4999-8999-999999999999"


def _req(**kw) -> list:
    row = [""] * len(REQUEST_HEADERS)
    for k, v in kw.items():
        row[REQUEST_HEADERS.index(k)] = v
    return row


def _fmt_cols(row_index: int, headers: list, names: tuple, fmt: dict) -> dict:
    return {(row_index, headers.index(n)): fmt for n in names}


def db_book_tabs(book) -> None:
    """DB スプレッドシートのタブ（A・B・C・D 区分）を足す。"""
    req_rows = [
        REQUEST_HEADERS,
        _req(requestId=R1, **{"requestType(overtime/holiday)": "overtime",
                              "status(submitted/approved/canceled)": "approved"},
             dept="機械加工", workerCode=1001, workerName="試験 一郎", targetDate=serial("2026-03-02"),
             submittedAt=serial("2026-03-02 08:15:30.250"), approvedAt=serial("2026-03-02 09:00:00"),
             approvedBy="approver-a@lineworks-local.info", approvedMinutes=60, reason="急ぎ作業を要するため",
             workContent="加工 1/2", workNo1="6401", orderNo1="LW00001", customer1="架空工業", product1="試作機",
             pdfGeneratedAt=serial("2026-03-02 18:40:01.500"), pdfFileId="pdf-file-1", pdfFolderId="pdf-folder-1",
             approvedBy2="admin-a@lineworks-local.info", approvedAt2=serial("2026-03-03 10:00:00")),
        _req(requestId=R2, **{"requestType(overtime/holiday)": "holiday",
                              "status(submitted/approved/canceled)": "approved"},
             dept="組立・塗装", workerCode="1002", workerName="試験 二郎", targetDate="2026-03-07",
             submittedAt="2026-03-05 12:00:00", approvedAt=serial("2026-03-05 13:00:00"),
             approvedBy="approver-b@lineworks-local.info", approvedMinutes=480, reason="その他（→補足理由必須）",
             reasonDetail="設備の据付", modifiedAt=serial("2026-03-06 09:30:00"),
             modifiedBy="approver-b@lineworks-local.info"),
        _req(requestId=R3, **{"requestType(overtime/holiday)": "overtime",
                              "status(submitted/approved/canceled)": "canceled"},
             dept="機械加工", workerCode=1003, workerName="試験 三郎", targetDate=serial("2026-03-09"),
             submittedAt=serial("2026-03-10 07:00:00"), approvedBy="approver-a@lineworks-local.info(却下)",
             approvedMinutes=30, reason="業者対応のため", modifiedAt="2026-03-10 08:00:00",
             modifiedBy="approver-a@lineworks-local.info"),
        _req(requestId=R4, **{"requestType(overtime/holiday)": "overtime",
                              "status(submitted/approved/canceled)": "submitted"},
             dept="機械加工", workerCode=1001, workerName="試験 一郎", targetDate=serial("2026-03-11"),
             submittedAt=serial("2026-03-11 06:30:00"), approvedMinutes=90, reason="顧客対応のため"),
    ]
    req_formats = {}
    for r in range(1, 5):
        req_formats.update(_fmt_cols(r, REQUEST_HEADERS, ("submittedAt", "approvedAt", "pdfGeneratedAt", "approvedAt2",
                                                          "modifiedAt"), DATETIME_FMT))
        req_formats.update(_fmt_cols(r, REQUEST_HEADERS, ("targetDate",), DATE_FMT))
    book.add_tab("Requests", req_rows, formats=req_formats, col_count=len(REQUEST_HEADERS))

    wl_rows = [
        [WORKLOG_NOTE, "", "", "", "", "", "", ""],
        WORKLOG_HEADERS,
        [R1, "2026-03-02 17:20:00", "2026-03-02 18:30:00", 70, 10, 60, "2026-03-02 18:30:05",
         "machine-01@lineworks-local.info"],
        [ORPHAN, serial("2026-02-10 08:00:00"), serial("2026-02-10 17:00:00"), 540, 60, 480,
         serial("2026-02-10 17:00:00"), "unknown"],
        [R2, "2026-03-07 08:00:00", "2026-03-07 17:00:00", 540, 60, 480, "2026-03-07 17:00:00", "sebastian-agent"],
        [R3, "", "", "", "", "", "2026-03-10 08:00:00", "delete"],
    ]
    wl_formats = {}
    for r in (2, 4, 5):
        wl_formats.update(_fmt_cols(r, WORKLOG_HEADERS, ("actualStartAt", "actualEndAt", "updatedAt"), TEXT_FMT))
    wl_formats.update(_fmt_cols(3, WORKLOG_HEADERS, ("actualStartAt", "actualEndAt", "updatedAt"), DATETIME_FMT))
    book.add_tab("WorkLogs", wl_rows, formats=wl_formats)

    bl_rows = [
        BATCH_HEADERS,
        [serial("2026-03-02 22:11:27.120"), "morningBatchNoMail", serial("2026-03-03"), 2, 0, 0, ""],
        [serial("2026-03-03 22:11:21"), "morningBatchNoMail", "2026-03-04", 1, 1, 1, "x: エラー\ny: エラー"],
    ]
    bl_formats = {(1, 0): DATETIME_FMT, (2, 0): DATETIME_FMT, (1, 2): DATE_FMT}
    book.add_tab("BatchLogs", bl_rows, formats=bl_formats)

    book.add_tab("Settings", [
        ["※総務宛先/テンプレID/保存先フォルダIDなどをここで管理", "", ""],
        ["key", "メールアドレス", "memo"],
        ["ADMIN_EMAILS", "old-admin@lineworks-local.info", "上の行（負ける）"],
        ["SOURCE_SSID", MASTER_SOURCE_SPREADSHEET_ID_FIXED, ""],
        ["GENERAL_AFFAIRS_EMAILS  ", "soumu@lineworks-local.info", "後ろに空白 2 つ"],
        ["TEMPLATE_SSID", STAGING_SID, ""],
        ["PDF_ROOT_FOLDER_ID", TEST_ENV["STAGING_PDF_FOLDER_ID"], ""],
        [" STAMP_HOLIDAY_FILE_ID  ", "stamp-file", "前 1 つ後ろ 2 つ"],
        ["ACCUMULATION_SS_ID", "wrong-ledger-id", "食い違い"],
        ["PDF_MODE", "direct で直接書込方式（未設定=XLOOKUP方式）", ""],
        ["ADMIN_EMAILS", "owner@lineworks-local.info,admin-a@lineworks-local.info", "下の行（勝つ）"],
    ])
    book.add_tab("DeptApprovers", [
        ["dept", "approverEmails", " approverName", "役職", "通知メール"],
        ["機械加工", "approver-a@lineworks-local.info", "承認 A", "主任", "notify-a@example.co.jp"],
        ["組立・塗装", "approver-b@lineworks-local.info,approver-c@lineworks-local.info", "承認 B", "係長",
         "notify-b@example.co.jp"],
    ])
    book.add_tab("ApproverMap", [
        ["※部署×Googleアカウントで承認画面の入室可否と承認権限を管理。", "", "", ""],
        ["部署", "承認者メール", "role(approver/admin)", "有効フラグ"],
    ])
    book.add_tab("StampMap", [["メール", "stampFileId", "備考"], ["approver-a@lineworks-local.info", "stamp-a", "A"]])
    book.add_tab("休憩マスタ", [["※実績計算用。", "", "", "", ""],
                               ["適用区分(overtime/holiday)", "実働分_下限", "実働分_上限", "休憩分", "備考"],
                               ["holiday", 0, 240, 0, "例"]])
    book.add_tab("申請書フォーム", [["テンプレ", "=Requests!A2"]])
    book.add_tab("操作", [["", ""], ["", ""], ["", R1]])
    book.add_tab("作業員マスタ", [["作業員コード", "氏名", "部署"], [9999, "古い 人", "旧部署"],
                                 [9998, "古い 人2", "旧部署"]] + [[9000 + i, f"古い{i}", "旧"] for i in range(100)])
    book.add_tab("業務NOマスタ", [["業務ＩＤ", "業務NO"], ["old", 1]])
    book.add_tab("工番マスタ", [["工番", "受注先"], ["OLD", "旧"]])
    book.add_tab("社内カレンダーマスタ", [["日付", "区分", "曜日", "備考"]])


def source_book_tabs(book, workers: int = 85) -> None:
    """マスタ元の 4 タブ（作業員は 1 列目が空の行・書式だけの列を含む）。"""
    w_rows = [["作業員コード", "氏名", "部署", "担当業務", "拠点", "スタッフ種類", ""]]
    for i in range(workers):
        w_rows.append([1001 + i, f"作業 {i:03d}", "機械加工" if i % 2 else "組立・塗装", "", "本社工場", "工場",
                       "書式だけの列の値" if i == 0 else ""])
        if i == 2:
            w_rows.append(["", "1 列目が空の行", "x", "", "", "", ""])
        if i == 3:
            w_rows.append([0, "1 列目が 0 の行", "x", "", "", "", ""])
    book.add_tab("作業員マスタ", w_rows)
    book.add_tab("業務NO.マスタ", [["業務ＩＤ", "業務NO", "業務名", "部署", "説明"],
                                ["bdcf58ab", 6401, "材料切断", "製缶・溶接", ""],
                                ["4e1bbc11", "0042", "仮組", "製缶・溶接", ""]])
    book.add_tab("工番マスタ", [["工番", "受注先", "納入先", "納入先住所", "品名", "数量", "取込日時"],
                             ["LW00001", "架空商事", "架空工業", "", "試作機", 1, serial("2025-09-01")],
                             ["LW00002", "架空商事", "架空工業", "", "量産機", 2, serial("2025-09-01 13:45:00")],
                             ["LW00003", "架空商事", "", "", "部品", 1, "9/1/2025"]],
                  formats={(1, 6): DATE_FMT, (2, 6): DATETIME_FMT})
    book.add_tab("社内カレンダーマスタ", [["日付", "区分", "曜日", "備考", ""],
                                    [serial("2026-04-04"), "休日", "Sat", "", ""],
                                    [serial("2026-04-05"), "休日", "Sun", "", ""],
                                    [serial("2026-04-29"), "祝日", "Wed", "昭和の日", ""],
                                    [serial("2026-04-11"), "出勤土曜", "Sat", "", ""]],
                  formats={(r, 0): DATE_FMT for r in range(1, 5)})


def make_fake(*, source_tz: str = "Asia/Tokyo", workers: int = 85) -> FakeSheets:
    """本番 DB（読む元）・ステージングの DB（書き込み先。同じブック）・マスタ元をそろえた偽物。"""
    fake = FakeSheets()
    book = fake.add_book(PROD_SPREADSHEET_ID_FIXED)
    db_book_tabs(book)
    fake.books[STAGING_SID] = book
    source = fake.add_book(MASTER_SOURCE_SPREADSHEET_ID_FIXED, tz=source_tz)
    source_book_tabs(source, workers=workers)
    return fake


def load_migrate_tool():
    """vps-app/tools/migrate_from_sheets.py を読み込む（tools はイメージに入らない。無ければ試験を飛ばす）。"""
    import importlib.util
    from pathlib import Path

    import pytest

    path = Path(__file__).resolve().parents[2] / "tools" / "migrate_from_sheets.py"
    if not path.exists():
        pytest.skip("tools/migrate_from_sheets.py がありません（VPS の試験は tools を付けて流す）")
    import sys

    if "migrate_from_sheets" in sys.modules:
        return sys.modules["migrate_from_sheets"]
    spec = importlib.util.spec_from_file_location("migrate_from_sheets", path)
    module = importlib.util.module_from_spec(spec)
    sys.modules["migrate_from_sheets"] = module
    spec.loader.exec_module(module)
    return module


def migrate(fake, engine, settings, **kw):
    """偽の本番 DB を移行の道具で DB に入れる（試験の準備）。報告を返す。"""
    tool = load_migrate_tool()
    reads = tool.read_sheets(fake.api(settings), PROD_SPREADSHEET_ID_FIXED)
    return tool.run(reads, engine, dry_run=False, **kw)
