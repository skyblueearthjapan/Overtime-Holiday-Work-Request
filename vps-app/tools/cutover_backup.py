"""切り替え前の控え: この環境の DB スプレッドシートと蓄積台帳を、Drive API の files.copy で日付つきの複製にする。

使い方（VPS。tools はイメージに入っていないので付けて動かす＝罠 5。docs/cutover-runbook.md の B-8）:
  docker compose run --rm --no-deps -T -e PYTHONPATH=/app -v /opt/overtime/tools:/app/tools:ro api \
      python tools/cutover_backup.py [--report-out /tmp/backup.json]

安全の決まり:
  - 複製する元は許可リストの 2 つだけ（APP_ENV=production なら本番の DB と台帳、staging ならステージングの DB と台帳）。
    ID を引数で受け取らない。どの呼び出しの前にも、元が許可リストにあるかを確かめる（違えば止める）。
  - 書くのは「複製を作る」だけ。元のファイルは変えない。複製はマイドライブの直下に置き、共有しない。
  - 複製の後、タブごとの A 列の行数を元と複製で比べる（違えばその冊は NG）。
  - 出すのは ID・名前・行数だけ（セルの中身・トークンは出さない）。
  - 委任トークン（DELEGATE_EMAIL）を使う。drive の許可が要る（google_delegate.DELEGATE_SCOPES）。
終了コード: 0＝すべて OK、1＝行数が合わない冊がある、2＝止めた（許可リスト外・API の失敗）。
"""
from __future__ import annotations

import argparse
import json
import sys
import urllib.parse
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Optional

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "api"))

from app import config as _config  # noqa: E402

DRIVE_API = "https://www.googleapis.com/drive/v3/files"
SHEETS_API = "https://sheets.googleapis.com/v4/spreadsheets"
BACKUP_PREFIX = "【切替前の控え】残業・休日出勤申請"


class BackupStop(RuntimeError):
    """控えを止める理由。"""


def allowed_sources(settings: Any) -> list[tuple[str, str]]:
    """この環境で複製してよい元（ラベル, ID）。"""
    if settings.is_staging:
        pairs = [("DB", settings.staging_spreadsheet_id), ("台帳", settings.staging_ledger_spreadsheet_id)]
        for _, fid in pairs:
            if not fid or fid in _config.PRODUCTION_IDS:
                raise BackupStop("ステージングの複製の元が本番の ID か空です")
        return pairs
    pairs = [("DB", _config.PROD_SPREADSHEET_ID_FIXED), ("台帳", _config.PROD_LEDGER_SPREADSHEET_ID_FIXED)]
    if settings.prod_spreadsheet_id != pairs[0][1] or settings.prod_ledger_spreadsheet_id != pairs[1][1]:
        raise BackupStop("本番の DB・台帳の ID が定数と違います")
    return pairs


def check_source(file_id: str, settings: Any) -> None:
    if file_id not in {fid for _, fid in allowed_sources(settings)}:
        raise BackupStop(f"許可リストに無いファイルです: {file_id}")


class Http:
    """Bearer トークンつきの JSON の呼び出し（httpx。試験では transport を差し替える）。"""

    def __init__(self, token_provider: Callable[[], str], transport: Any = None, timeout: float = 120.0):
        import httpx

        self._token_provider = token_provider
        self._client = httpx.Client(transport=transport, timeout=timeout)

    def call(self, method: str, url: str, *, params: Any = None, body: Any = None) -> dict:
        headers = {"Authorization": "Bearer " + self._token_provider()}
        resp = self._client.request(method, url, params=params, json=body, headers=headers)
        if resp.status_code >= 400:
            raise BackupStop(f"{method} {url.split('?')[0]} が失敗しました（{resp.status_code}）")
        return resp.json() if resp.content else {}


def tab_row_counts(http: Http, spreadsheet_id: str) -> dict[str, int]:
    """タブごとの A 列の行数（値のある最後の行まで）。"""
    meta = http.call("GET", f"{SHEETS_API}/{spreadsheet_id}", params={"fields": "sheets.properties.title"})
    titles = [s["properties"]["title"] for s in meta.get("sheets", [])]
    if not titles:
        return {}
    params = [("ranges", "'" + t.replace("'", "''") + "'!A:A") for t in titles] + [("majorDimension", "COLUMNS")]
    data = http.call("GET", f"{SHEETS_API}/{spreadsheet_id}/values:batchGet", params=params)
    ranges = data.get("valueRanges", [])
    out = {}
    for i, t in enumerate(titles):
        values = (ranges[i].get("values") if i < len(ranges) else None) or [[]]
        out[t] = len(values[0])
    return out


def copy_file(http: Http, file_id: str, name: str, settings: Any) -> dict:
    check_source(file_id, settings)  # 呼ぶ直前にもう一度
    return http.call("POST", f"{DRIVE_API}/{urllib.parse.quote(file_id)}/copy",
                     params={"supportsAllDrives": "true", "fields": "id,name"},
                     body={"name": name, "parents": ["root"]})


def run(http: Http, settings: Any, *, now: Optional[datetime] = None) -> dict:
    stamp = (now or datetime.now()).strftime("%Y-%m-%d %H%M")
    report: dict = {"appEnv": settings.app_env, "backups": [], "ok": True}
    for label, fid in allowed_sources(settings):
        check_source(fid, settings)
        before = tab_row_counts(http, fid)
        copied = copy_file(http, fid, f"{BACKUP_PREFIX} {label} {stamp}", settings)
        after = tab_row_counts(http, copied["id"])
        tabs = {t: {"source": n, "copy": after.get(t), "ok": after.get(t) == n} for t, n in before.items()}
        ok = bool(tabs) and all(v["ok"] for v in tabs.values())
        report["backups"].append({"label": label, "sourceId": fid, "copyId": copied["id"], "name": copied.get("name"),
                                  "tabs": tabs, "ok": ok})
        report["ok"] = report["ok"] and ok
    return report


def main(argv: Optional[list] = None, *, http: Optional[Http] = None, settings: Any = None) -> int:
    parser = argparse.ArgumentParser(description="切り替え前の控え（DB と台帳を Drive で複製する）")
    parser.add_argument("--report-out", help="報告の JSON の書き出し先")
    args = parser.parse_args(argv)
    try:
        if settings is None:
            from app.config import get_settings

            settings = get_settings()
        allowed_sources(settings)  # トークンを取る前に止める
        if http is None:
            from app.google_delegate import get_access_token

            token = get_access_token(settings=settings)
            http = Http(lambda: token)
        report = run(http, settings)
    except BackupStop as error:
        print(f"止めました: {error}", file=sys.stderr)
        return 2
    except Exception as error:  # noqa: BLE001 - 委任が無いなど
        print(f"止めました: {type(error).__name__}", file=sys.stderr)
        return 2
    text = json.dumps(report, ensure_ascii=False, indent=1)
    if args.report_out:
        Path(args.report_out).write_text(text, encoding="utf-8")
    print(text)
    for b in report["backups"]:
        print(f"{b['label']} 控えの ID: {b['copyId']}  {'OK' if b['ok'] else 'NG'}", file=sys.stderr)
    return 0 if report["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
