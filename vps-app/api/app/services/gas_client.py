"""GAS の内部用の入口を呼ぶ（決定 B・実装仕様 7-1）。担当 F の土台。担当 M（services/gas_jobs.py）が使う。

本文: {"internal": true, "token": GAS_INTERNAL_TOKEN, "action": "internal.<名前>", "jobId": "<uuid>",
       "requestId": "...", "expect": {...}, "args": {...}}
返事: いつも HTTP 200 の JSON（成功 {"ok": true, ...}、失敗 {"ok": false, "error": "<コード>", "retryable": ...}）。
GAS の /exec への POST は 302 で script.googleusercontent.com へ送られ、そこを GET で読むと JSON が返る
（httpx は 302 の後を GET で追う）。

- 通信の失敗・時間切れ・200 以外・JSON でない返事は GasTransportError（「結果が分からない」。
  担当 M は同じ jobId で送り直す。GAS は jobId で同じ返事を返す＝冪等）。
- token はログにも例外の文にも出さない。
- 呼ぶ前に URL が環境の許可リストに合うか確かめる（staging は本番の GAS を呼ばない）。
"""
from __future__ import annotations

import json
import logging
import time
from collections.abc import Callable, Mapping
from typing import Any

import httpx

from app import config as _config
from app.config import Settings

log = logging.getLogger(__name__)

ACTIONS = frozenset({
    "internal.ping",
    "internal.afterApprove",
    "internal.afterComplete",
    "internal.regenerate",
    "internal.ledgerUpsert",
    "internal.afterDelete",
    "internal.batchPdf",
    "internal.ledgerRewriteDay",
    "internal.ledgerRebuild",
})


class GasTransportError(RuntimeError):
    """GAS に届かなかった・返事を読めなかった（結果は分からない）。timeout=True は時間切れ。"""

    def __init__(self, message: str, *, timeout: bool = False, status: int | None = None):
        super().__init__(message)
        self.timeout = timeout
        self.status = status


class GasClientConfigError(RuntimeError):
    """URL が環境の許可リストに合わない（staging から本番の GAS を呼ぶなど）。"""


def assert_gas_url_allowed(url: str, settings: Settings) -> str:
    deployment = _config.gas_deployment_id(url)
    if not deployment:
        raise GasClientConfigError("GAS_INTERNAL_URL の形が正しくありません")
    if settings.is_staging and deployment in _config.PRODUCTION_IDS:
        raise GasClientConfigError("ステージングから本番の GAS は呼べません")
    if not settings.is_staging and deployment in _config.KNOWN_STAGING_IDS:
        raise GasClientConfigError("本番からステージングの GAS は呼べません")
    return deployment


class GasClient:
    """http_factory は httpx.Client を作る関数（試験では MockTransport を入れる）。"""

    def __init__(self, url: str, token: str, *, timeout: float = 330.0, settings: Settings | None = None,
                 http_factory: Callable[..., httpx.Client] | None = None, monotonic: Callable[[], float] = time.monotonic):
        if settings is not None:
            assert_gas_url_allowed(url, settings)
        if not token:
            raise GasClientConfigError("GAS_INTERNAL_TOKEN が空です")
        self.url = url
        self._token = token
        self.timeout = float(timeout)
        self._http_factory = http_factory or (lambda **kw: httpx.Client(**kw))
        self._monotonic = monotonic

    def __repr__(self) -> str:  # token を出さない
        return f"GasClient(url={self.url!r}, timeout={self.timeout})"

    @classmethod
    def from_settings(cls, settings: Settings, **kwargs: Any) -> "GasClient":
        return cls(settings.gas_internal_url, settings.gas_internal_token,
                   timeout=settings.gas_call_timeout_seconds, settings=settings, **kwargs)

    def call(self, action: str, *, job_id: str, request_id: str = "", expect: Mapping[str, Any] | None = None,
             args: Mapping[str, Any] | None = None, timeout: float | None = None) -> dict:
        """1 件呼んで返事の JSON（dict）を返す。通信の失敗・時間切れは GasTransportError。"""
        if action not in ACTIONS:
            raise ValueError(f"知らない action です: {action}")
        if not job_id:
            raise ValueError("job_id が空です（GAS は jobId で二重の処理を防ぐ）")
        body: dict[str, Any] = {"internal": True, "token": self._token, "action": action, "jobId": str(job_id),
                                "requestId": request_id or ""}
        if expect is not None:
            body["expect"] = dict(expect)
        if args is not None:
            body["args"] = dict(args)
        limit = float(timeout if timeout is not None else self.timeout)
        started = self._monotonic()
        try:
            with self._http_factory(timeout=httpx.Timeout(limit, connect=min(30.0, limit)),
                                    follow_redirects=True) as client:
                resp = client.post(self.url, content=json.dumps(body, ensure_ascii=False).encode("utf-8"),
                                   headers={"Content-Type": "application/json"})
        except httpx.TimeoutException as exc:
            raise GasTransportError(f"GAS の呼び出しが時間切れになりました（{action}、{limit:.0f} 秒）",
                                    timeout=True) from None
        except httpx.HTTPError as exc:
            raise GasTransportError(f"GAS に届きませんでした（{action}）: {type(exc).__name__}") from None
        elapsed = self._monotonic() - started
        if resp.status_code != 200:
            raise GasTransportError(f"GAS の返事が 200 ではありません（{action}）: {resp.status_code}",
                                    status=resp.status_code)
        try:
            data = resp.json()
        except ValueError:
            raise GasTransportError(f"GAS の返事が JSON ではありません（{action}）") from None
        if not isinstance(data, dict):
            raise GasTransportError(f"GAS の返事が JSON のオブジェクトではありません（{action}）")
        log.info("gas %s job=%s request=%s ok=%s %.1fs", action, job_id, request_id, data.get("ok"), elapsed)
        return data
