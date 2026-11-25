"""services/gas_client.GasClient（GAS の内部用の入口）: 本文の形・時間切れ・200 以外・JSON でない・URL の許可リスト。"""
from __future__ import annotations

import json

import httpx
import pytest

from app.config import PROD_GAS_DEPLOYMENT_IDS_FIXED, STAGING_GAS_DEPLOYMENT_IDS_KNOWN
from app.services.gas_client import GasClient, GasClientConfigError, GasTransportError, assert_gas_url_allowed
from conftest import TEST_ENV, make_settings

URL = TEST_ENV["GAS_INTERNAL_URL"]
TOKEN = TEST_ENV["GAS_INTERNAL_TOKEN"]


def client_with(handler, settings=None, **kw):
    transport = httpx.MockTransport(handler)
    factory = lambda **opts: httpx.Client(transport=transport, **opts)  # noqa: E731
    if settings is not None:
        return GasClient.from_settings(settings, http_factory=factory, **kw)
    return GasClient(URL, TOKEN, http_factory=factory, **kw)


def test_call_sends_internal_body_and_follows_redirect(settings):
    seen = {}

    def handler(request: httpx.Request):
        if request.url.host == "script.google.com":
            seen["method"] = request.method
            seen["body"] = json.loads(request.content)
            seen["ct"] = request.headers["content-type"]
            return httpx.Response(302, headers={"location": "https://script.googleusercontent.com/macros/echo?x=1"})
        seen["followed"] = request.method
        return httpx.Response(200, json={"ok": True, "action": "internal.afterApprove", "jobId": "j-1",
                                         "pdf": {"status": "created", "pdfFileId": "F"}})

    gas = client_with(handler, settings)
    result = gas.call("internal.afterApprove", job_id="j-1", request_id="R1", expect={"status": "approved"},
                      args={"source": "screen"})
    assert result["pdf"]["status"] == "created"
    assert seen["method"] == "POST" and seen["followed"] == "GET" and seen["ct"] == "application/json"
    assert seen["body"] == {"internal": True, "token": TOKEN, "action": "internal.afterApprove", "jobId": "j-1",
                            "requestId": "R1", "expect": {"status": "approved"}, "args": {"source": "screen"}}
    assert gas.timeout == 330


def test_failure_json_is_returned_as_is():
    gas = client_with(lambda r: httpx.Response(200, json={"ok": False, "error": "row_not_synced", "retryable": True}))
    assert gas.call("internal.ping", job_id="p-1") == {"ok": False, "error": "row_not_synced", "retryable": True}


def test_timeout_is_transport_error():
    def handler(request):
        raise httpx.ReadTimeout("slow", request=request)

    with pytest.raises(GasTransportError) as exc:
        client_with(handler).call("internal.batchPdf", job_id="j", timeout=5)
    assert exc.value.timeout is True
    assert TOKEN not in str(exc.value)


def test_connect_error_is_transport_error():
    def handler(request):
        raise httpx.ConnectError("down", request=request)

    with pytest.raises(GasTransportError) as exc:
        client_with(handler).call("internal.ping", job_id="j")
    assert exc.value.timeout is False


@pytest.mark.parametrize("response", [httpx.Response(500, text="err"), httpx.Response(200, text="<html>"),
                                      httpx.Response(200, json=[1, 2])])
def test_bad_responses_are_transport_errors(response):
    with pytest.raises(GasTransportError):
        client_with(lambda r: response).call("internal.ping", job_id="j")


def test_arguments_are_checked():
    gas = client_with(lambda r: httpx.Response(200, json={"ok": True}))
    with pytest.raises(ValueError):
        gas.call("internal.unknown", job_id="j")
    with pytest.raises(ValueError):
        gas.call("internal.ping", job_id="")
    assert TOKEN not in repr(gas)


def test_url_allow_list(settings):
    with pytest.raises(GasClientConfigError):
        assert_gas_url_allowed(f"https://script.google.com/macros/s/{PROD_GAS_DEPLOYMENT_IDS_FIXED[0]}/exec", settings)
    prod = make_settings(app_env="production", staging_spreadsheet_id="", staging_pdf_folder_id="",
                         staging_ledger_spreadsheet_id="",
                         gas_internal_url="https://script.google.com/macros/s/AKfycbTESTprod/exec")
    with pytest.raises(GasClientConfigError):
        assert_gas_url_allowed(f"https://script.google.com/macros/s/{STAGING_GAS_DEPLOYMENT_IDS_KNOWN[0]}/exec", prod)
    with pytest.raises(GasClientConfigError):
        assert_gas_url_allowed("https://example.com/", prod)
    assert assert_gas_url_allowed(prod.gas_internal_url, prod) == "AKfycbTESTprod"
    with pytest.raises(GasClientConfigError):
        GasClient(URL, "", settings=settings)
