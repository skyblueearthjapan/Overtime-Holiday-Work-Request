"""GAS（JavaScript）と同じ結果になる小道具（実装仕様 8-3・8-4）。

- js_round(x) = Math.round（.5 は上へ。負の数も JS と同じ）。Python の組み込みの丸め（偶数への丸め）は使わない。
- js_floor(x) = Math.floor。
- js_string(v) = String(v)（1.0 → "1"、true → "true"、null/undefined → ""）。
- normalize_gas(s) = src/config.gs の normalize_（前後の空白を取り、全角英数を半角に）。
- js_sort_key(s) = JS の文字の比べ（UTF-16 の符号単位の順）と同じ並びになるキー。
"""
from __future__ import annotations

import math
import re
from typing import Any

# JS の String.prototype.trim が取る空白（WhiteSpace と LineTerminator）
_JS_WHITESPACE = (
    "\u0009\u000a\u000b\u000c\u000d          "
    "        　﻿"
)
# JS の正規表現の \s と同じ文字の組（Python の \s は U+001C・U+0085 を含み U+FEFF を含まない）
JS_WHITESPACE_RE = re.compile("[" + re.escape(_JS_WHITESPACE) + "]+")
_FULLWIDTH_ALNUM = re.compile(r"[Ａ-Ｚａ-ｚ０-９]")


def js_round(x: float) -> int:
    """Math.round と同じ（floor(x + 0.5)）。14.5 → 15、-14.5 → -14。"""
    return int(math.floor(x + 0.5))


def js_floor(x: float) -> int:
    return int(math.floor(x))


def _js_number_string(value: float) -> str:
    if math.isnan(value):
        return "NaN"
    if math.isinf(value):
        return "Infinity" if value > 0 else "-Infinity"
    if value == int(value) and abs(value) < 1e21:
        return str(int(value))
    text = repr(value)
    if "e" in text:
        mantissa, exp = text.split("e")
        exp_num = int(exp)
        return f"{mantissa}e{'+' if exp_num > 0 else '-'}{abs(exp_num)}"
    return text


def js_string(value: Any) -> str:
    """String(value) と同じ文字（null・undefined にあたる None は ''。GAS の `String(x ?? '')` の形）。"""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        return _js_number_string(value)
    if isinstance(value, (list, tuple)):
        # Array.prototype.toString: 要素を ',' でつなぐ（null・undefined は空。正しさ 低-2）
        return ",".join(js_string(v) for v in value)
    if isinstance(value, dict):
        return "[object Object]"
    return str(value)


def js_property_key(value: Any) -> str:
    """オブジェクトのキーにしたときの文字（`obj[value]`）。null は 'null'（String(null)）。"""
    if value is None:
        return "null"
    if isinstance(value, (list, tuple)):
        return ",".join("" if v is None else js_string(v) for v in value)
    return js_string(value)


def _is_array_index(key: str) -> bool:
    return key.isdigit() and key.isascii() and (key == "0" or not key.startswith("0")) and int(key) < 2**32 - 1


def js_object_key_order(keys: list[str]) -> list[str]:
    """`for (k in obj)` の順: 配列の添字に見える数字のキーを小さい順に先に、残りは入れた順。"""
    index_keys = sorted((k for k in keys if _is_array_index(k)), key=int)
    return index_keys + [k for k in keys if not _is_array_index(k)]


def js_utf16_units(text: str) -> list[str]:
    """文字列を JS の添字（UTF-16 の符号単位）ごとに分ける（サロゲートのペアは 2 つになる）。"""
    data = text.encode("utf-16-be", "surrogatepass")
    return [data[i:i + 2].decode("utf-16-be", "surrogatepass") for i in range(0, len(data), 2)]


def js_array_like_items(value: Any) -> list:
    """`for (k = 0; k < x.length; k++) x[k]` で読む値（配列はそのまま、文字列は 1 字ずつ、ほかは length が無いので空）。"""
    if isinstance(value, (list, tuple)):
        return list(value)
    if isinstance(value, str):
        return js_utf16_units(value)
    return []


def js_trim(text: str) -> str:
    return text.strip(_JS_WHITESPACE)


def normalize_gas(value: Any) -> str:
    """src/config.gs の normalize_ と同じ: String(s ?? '').trim() の後、全角英数（Ａ-Ｚａ-ｚ０-９）を半角に。"""
    text = js_trim(js_string(value))
    return _FULLWIDTH_ALNUM.sub(lambda m: chr(ord(m.group(0)) - 0xFEE0), text)


def js_sort_key(value: Any) -> tuple[int, ...]:
    """JS の既定の並び（Array.prototype.sort・`<`）と同じ順になるキー（UTF-16 の符号単位の並び）。"""
    data = js_string(value).encode("utf-16-be", "surrogatepass")
    return tuple(int.from_bytes(data[i:i + 2], "big") for i in range(0, len(data), 2))
