"""localeCompare(…, 'ja') の近似（ICU の日本語の照合。実装仕様 8-4・gap-1 G11）。

使う所は 2 か所だけ: 総務の日次一覧の並び（dept + workerName）と、月次サマリーの部署の円グラフの並び。
PyICU は入れない。代わりに ICU の ja の決まりのうち、ここで出る文字に効くものを写す:

- 一次（文字の種類と順）: 空白・記号 < 数字 < ラテン文字（大文字小文字・全角半角を区別しない）< ほかの文字
  < かな（ひらがな・カタカナ・半角カナを区別しない。五十音順）< 漢字（JIS X 0208 の順＝EUC-JP の符号順）
  < JIS X 0208 に無い漢字（符号位置の順）。
- 二次: 濁点・半濁点（清音 < 濁音 < 半濁音）。
- 三次: 小書き < 大きい字、ひらがな < カタカナ < 半角カナ、ラテンは小文字 < 大文字、半角 < 全角。
前の段で差が出なければ次の段を比べる（ICU と同じ）。node（ICU 入り）で並べた結果との一致を試験で確かめる。
長音記号「ー」の細かい扱いなど、ここに無い決まりは写していない（受け入れる差）。
"""
from __future__ import annotations

import unicodedata

_SMALL_KANA = {
    "ぁ": "あ", "ぃ": "い", "ぅ": "う", "ぇ": "え", "ぉ": "お", "っ": "つ", "ゃ": "や", "ゅ": "ゆ", "ょ": "よ",
    "ゎ": "わ", "ゕ": "か", "ゖ": "け",
}
_SYMBOLS = " _-,;:!?.'\"()[]{}@*/\\&#%`^+<=>|~$"


def _char_key(ch: str) -> tuple[tuple, int, tuple]:
    folded = unicodedata.normalize("NFC", unicodedata.normalize("NFKC", ch))
    wide = 1 if folded != ch else 0  # 全角・半角の違い（三次）
    if len(folded) != 1:
        folded = folded[:1] or ch
    c = folded
    if c.isspace():
        return (0, 0, 0), 0, (0, wide)
    if c in _SYMBOLS:
        return (0, 1, _SYMBOLS.index(c)), 0, (0, wide)
    if c.isdigit() and c.isascii():
        return (1, int(c)), 0, (0, wide)
    if c.isascii() and c.isalpha():
        return (2, ord(c.lower())), 0, (1 if c.isupper() else 0, wide)
    code = ord(c)
    if 0x3041 <= code <= 0x30FA and c != "・":
        # カタカナ → ひらがな、半角カナ（NFKC 前の元の字）は三次で後ろ
        script = 0
        if 0x30A1 <= code <= 0x30FA:
            c = chr(code - 0x60) if code - 0x60 >= 0x3041 else c
            script = 1
        if 0xFF61 <= ord(ch[0]) <= 0xFF9F:
            script = 2
        small = 0 if c in _SMALL_KANA else 1
        c = _SMALL_KANA.get(c, c)
        decomposed = unicodedata.normalize("NFD", c)
        base = decomposed[0]
        voiced = 0
        if "゙" in decomposed:
            voiced = 1
        elif "゚" in decomposed:
            voiced = 2
        return (4, ord(base)), voiced, (small, script)
    try:
        encoded = c.encode("euc_jp")
    except UnicodeEncodeError:
        encoded = b""
    if len(encoded) == 2 and encoded[0] >= 0xB0 and unicodedata.category(c) == "Lo":
        return (5, encoded[0], encoded[1]), 0, (0, wide)
    if unicodedata.category(c) == "Lo" and 0x3400 <= code:
        return (6, code), 0, (0, wide)
    if unicodedata.category(c).startswith(("P", "S")):
        return (0, 2, code), 0, (0, wide)
    return (3, code), 0, (0, wide)


def _segments(text: str) -> list[str]:
    """半角カナと続く半角の濁点（ﾞ・ﾟ）は 1 字として扱う。"""
    out: list[str] = []
    for ch in text:
        if ch in "ﾞﾟ" and out and 0xFF61 <= ord(out[-1][-1]) <= 0xFF9D:
            out[-1] += ch
        else:
            out.append(ch)
    return out


def ja_key(text: str) -> tuple:
    keys = [_char_key(seg) for seg in _segments(text or "")]
    return (tuple(k[0] for k in keys), tuple(k[1] for k in keys), tuple(k[2] for k in keys))
