"""画面の静的な中身の試験（仕様 3-6・3-8・3-9・12-2 の U）。ファイルを読むだけで、app は起動しない。"""
from __future__ import annotations

import re
from pathlib import Path

import pytest

VPS = Path(__file__).resolve().parents[2]
TPL = VPS / "api" / "app" / "templates" / "pages"
JS = VPS / "web" / "static" / "js"
CSS = VPS / "web" / "static" / "css" / "mobile.css"

PAGES = ["top", "approver", "approve2", "admin", "staffSummary"]
ALL_TEMPLATES = PAGES + ["no_auth"]
ON_ATTR = re.compile(r"<[A-Za-z][^>]*\son[a-z]+\s*=", re.S)


def read(p: Path) -> str:
    return p.read_text(encoding="utf-8")


@pytest.mark.parametrize("page", ALL_TEMPLATES)
def test_templates_have_no_inline_handlers_or_scripts(page):
    html = read(TPL / f"{page}.html")
    assert not ON_ATTR.search(html), "on 属性が残っている"
    # 実行される <script> は外部ファイルだけ（JSON の塊は実行されない）
    for tag in re.findall(r"<script\b[^>]*>", html):
        assert ('src="/static/js/' in tag) or ('type="application/json"' in tag), tag
    for body in re.findall(r"<script\b(?![^>]*\bsrc=)(?![^>]*application/json)[^>]*>(.*?)</script>", html, re.S):
        assert body.strip() == ""
    assert "<?" not in html, "GAS のテンプレートの書き方が残っている"
    assert "javascript:" not in html


@pytest.mark.parametrize("page", ALL_TEMPLATES)
def test_titles_and_mobile_css(page):
    html = read(TPL / f"{page}.html")
    title = re.findall(r"<title>(.*?)</title>", html)
    assert title == (["権限エラー"] if page == "no_auth" else ["残業・休日出勤申請"])
    assert '<meta name="viewport" content="width=device-width, initial-scale=1"' in html
    # mobile.css は head の最後の link
    head = html.split("</head>")[0]
    assert head.rstrip().endswith('<link rel="stylesheet" href="/static/css/mobile.css">')


@pytest.mark.parametrize("page", PAGES)
def test_page_scripts_exist(page):
    html = read(TPL / f"{page}.html")
    srcs = re.findall(r'<script src="(/static/js/[^"]+)"></script>', html)
    assert srcs == ["/static/js/gas-shim.js", f"/static/js/{page}.js"]
    for s in srcs:
        assert (VPS / "web" / "static" / s[len("/static/"):]).is_file()


@pytest.mark.parametrize("page", PAGES)
def test_js_has_no_onclick_strings_or_onclick_selectors(page):
    js = read(JS / f"{page}.js")
    assert not re.search(r"\son(click|input|change|mouseenter|mouseleave)\s*=", js)
    assert "[onclick*=" not in js
    assert "<?" not in js


@pytest.mark.parametrize("page", ["top", "approver", "approve2", "admin"])
def test_eschtml_escapes_single_quote(page):
    js = read(JS / f"{page}.js")
    body = re.search(r"function escHtml\(s\)\{(.*?)\n\}", js, re.S).group(1)
    for frag in ("&amp;", "&lt;", "&gt;", "&quot;", "&#39;"):
        assert frag in body


def test_staff_summary_esc_escapes_single_quote():
    js = read(JS / "staffSummary.js")
    assert "/[&<>\"']/g" in js and "\"'\":'&#39;'" in js


@pytest.mark.parametrize("page", PAGES)
def test_no_safari_unsafe_date_strings(page):
    """`new Date('yyyy-MM-dd HH:mm:ss')` のような空白入りの文字を Date に渡さない（Safari で Invalid Date）。"""
    js = read(JS / f"{page}.js")
    for m in re.finditer(r"new Date\(([^()]*(?:\([^()]*\)[^()]*)*)\)", js):
        arg = m.group(1)
        lits = re.findall(r"'([^']*)'|\"([^\"]*)\"", arg)
        for a, b in lits:
            s = a or b
            assert not re.search(r"\d{4}-\d{2}-\d{2} \d", s), m.group(0)
            assert not re.match(r"^ \d", s), m.group(0)  # 'yyyy-MM-dd' + ' HH:mm' の形


def test_approver_finds_batch_button_by_id():
    js = read(JS / "approver.js")
    html = read(TPL / "approver.html")
    assert "document.getElementById('btnApproveAll-' + type)" in js
    assert 'id="btnApproveAll-overtime"' in html and 'id="btnApproveAll-holiday"' in html


def test_admin_manual_entry_button_uses_json_args():
    js = read(JS / "admin.js")
    assert "openManualEntry(\\'" not in js
    assert "data-on-click=\"openManualEntry\" data-args=\"'+__vpsArgs(it.requestId, it.workerName," in js


def test_mobile_css_rules_are_all_inside_narrow_media():
    css = re.sub(r"/\*.*?\*/", "", read(CSS), flags=re.S)
    depth = 0
    i = 0
    top_level_blocks = []
    while i < len(css):
        if css[i] == "{":
            if depth == 0:
                start = css.rfind("}", 0, i) + 1
                top_level_blocks.append(css[start:i].strip())
            depth += 1
        elif css[i] == "}":
            depth -= 1
        i += 1
    assert depth == 0
    assert top_level_blocks, "規則が無い"
    for sel in top_level_blocks:
        m = re.fullmatch(r"@media\s*\(max-width:\s*(\d+)px\)", sel)
        assert m and int(m.group(1)) <= 820, sel
    assert "-webkit-appearance: none" in css and "::-webkit-date-and-time-value" in css


def test_templates_are_generated_from_src():
    """templates と js は src/ から port-pages.js で作った物と同じ（手で直していない）。node が無ければ飛ばす。"""
    import shutil
    import subprocess

    node = shutil.which("node")
    if not node or not (VPS.parent / "src" / "top.html").is_file():
        pytest.skip("node か src/ が無い（イメージの中など）")
    r = subprocess.run([node, str(VPS / "web" / "dev" / "port-pages.js"), "--check"], capture_output=True, text=True, encoding="utf-8")
    assert r.returncode == 0, r.stderr
