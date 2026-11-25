"""配備の設定の確認（nginx・compose・Dockerfile・バックアップ・.env.example・画面の CSP の決まり）。
ファイルを読むだけ。リポジトリの中で動かすときだけ（イメージの中では飛ばす）。

画面のファイル（app/templates/pages/*.html・web/static/css/mobile.css）は担当 U が作る。まだ無ければその試験は飛ばす。
"""
from __future__ import annotations

import os
import re
from pathlib import Path

import pytest

VPS_DIR = Path(os.environ.get("OVERTIME_VPS_DIR") or Path(__file__).resolve().parents[2])
WEB = VPS_DIR / "web"
API = VPS_DIR / "api"

pytestmark = pytest.mark.skipif(not (WEB / "nginx.conf").exists(), reason="リポジトリの外（イメージの中）では確かめない")


def read(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def csp() -> dict[str, str]:
    value = re.search(r'Content-Security-Policy "([^"]+)"', read(WEB / "nginx.conf")).group(1)
    return {d.split()[0]: " ".join(d.split()[1:]) for d in (x.strip() for x in value.split(";")) if d}


def test_csp_is_exactly_the_spec():
    assert csp() == {
        "default-src": "'self'", "script-src": "'self'", "style-src": "'self' 'unsafe-inline'",
        "img-src": "'self' data:", "connect-src": "'self'", "frame-ancestors": "'self'", "base-uri": "'self'",
        "form-action": "'self'",
    }


def test_nginx_headers_only_at_server_level_and_proxies():
    conf = read(WEB / "nginx.conf")
    for block in re.findall(r"location [^{]+\{([^}]*)\}", conf):
        assert "add_header" not in block  # location の add_header は server の分を消す
    for header in ("X-Content-Type-Options \"nosniff\"", "Strict-Transport-Security", "Referrer-Policy",
                   "Permissions-Policy"):
        assert conf.count(header) == 1
    assert "limit_req_zone $overtime_client_ip zone=overtime_api:10m rate=10r/s;" in conf
    assert "limit_req zone=overtime_api burst=30 nodelay;" in conf
    assert "$http_cf_connecting_ip" in conf
    assert "$proxy_add_x_forwarded_for" not in conf
    locations = dict(re.findall(r"^\s*location ([^{]+?)\s*\{([^}]*)\}", conf, flags=re.M))
    assert "proxy_pass" not in locations["/static/"] and "expires -1;" in locations["/static/"]
    for loc in ("/api/", "/auth/", "= /healthz", "/"):
        assert "proxy_pass http://overtime_api;" in locations[loc], loc
        assert "proxy_set_header X-Forwarded-For $overtime_client_ip;" in locations[loc]


def test_compose_names_network_and_no_host_ports():
    compose = read(VPS_DIR / "docker-compose.yml")
    assert re.search(r"^name: overtime$", compose, flags=re.M)
    names = re.findall(r"container_name: (\S+)", compose)
    assert sorted(names) == sorted(["overtime-db", "overtime-api", "overtime-worker", "overtime-web",
                                    "overtime-cloudflared", "overtime-backup"])
    assert "name: overtime-pgdata" in compose and "name: overtime-net" in compose
    assert "subnet: ${OVERTIME_NET_SUBNET:-10.93.0.0/24}" in compose
    assert "ipv4_address: ${OVERTIME_WEB_IP:-10.93.0.10}" in compose
    assert "FORWARDED_ALLOW_IPS: ${FORWARDED_ALLOW_IPS:-10.93.0.10}" in compose
    assert compose.count("image: overtime-app:${APP_IMAGE_TAG:-latest}") == 2  # api と worker は同じイメージ
    assert not re.search(r"^\s+ports:", compose, flags=re.M)  # ホストにポートを出さない
    image = re.search(r"cloudflare/cloudflared:(\S+)", compose).group(1)
    assert re.fullmatch(r"\d{4}\.\d+\.\d+", image)
    worker = compose.split("\n  worker:", 1)[1].split("\n  web:", 1)[0]
    assert 'command: ["python", "-m", "app.worker"]' in worker
    assert "disable: true" in worker and "condition: service_healthy" in worker
    assert "./web/static:/usr/share/nginx/html/static:ro" in compose


def test_compose_and_env_example_pass_every_setting_key():
    """config.Settings の値はすべて compose から api・worker に渡し、.env.example に書く。"""
    from app.config import Settings

    compose = read(VPS_DIR / "docker-compose.yml")
    env_example = read(VPS_DIR / ".env.example")
    for field in Settings.model_fields:
        key = field.upper()
        assert f"  {key}: " in compose, key
        if key != "DATABASE_URL":
            assert re.search(rf"^{key}=", env_example, flags=re.M), key
    for key in ("POSTGRES_USER", "POSTGRES_PASSWORD", "POSTGRES_DB", "APP_IMAGE_TAG", "TUNNEL_TOKEN",
                "BACKUPS_HOST_PATH", "BACKUP_KEEP_DAYS", "OVERTIME_NET_SUBNET", "OVERTIME_WEB_IP",
                "FORWARDED_ALLOW_IPS"):
        assert re.search(rf"^{key}=", env_example, flags=re.M), key


def test_env_example_has_no_secrets():
    env = dict(line.split("=", 1) for line in read(VPS_DIR / ".env.example").splitlines()
               if line and not line.startswith("#"))
    for key in ("POSTGRES_PASSWORD", "GOOGLE_OAUTH_CLIENT_SECRET", "SESSION_SECRET", "FERNET_KEY",
                "GAS_INTERNAL_TOKEN", "SEBASTIAN_API_TOKEN", "TUNNEL_TOKEN"):
        assert env[key] == "CHANGE_ME", key
    for key in ("MAIL_REDIRECT_TO", "MAIL_ALLOWED_DOMAINS", "EXPORT_HOURLY_UNTIL"):  # このアプリでは持たない
        assert key not in env
    assert env["SESSION_MAX_AGE_DAYS"] == "90" and env["SESSION_IDLE_HOURS"] == "2160"


def test_dockerfile_base_without_chromium_and_runtime_stage():
    dockerfile = read(API / "Dockerfile")
    assert "FROM python:3.12-slim-bookworm AS base" in dockerfile
    code = re.sub(r"#.*", "", dockerfile)
    for word in ("playwright", "chromium", "fonts-noto-cjk"):
        assert word not in code
    assert "--forwarded-allow-ips='*'" not in dockerfile
    stages = re.split(r"^FROM ", dockerfile, flags=re.M)
    assert stages[-1].startswith("base AS runtime") and "USER app" in stages[-1]
    base = next(s for s in stages if s.startswith("python:"))
    assert "tests" not in re.sub(r"#.*", "", base) and "requirements-dev" not in base
    reqs = re.sub(r"#.*", "", read(API / "requirements.txt"))
    for line in (x for x in reqs.splitlines() if x and not x.startswith("#")):
        assert "==" in line, line  # 版はすべて固定
    for word in ("playwright", "segno", "lxml"):
        assert word not in reqs
    assert re.search(r"^jinja2==", reqs, flags=re.M)


def test_backup_runner_uses_umask_077_and_overtime_prefix():
    script = read(VPS_DIR / "backup" / "backup_entrypoint.sh")
    runner = script.split('cat > "$RUNNER" <<EOF', 1)[1].split("\nEOF", 1)[0]
    assert "\numask 077\n" in runner
    assert "/backups/overtime_" in runner and "'overtime_*.dump'" in runner
    assert '-lt 1024' in runner and "FAILED" in runner
    assert "PGPASSFILE" in runner  # cron に環境変数が渡らなくても動く（罠 26）
    assert "0 3 * * *" in script


def test_runbook_uses_git_archive_not_git_pull():
    runbook = read(VPS_DIR / "docs" / "RUNBOOK.md")
    assert "git archive" in runbook and "/opt/overtime" in runbook
    assert "git pull" not in runbook and "/opt/overtime/vps-app" not in runbook
    assert "srv1508169" in runbook and "srv1300618" in runbook


def test_gitattributes_keeps_lf_for_vps_files():
    attrs = read(VPS_DIR / ".gitattributes")
    for pattern in ("*.sh text eol=lf", "*.conf text eol=lf", "*.py text eol=lf", "Dockerfile text eol=lf"):
        assert pattern in attrs


# ===== 画面のファイル（担当 U。あるときだけ） =====

PAGES = API / "app" / "templates" / "pages"
MOBILE_CSS = WEB / "static" / "css" / "mobile.css"
ON_ATTR = re.compile(r"<[^>]*\son[a-z]+\s*=", re.I)


@pytest.mark.skipif(not PAGES.exists(), reason="画面のテンプレートがまだ無い（担当 U）")
def test_pages_have_no_inline_script_or_on_attributes():
    pages = sorted(PAGES.glob("*.html"))
    assert pages
    for page in pages:
        html = read(page)
        for m in re.finditer(r"<script\b([^>]*)>(.*?)</script>", html, flags=re.S | re.I):
            attrs, body = m.group(1), m.group(2)
            if "src=" in attrs:
                assert not body.strip(), page.name
            else:
                assert 'type="application/json"' in attrs, page.name  # 実行されない JSON の塊だけ
        assert not ON_ATTR.search(html), page.name


@pytest.mark.skipif(not MOBILE_CSS.exists(), reason="mobile.css がまだ無い（担当 U）")
def test_mobile_css_rules_are_all_inside_max_width_820_media():
    css = re.sub(r"/\*.*?\*/", "", read(MOBILE_CSS), flags=re.S)
    depth, i, top_level = 0, 0, []
    while i < len(css):
        if css.startswith("@media", i) and depth == 0:
            end = css.index("{", i)
            query = css[i:end]
            widths = [int(w) for w in re.findall(r"max-width\s*:\s*(\d+)px", query)]
            assert widths and max(widths) <= 820, query
            assert "min-width" not in query, query
            depth, i = 1, end + 1
            continue
        ch = css[i]
        if ch == "{":
            if depth == 0:
                top_level.append(css[max(0, i - 40):i])
            depth += 1
        elif ch == "}":
            depth -= 1
        i += 1
    assert top_level == []  # @media の外に規則が無い（PC の見た目を変えない）
