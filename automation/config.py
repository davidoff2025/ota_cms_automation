"""Load runtime config from config/.env (credentials never hard-coded)."""
from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv

_PROJECT_ROOT = Path(__file__).resolve().parents[1]
load_dotenv(_PROJECT_ROOT / "config" / ".env")


def _req(name: str) -> str:
    val = os.getenv(name, "").strip()
    if not val or val == "__set_me__":
        raise SystemExit(
            f"缺少配置 {name}。请在 config/.env 中填写（可从 config/.env.example 复制）。"
        )
    return val


def get(name: str, default: str = "") -> str:
    return os.getenv(name, default).strip()


# URLs (have sane defaults so only credentials are mandatory).
CMS_LOGIN_URL = get("CMS_LOGIN_URL", "http://cms.intsmarthub.com:31597/microop/user/login/")
CMS_OTA_URL = get("CMS_OTA_URL", "http://cms.intsmarthub.com:31597/basicservice/otamanage/upgrade-ota")

HEADED = get("CMS_HEADED", "true").lower() in ("1", "true", "yes", "on")
ARTIFACTS_DIR = _PROJECT_ROOT / get("CMS_ARTIFACTS_DIR", "./artifacts").lstrip("./")


def cms_username() -> str:
    return _req("CMS_USERNAME")


def cms_password() -> str:
    return _req("CMS_PASSWORD")


# DASS (FineReport fill-in site).
DASS_LOGIN_URL = get(
    "DASS_LOGIN_URL",
    "http://52.24.159.106:8080/dass/ReportServer?op=fs_load&cmd=fs_signin",
)
DASS_URL = get("DASS_URL", "http://52.24.159.106:8080/dass/ReportServer?op=fs")


def dass_username() -> str:
    return _req("DASS_USERNAME")


def dass_password() -> str:
    return _req("DASS_PASSWORD")
