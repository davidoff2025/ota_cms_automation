"""Runtime configuration — 12-factor style so the same image runs locally and on AWS.

Every deployment-specific value comes from an environment variable with a sane
local default. On AWS you override these via the task/instance environment (or
Secrets Manager / SSM Parameter Store injected as env vars); nothing here needs
code changes to move from 127.0.0.1 to a server.
"""
from __future__ import annotations

import os
from pathlib import Path

# Load a local .env (if present) BEFORE reading any env vars, so secrets like the
# FTP password can live in an untracked file instead of being exported manually.
# On AWS, real environment variables take precedence (override=False).
try:
    from dotenv import load_dotenv

    _ENV_FILE = Path(__file__).resolve().parents[2] / "config" / ".env"
    load_dotenv(_ENV_FILE, override=False)
except Exception:
    # python-dotenv is optional; without it, plain OS env vars still work.
    pass


def _bool(name: str, default: bool) -> bool:
    val = os.getenv(name)
    if val is None:
        return default
    return val.strip().lower() in ("1", "true", "yes", "on")


# --- HTTP server ---------------------------------------------------------
HOST = os.getenv("OTA_HOST", "127.0.0.1")          # AWS: 0.0.0.0
PORT = int(os.getenv("OTA_PORT", "8000"))

# --- Storage -------------------------------------------------------------
# Where generated Excel files and run artifacts live. Local: ./data. On AWS,
# point this at an EFS mount, or set OTA_STORAGE_BACKEND=s3 to use a bucket.
DATA_DIR = Path(os.getenv("OTA_DATA_DIR", str(Path(__file__).resolve().parent / "data")))
STORAGE_BACKEND = os.getenv("OTA_STORAGE_BACKEND", "local")  # local | s3
OUTPUT_S3_BUCKET = os.getenv("OTA_OUTPUT_S3_BUCKET", "")     # used when backend=s3

# --- OTA domain constants (unchanged from v3) ---------------------------
AWS_PREFIX = os.getenv("OTA_AWS_PREFIX", "https://ota-dl.vidaahub.com")
S3_UPLOAD_BUCKET = os.getenv("OTA_S3_UPLOAD_BUCKET", "fam-media-andr")

# --- FTP (download source ZIPs) -----------------------------------------
# Never hard-code real secrets; supply via env on every environment.
FTP_USER = os.getenv("OTA_FTP_USER", "ftpuser")
FTP_PASS = os.getenv("OTA_FTP_PASS", "")

# --- Templates / history storage ----------------------------------------
# Local defaults to the user home (v3 behaviour); on a server set these to a
# writable, persisted path (EFS) so they survive container restarts.
MESSAGE_TEMPLATES_FILE = os.getenv(
    "OTA_MESSAGE_TEMPLATES_FILE",
    str(DATA_DIR / "ota_message_templates.json"),
)
DEVICE_HISTORY_FILE = os.getenv(
    "OTA_DEVICE_HISTORY_FILE",
    str(DATA_DIR / "device_history.json"),
)

DEBUG = _bool("OTA_DEBUG", True)


def ensure_dirs() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    (DATA_DIR / "outputs").mkdir(parents=True, exist_ok=True)
