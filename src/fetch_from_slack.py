"""Slack fetcher for the cloud pipeline. Downloads today's execution_report_*.csv."""
from __future__ import annotations

import json
import os
import re
import socket
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

SLACK_API = "https://slack.com/api"
FILE_PATTERN = re.compile(r"^execution_report_(\d{4}-\d{2}-\d{2})\.csv$", re.IGNORECASE)
RETRY_ATTEMPTS = 5
RETRY_INITIAL_DELAY_S = 5


def _http(url: str, *, token: str | None = None, raw: bool = False):
    delay = RETRY_INITIAL_DELAY_S
    last_error = None
    for attempt in range(1, RETRY_ATTEMPTS + 1):
        try:
            req = urllib.request.Request(url)
            if token:
                req.add_header("Authorization", f"Bearer {token}")
            with urllib.request.urlopen(req, timeout=90) as resp:
                data = resp.read()
            return data if raw else json.loads(data.decode("utf-8"))
        except (urllib.error.URLError, socket.timeout, TimeoutError) as e:
            last_error = e
            if attempt == RETRY_ATTEMPTS:
                break
            time.sleep(delay)
            delay *= 2
    raise last_error


def _is_scheduled_run(message: dict) -> bool:
    """Only trust the automatic daily bot post, not manual/on-demand re-runs.

    The bot's own fallback text (and blocks) contain "triggered by schedule"
    for the real cron-fired post, vs "triggered by @some-user" for manual runs.
    We check the whole serialized message defensively since Block Kit messages
    put the readable text in nested blocks, not always the top-level `text` field.
    """
    blob = json.dumps(message).lower()
    return "triggered by schedule" in blob


def _find_latest_csv(token: str, channel_id: str) -> dict | None:
    params = urllib.parse.urlencode({"channel": channel_id, "limit": 50})
    resp = _http(f"{SLACK_API}/conversations.history?{params}", token=token)
    if not resp.get("ok"):
        raise RuntimeError(f"Slack conversations.history failed: {resp.get('error')}")
    for m in resp.get("messages", []):
        if not _is_scheduled_run(m):
            continue
        for f in m.get("files", []) or []:
            if FILE_PATTERN.match(f.get("name", "")):
                return f
    return None


def fetch_latest_csv(dest_dir: Path, token: str, channel_id: str | None = None) -> Path | None:
    channel_id = channel_id or os.environ.get("SLACK_CHANNEL_ID")
    if not channel_id:
        raise ValueError("channel_id not provided and SLACK_CHANNEL_ID env var not set")
    f = _find_latest_csv(token, channel_id)
    if not f:
        return None
    url = f.get("url_private_download") or f.get("url_private")
    if not url:
        raise RuntimeError("Slack file object missing url_private_download")
    blob = _http(url, token=token, raw=True)
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / f["name"]
    dest.write_bytes(blob)
    return dest
