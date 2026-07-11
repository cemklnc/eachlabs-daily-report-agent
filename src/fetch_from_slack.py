"""Slack fetcher for the cloud pipeline.

Downloads the most recent execution_report_*.csv posted by the Canbo-Helper
bot in the target channel and returns its local path.

No persistent inbox/archive here — the caller writes to a temp dir and
uploads results elsewhere.
"""

from __future__ import annotations

import json
import os
import re
import socket
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path

SLACK_API = "https://slack.com/api"
FILE_PATTERN = re.compile(r"^execution_report_(\d{4}-\d{2}-\d{2})\.csv$", re.IGNORECASE)

# Retry config — GitHub Actions runners get transient DNS/timeout failures too
RETRY_ATTEMPTS = 5
RETRY_INITIAL_DELAY_S = 5


def _http(url: str, *, token: str | None = None, raw: bool = False):
    """HTTP GET with retry-with-backoff on transient network errors."""
    delay = RETRY_INITIAL_DELAY_S
    last_error: Exception | None = None
    for attempt in range(1, RETRY_ATTEMPTS + 1):
        try:
            req = urllib.request.Request(url)
            if token:
                req.add_header("Authorization", f"Bearer {token}")
            with urllib.request.urlopen(req, timeout=90) as resp:
                data = resp.read()
            if raw:
                return data
            return json.loads(data.decode("utf-8"))
        except (urllib.error.URLError, socket.timeout, TimeoutError) as e:
            last_error = e
            if attempt == RETRY_ATTEMPTS:
                break
            time.sleep(delay)
            delay *= 2
    assert last_error is not None
    raise last_error


def _find_latest_csv(token: str, channel_id: str) -> dict | None:
    """Walk recent messages, return the newest execution_report_*.csv file object."""
    params = urllib.parse.urlencode({"channel": channel_id, "limit": 50})
    resp = _http(f"{SLACK_API}/conversations.history?{params}", token=token)
    if not resp.get("ok"):
        raise RuntimeError(f"Slack conversations.history failed: {resp.get('error')}")
    for m in resp.get("messages", []):
        for f in m.get("files", []) or []:
            if FILE_PATTERN.match(f.get("name", "")):
                return f
    return None


def fetch_latest_csv(dest_dir: Path, token: str, channel_id: str | None = None) -> Path | None:
    """Download the most recent execution_report CSV from the Slack channel.

    Returns the local path, or None if no matching file was found.
    channel_id defaults to SLACK_CHANNEL_ID env var.
    """
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


if __name__ == "__main__":
    import sys
    import tempfile

    tok = os.environ.get("SLACK_USER_TOKEN")
    if not tok:
        print("ERROR: SLACK_USER_TOKEN env var not set")
        sys.exit(3)
    with tempfile.TemporaryDirectory() as tmp:
        path = fetch_latest_csv(Path(tmp), tok)
        if path:
            print(json.dumps({
                "status": "ok",
                "name": path.name,
                "bytes": path.stat().st_size,
                "fetched_at": datetime.utcnow().isoformat(timespec="seconds") + "Z",
            }, indent=2))
        else:
            print(json.dumps({"status": "no_file"}))
            sys.exit(2)
