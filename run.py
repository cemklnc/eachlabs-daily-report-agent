"""Daily orchestrator — fetch → build → upload → notify.

Run once per day via GitHub Actions cron (or locally via `python run.py`).

Env vars required:
  SLACK_USER_TOKEN                        Slack user OAuth token (xoxp-)
  SLACK_CHANNEL_ID                        Channel to fetch from + post to
  GDRIVE_FOLDER_ID                        Drive folder for uploads
  GOOGLE_APPLICATION_CREDENTIALS_JSON     Service account JSON (Actions), OR
  GOOGLE_APPLICATION_CREDENTIALS          Path to service account json (local)

Exit codes:
  0  ok (or no new file — treated as success)
  2  no CSV found in Slack history
  3  missing required env var
  4  runtime failure (network, API, etc.)
"""

from __future__ import annotations

import os
import sys
import tempfile
import traceback
from pathlib import Path

# Make ./src importable without needing to install as package
sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from fetch_from_slack import fetch_latest_csv          # noqa: E402
from daily_report import build_report_from_csv         # noqa: E402
from upload_drive import upload_file, file_exists_in_folder  # noqa: E402
from notify_slack import send_message                  # noqa: E402


REQUIRED_ENV = [
    "SLACK_USER_TOKEN",
    "SLACK_BOT_TOKEN",
    "SLACK_CHANNEL_ID",
    "GDRIVE_FOLDER_ID",
]


def _check_env() -> None:
    missing = [k for k in REQUIRED_ENV if not os.environ.get(k)]
    if missing:
        raise SystemExit(f"Missing env vars: {', '.join(missing)}\n(exit 3)")
    if not (
        os.environ.get("GOOGLE_APPLICATION_CREDENTIALS_JSON")
        or os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
    ):
        raise SystemExit(
            "Missing Google credentials. Set GOOGLE_APPLICATION_CREDENTIALS_JSON "
            "(JSON content) or GOOGLE_APPLICATION_CREDENTIALS (path).\n(exit 3)"
        )


def main() -> int:
    try:
        _check_env()
    except SystemExit as e:
        print(str(e))
        return 3

    slack_token = os.environ["SLACK_USER_TOKEN"]        # reads Slack history/files
    slack_bot_token = os.environ["SLACK_BOT_TOKEN"]      # posts as the bot, not personal account
    slack_channel = os.environ["SLACK_CHANNEL_ID"]
    drive_folder = os.environ["GDRIVE_FOLDER_ID"]

    with tempfile.TemporaryDirectory(prefix="eachlabs-daily-") as tmp:
        tmp_path = Path(tmp)

        # 1. Fetch today's CSV from Slack
        print("→ Fetching latest CSV from Slack…", flush=True)
        try:
            csv_path = fetch_latest_csv(tmp_path, slack_token, slack_channel)
        except Exception as e:
            print(f"  ERROR: Slack fetch failed: {type(e).__name__}: {e}")
            traceback.print_exc()
            return 4
        if not csv_path:
            print("  No matching CSV in recent history. Exiting cleanly.")
            return 0
        print(f"  ✓ Downloaded: {csv_path.name} ({csv_path.stat().st_size:,} bytes)")

        # 2. Build the report
        print("→ Building daily report…", flush=True)
        try:
            report_path, report_date = build_report_from_csv(csv_path, tmp_path)
        except Exception as e:
            print(f"  ERROR: Report build failed: {type(e).__name__}: {e}")
            traceback.print_exc()
            return 4
        print(f"  ✓ Generated: {report_path.name} ({report_path.stat().st_size:,} bytes)")

        # 3. Idempotency check — skip upload if a same-named report is already in Drive
        try:
            existing = file_exists_in_folder(report_path.name, drive_folder)
        except Exception as e:
            print(f"  ERROR: Drive lookup failed: {type(e).__name__}: {e}")
            traceback.print_exc()
            return 4

        if existing:
            file_link = existing["webViewLink"]
            print(f"  ⚠ Same-named report already in Drive: {file_link}")
        else:
            print("→ Uploading to Google Drive…", flush=True)
            try:
                uploaded = upload_file(report_path, drive_folder)
            except Exception as e:
                print(f"  ERROR: Drive upload failed: {type(e).__name__}: {e}")
                traceback.print_exc()
                return 4
            file_link = uploaded["webViewLink"]
            print(f"  ✓ Uploaded: {file_link}")

        # 4. Post to Slack
        print("→ Notifying Slack…", flush=True)
        message = (
            f"*Eachlabs daily report — {report_date}*\n"
            f"<{file_link}|Open report>"
        )
        try:
            result = send_message(slack_channel, message, slack_bot_token)
        except Exception as e:
            print(f"  ERROR: Slack post failed: {type(e).__name__}: {e}")
            traceback.print_exc()
            return 4

        if not result.get("ok"):
            print(f"  ERROR: Slack API returned not-ok: {result.get('error')}")
            return 4

        ts = result.get("ts", "").replace(".", "")
        print(f"  ✓ Posted: https://eachlabs.slack.com/archives/{slack_channel}/p{ts}")

    print("Done.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
