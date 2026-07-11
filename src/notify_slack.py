"""Post a message to a Slack channel using chat:write scope.

Uses the same user OAuth token as the fetcher.
"""

from __future__ import annotations

import json
import os
import urllib.error
import urllib.request

SLACK_API = "https://slack.com/api"


def send_message(channel_id: str, text: str, token: str) -> dict:
    """Post a message. Returns the raw Slack response dict (must check .get('ok'))."""
    payload = {
        "channel": channel_id,
        "text": text,
        "mrkdwn": True,
    }
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(f"{SLACK_API}/chat.postMessage", data=data, method="POST")
    req.add_header("Content-Type", "application/json; charset=utf-8")
    req.add_header("Authorization", f"Bearer {token}")
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode("utf-8"))


if __name__ == "__main__":
    import sys
    channel = os.environ["SLACK_CHANNEL_ID"]
    token = os.environ["SLACK_USER_TOKEN"]
    text = sys.argv[1] if len(sys.argv) > 1 else "test message"
    print(json.dumps(send_message(channel, text, token), indent=2))
