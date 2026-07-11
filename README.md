# Eachlabs Daily Execution Report — Cloud Edition

Automated daily pipeline that fetches an execution report from Slack, builds a
multi-tab Excel analysis, uploads it to Google Drive, and posts the link back
to Slack. Runs entirely on GitHub Actions — no laptop needs to be online.

## What it does, once per day

1. **Fetch** the newest `execution_report_YYYY-MM-DD.csv` from Slack `#custom-report`
2. **Build** a formatted xlsx with: Summary, Cockpit (MoM + projection), Monthly cockpit, Top movers up/down, First-time today, Drop-offs, MTD pivot, Monthly pivot, Model usage, Seedance 2.0 Usage, Raw, Source
3. **Upload** the xlsx to a Google Drive folder via service account
4. **Post** a link message to Slack

Schedule: **07:00 Europe/Istanbul** (04:00 UTC).
Change the cron in `.github/workflows/daily.yml` to adjust.

## Repo layout

```
eachlabs-daily/
├── .github/workflows/daily.yml   GitHub Actions schedule + workflow
├── src/
│   ├── fetch_from_slack.py       Slack file download (with retry)
│   ├── daily_report.py           Report builder (12 tabs)
│   ├── upload_drive.py           Google Drive upload
│   └── notify_slack.py           Slack chat.postMessage
├── run.py                         Orchestrator: fetch → build → upload → notify
├── requirements.txt
├── .env.example                   Template — real .env is gitignored
├── .gitignore
└── README.md                      This file
```

## Setup checklist

Work through these in order. Everything is one-time.

- [ ] GitHub repo created (private, personal account)
- [ ] Files pushed to the repo
- [ ] Slack app + User OAuth token in hand (scopes below)
- [ ] Google Cloud project created
- [ ] Google Drive API enabled
- [ ] Service account created + JSON key downloaded
- [ ] Drive folder shared with service account email as Editor
- [ ] 4 GitHub Secrets added
- [ ] Manual test run via **Run workflow** button
- [ ] Cron test — wait for next 04:00 UTC or trigger manually

## Step 1 — Create the GitHub repo

1. https://github.com/new
2. Name: `eachlabs-daily` (or whatever you prefer)
3. **Private**
4. Do NOT initialize with README (we already have one)
5. Create

Then locally:

```bash
cd ~/Documents/eachlabs-daily-cloud
git init
git add .
git commit -m "Initial cloud pipeline"
git branch -M main
git remote add origin https://github.com/YOUR-USERNAME/eachlabs-daily.git
git push -u origin main
```

## Step 2 — Slack app + token

You may already have a Slack app from the local setup. If yes, reuse that token. If not:

1. https://api.slack.com/apps → **Create New App** → **From scratch**
2. Name: `Eachlabs Daily`, workspace: Eachlabs
3. Left sidebar → **OAuth & Permissions**
4. Scroll to **User Token Scopes** and add these four:
   - `groups:history` — read messages in private channels
   - `groups:read` — list private channels
   - `files:read` — read files
   - `chat:write` — post messages
5. Scroll up → **Install to Workspace** → authorize
6. Copy the **User OAuth Token** (`xoxp-...`). You'll paste it as a GitHub Secret in Step 5.

Also copy your channel ID:
- In Slack, open `#custom-report` → click the channel name at the top → scroll to bottom → **Channel ID** (looks like `C0A84HY8B41`)

## Step 3 — Google Cloud project + service account

1. Go to https://console.cloud.google.com/
2. Top bar → project dropdown → **New Project**
   - Name: `eachlabs-daily`
   - Create
3. Left menu → **APIs & Services → Library**
4. Search **Google Drive API** → **Enable**
5. Left menu → **APIs & Services → Credentials → Create Credentials → Service account**
   - Name: `eachlabs-daily-writer`
   - Skip the optional roles page (we grant access at the folder level instead)
   - Done
6. In the service accounts list, click your new account → **Keys** tab → **Add Key → Create new key → JSON → Create**
7. A JSON file downloads (e.g. `eachlabs-daily-abc123.json`). **Never commit this.**
8. Open the JSON, find `"client_email"` (looks like `eachlabs-daily-writer@eachlabs-daily.iam.gserviceaccount.com`). Copy that email.

## Step 4 — Create the Drive folder and share it

1. https://drive.google.com/ → **New → New folder**
2. Name: `Eachlabs Daily Reports` (or whatever)
3. Right-click the folder → **Share**
4. Paste the service account email (from step 3.8)
5. Role: **Editor**
6. Uncheck "Notify people" (no email will arrive at that address anyway)
7. Share
8. Open the folder in Drive. The URL is `https://drive.google.com/drive/folders/{FOLDER_ID}`. Copy `{FOLDER_ID}` — you'll paste it as a secret.

## Step 5 — GitHub Secrets

**What's a GitHub Secret?** It's an encrypted environment variable stored per-repo. GitHub Actions can read them at runtime but they never appear in logs or the repo. You add them in the repo web UI — not in your code.

1. In your repo on GitHub, click **Settings** (top right, next to Insights)
2. Left sidebar → **Secrets and variables → Actions**
3. Click **New repository secret** — you'll do this four times

Add these four secrets exactly:

| Secret name | Value |
|---|---|
| `SLACK_USER_TOKEN` | Your `xoxp-...` token from Step 2.6 |
| `SLACK_CHANNEL_ID` | Your channel ID from Step 2 (e.g. `C0A84HY8B41`) |
| `GDRIVE_FOLDER_ID` | Folder ID from Step 4.8 |
| `GOOGLE_APPLICATION_CREDENTIALS_JSON` | The **entire content** of the JSON file from Step 3.6, pasted as-is (multiline is fine) |

For the JSON secret, open the file in a text editor and copy everything from `{` through `}`. Paste into the Value field. GitHub stores it as an opaque string — the workflow reads it via `os.environ["GOOGLE_APPLICATION_CREDENTIALS_JSON"]` and parses the JSON inside the app.

## Step 6 — First run

1. In your repo → **Actions** tab
2. Left sidebar → **Daily execution report**
3. Top right → **Run workflow** → **Run workflow**
4. Watch it run. Green check = success. Click into the run to see step-by-step output.

If successful:
- The report shows up in your Drive folder
- A message appears in `#custom-report`

## Local development (optional)

If you want to test on your Mac before pushing:

```bash
cd ~/Documents/eachlabs-daily-cloud
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# Copy the template and fill in real values
cp .env.example .env
# edit .env with your token, channel, folder id, and either the JSON string
# or a path to the downloaded service account file

# Load env and run
set -a; source .env; set +a
python run.py
```

## Common problems

| Symptom | Likely cause | Fix |
|---|---|---|
| Workflow runs at wrong hour | Cron is UTC, not local | Convert your local time to UTC. Turkey = UTC+3 year-round |
| `File not found` on Drive upload | Folder not shared with service account | Re-share the folder with `client_email` from the JSON, role Editor |
| `Google Drive API has not been used` | API not enabled in the project | Google Cloud Console → APIs & Services → Library → Google Drive API → Enable |
| `invalid_auth` from Slack | Token typo or missing scope | Reinstall the Slack app to your workspace after adding scopes; new token is issued |
| `.env` shows up on GitHub | `.gitignore` wasn't respected on first commit | Delete `.env` from history: `git rm --cached .env; git commit; git push`. Rotate all leaked secrets immediately |
| JSON secret stored wrong | Trailing whitespace or wrong quoting | Re-paste the raw JSON with no wrapping quotes. GitHub handles multiline pastes |
| Actions billing warning | Exceeded free minutes | Personal free tier is 2000 min/mo private repos — this pipeline uses ~2 min/day. Should stay well under |
| Cron runs but doesn't post | Missing `chat:write` scope | Add scope → reinstall app → replace token secret |
| Manual `run.py` works, Actions fails | Different Python versions or missing dep | Pin `python-version: '3.11'` in workflow (already done); check requirements.txt matches local |
| Report has wrong "today" | CSV data lags one day (Canbo-Helper posts previous day's data) | This is expected — a CSV named 07-10 contains data through 07-09 |

## Migrating from the current local setup

You can keep both running in parallel during the transition. To stop the local pipeline once cloud is verified:

```bash
# Stop the local launchd fetcher
launchctl unload ~/Library/LaunchAgents/com.eachlabs.slack-fetch.plist

# Optionally remove the cron wake
sudo pmset repeat cancel

# The Cowork scheduled task can be disabled from Cowork → Scheduled tasks
```

The local `~/Documents/eachlabs-daily/` folder and its archived CSVs can stay
for reference — they don't interfere with anything.
