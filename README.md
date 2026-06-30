# R2‑D2 Bot

Lightweight Discord moderation and utility bot focused on strike tracking,
placement logging, role/channel utilities, and robust Google Sheets failover.

Key features
- Strike management (writes to Google Sheets and a local Excel fallback)
- Strike appeals and demotion handling
- Placement logging with evidence attachments (files are hosted on Discord)
- Category/role duplication and permission migration helpers
- ID utilities for quick role/channel lookups

Requirements
- Python 3.10+
- See `requirements.txt` for pinned dependencies

Install
```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Quickstart (local)
1. Copy `fields.py.example` → `fields.py` (or edit `fields.py`) and fill values.
2. Place your Google service account JSON (recommended name: `credentials.json`).
3. Run:
```bash
python bot.py
```

Configuration
All runtime configuration lives in `fields.py`. At minimum you should set:
- `BOT_TOKEN` — your bot token
- `GUILD_ID` — guild to sync slash commands to
- Channel and role IDs: `LOG_CHANNEL_ID`, `PLACE_CHANNEL_ID`, `PLACEMENT_CMD_CHANNEL`, `SHEETS_ALERT_CHANNEL`, `STRIKE_1_ROLE_ID`, `STRIKE_2_ROLE_ID`, `CAN_VIEW_ALL_ROLE`, `PLACEMENT_STAFF` and duplication permission roles (`DUP_ROLE_1/2/3`).
- `GSHEET_CREDS_FILE` — path to service account JSON (default: `credentials.json`)
- `GSHEET_ID` — Google Sheet ID for strike logs
- `SPREADSHEET_PATH` — local backup path (default: `strike_log.xlsx`)

fields.py (example snippet)
```python
BOT_TOKEN = ""
GUILD_ID = 0

LOG_CHANNEL_ID = 0
PLACE_CHANNEL_ID = 0
PLACEMENT_CMD_CHANNEL = 0
SHEETS_ALERT_CHANNEL = 0

STRIKE_1_ROLE_ID = 0
STRIKE_2_ROLE_ID = 0

CAN_VIEW_ALL_ROLE = 0
PLACEMENT_STAFF = 0

DUP_ROLE_1 = 0
DUP_ROLE_2 = 0
DUP_ROLE_3 = 0

PLACEMENT_QUEUE_ROLE_IDS = [0, 0]

GSHEET_CREDS_FILE = "credentials.json"
GSHEET_ID = ""

SPREADSHEET_PATH = "strike_log.xlsx"
```

Google Sheets setup (brief)
1. Create a Google Cloud project and enable Sheets & Drive APIs.
2. Create a service account, add a JSON key, and save it as `credentials.json`.
3. Share your Google Sheet with the service account `client_email` with Editor access.
4. Put the sheet ID into `GSHEET_ID`.

Security
- Never commit `credentials.json` or `strike_log.xlsx` to version control.
- Add these to `.gitignore`:
```
credentials.json
strike_log.xlsx
__pycache__/
*.pyc
```

Permissions & Bot role
- The bot requires these permissions: View Channels, Send Messages, Embed Links, Read Message History, Manage Roles, Manage Channels, Moderate Members, and Attach Files.
- Ensure the bot's highest role is above any role it needs to manage/assign.

Commands overview
- Moderation: `/strike`, `/viewstrikes`, `/strikeappeal`, `/updatestrikes`
- Placement: `/place`, `/skipplace`
- Duplication / migration: `/dupcategory`, `/duprole`, `/cutchannelperms`, `/cutcategoryperms`
- Utilities: `/roleids`, `/channelids`

Behavior notes
- On write, the bot attempts Google Sheets first then a local Excel fallback. If either destination fails, moderators are alerted in `SHEETS_ALERT_CHANNEL`.
- Use `/updatestrikes` to sync local backup to Google Sheets after recovering cloud access.

Running in production
- Run under a process manager (systemd, pm2, Docker) to ensure automatic restarts.

Troubleshooting
- If slash commands don't appear, ensure `GUILD_ID` is set and the bot has `applications.commands` scope; the bot syncs commands to the configured guild on startup.
- Google Sheets errors usually mean `GSHEET_CREDS_FILE` or `GSHEET_ID` are incorrect, or the service account wasn't shared with the sheet.

Files
- `bot.py` — main bot implementation
- `fields.py` — configuration values you must edit
- `requirements.txt` — Python dependencies
- `credentials.json` — Google service account (keep secret)
- `strike_log.xlsx` — local backup created automatically

License
See `LICENSE.md`.

If you'd like, I can:
- generate a `fields.py.example` from the variables found in `bot.py` and `fields.py`,
- add a minimal systemd service or Dockerfile for deployment,
- or run a quick static check on `bot.py` for obvious issues.
