# R2-D2 Bot

A moderation and utility bot for Discord featuring:

- Strike management with Google Sheets + local backup
- Strike appeals
- Placement logging with evidence attachments
- Category duplication
- Role duplication
- Channel permission migration
- Role and channel ID utilities
- Automatic Google Sheets failover and recovery

---

# Requirements

## Python

Python 3.10+ is recommended.

Install dependencies:

```bash
pip install -r requirements.txt
```

Required packages:

- discord.py
- openpyxl
- gspread
- google-auth

---

# Project Structure

```text
.
├── bot.py
├── fields.py
├── requirements.txt
├── credentials.json
└── strike_log.xlsx
```

---

# Discord Bot Setup

## 1. Create the Application

Go to:

https://discord.com/developers/applications

1. Click **New Application**
2. Enter a name
3. Click **Create**

---

## 2. Create the Bot

Go to:

**Bot → Add Bot**

Then:

- Reset Token
- Copy the token

You will place this into:

```python
BOT_TOKEN = "YOUR_TOKEN"
```

inside `fields.py`.

---

## 3. Enable Privileged Intents

Under **Bot → Privileged Gateway Intents**:

Enable:

✅ Server Members Intent

No other intents are required.

---

## 4. Invite the Bot

Go to:

**OAuth2 → URL Generator**

### Scopes

- bot
- applications.commands

### Bot Permissions

- View Channels
- Send Messages
- Embed Links
- Read Message History
- Manage Roles
- Manage Channels
- Moderate Members
- Attach Files

Open the generated URL and invite the bot.

---

# IMPORTANT: Role Hierarchy

The bot's role MUST be above:

- Strike roles
- Placement roles
- Any role the bot will assign
- Any role whose permissions will be modified

Otherwise Discord will reject the action.

---

# Google Sheets Setup

The strike system uses both:

- Google Sheets
- Local Excel backup (`strike_log.xlsx`)

---

## Create the Sheet

Go to:

https://sheets.google.com

Create a new **native Google Sheet**.

Do NOT use an uploaded `.xlsx` file.

Copy the Sheet ID:

```text
https://docs.google.com/spreadsheets/d/SHEET_ID_HERE/edit
```

---

## Enable APIs

Go to:

https://console.cloud.google.com

Create or select a project.

Enable:

- Google Sheets API
- Google Drive API

---

## Create Service Account

Go to:

APIs & Services → Credentials → Create Credentials → Service Account

Then:

1. Open the service account
2. Go to Keys
3. Add Key
4. JSON

Save the file as:

```text
credentials.json
```

in the bot folder.

---

## Share the Sheet

Open:

```json
"client_email"
```

inside `credentials.json`.

Copy the email.

Open your Google Sheet:

Share → Add People

Paste the email and give it:

Editor permissions.

---

# Security

Never upload:

- credentials.json
- strike_log.xlsx

to a public repository.

Recommended `.gitignore`:

```gitignore
credentials.json
strike_log.xlsx
__pycache__/
*.pyc
```

---

# Enable Developer Mode

Discord:

User Settings → Advanced → Developer Mode

This allows you to copy:

- Server IDs
- Channel IDs
- Role IDs

---

# Configuration

All configuration is stored inside:

```text
fields.py
```

Example:

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

PLACEMENT_QUEUE_ROLE_IDS = [
    0,
    0,
]

GSHEET_CREDS_FILE = "credentials.json"
GSHEET_ID = ""

SPREADSHEET_PATH = "strike_log.xlsx"
```

---

# Configuration Variables

| Variable | Description |
|----------|-------------|
| BOT_TOKEN | Discord bot token |
| GUILD_ID | Server ID |
| LOG_CHANNEL_ID | Strike log channel |
| PLACE_CHANNEL_ID | Placement log channel |
| PLACEMENT_CMD_CHANNEL | Channel where placement commands can be used |
| SHEETS_ALERT_CHANNEL | Google Sheets failure alerts |
| STRIKE_1_ROLE_ID | Strike 1 role |
| STRIKE_2_ROLE_ID | Strike 2 role |
| CAN_VIEW_ALL_ROLE | Moderator role |
| PLACEMENT_STAFF | Additional placement role |
| DUP_ROLE_1 | Duplication permission role |
| DUP_ROLE_2 | Duplication permission role |
| DUP_ROLE_3 | Duplication permission role |
| PLACEMENT_QUEUE_ROLE_IDS | Queue roles automatically removed during placement |
| GSHEET_CREDS_FILE | Path to service account JSON |
| GSHEET_ID | Google Sheet ID |
| SPREADSHEET_PATH | Local strike backup |

---

# Running the Bot

```bash
python bot.py
```

On first startup:

- slash commands sync automatically
- `strike_log.xlsx` is created automatically if missing
- Google Sheet headers are created automatically if the sheet is empty

---

# Commands

## Moderation

- `/strike`
- `/viewstrikes`
- `/strikeappeal`
- `/updatestrikes`

## Placement

- `/place`
- `/skipplace`

## Duplication Utilities

- `/dupcategory`
- `/duprole`
- `/cutchannelperms`

## ID Utilities

- `/roleids`
- `/channelids`

---

# Google Sheets Failover

Every strike write attempts:

1. Google Sheets
2. Local Excel backup

If either fails:

- an alert is posted in `SHEETS_ALERT_CHANNEL`
- moderators are pinged
- data is preserved whenever at least one destination succeeds

When Google Sheets comes back online, run:

```text
/updatestrikes
```

to restore the cloud copy.

---

# Files

| File | Purpose |
|------|----------|
| bot.py | Main bot source |
| fields.py | Configuration values |
| requirements.txt | Python dependencies |
| credentials.json | Google service account credentials |
| strike_log.xlsx | Local backup database |

---

# License

This project is source-available. See `LICENSE.md` for licensing information.
