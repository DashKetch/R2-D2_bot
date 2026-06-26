# R2-D2 Bot — Setup & Command Reference

## Requirements

```
pip install -r requirements.txt
```

Dependencies: `discord.py`, `openpyxl`, `gspread`, `google-auth`

---

## Initial Setup

### 1. Discord Developer Portal

1. Go to [discord.com/developers/applications](https://discord.com/developers/applications)
2. Click **New Application** → give it a name → **Create**
3. Go to **Bot** → **Add Bot** → **Reset Token** → copy your `BOT_TOKEN`
4. Under **Privileged Gateway Intents**, enable **Server Members Intent**
5. Go to **OAuth2 → URL Generator**:
   - Scopes: `bot`, `applications.commands`
   - Permissions: `Manage Roles`, `Manage Channels`, `Send Messages`, `Embed Links`, `Read Message History`
6. Copy the generated URL, open it in a browser, and invite the bot to your server

> **Role hierarchy:** The bot's role must be placed **above** all Strike roles and any role it needs to assign in Server Settings → Roles. Otherwise it will be unable to assign them.

### 2. Google Sheets

1. Go to [sheets.google.com](https://sheets.google.com) and create a new **native** Google Sheet (not an uploaded .xlsx — if you uploaded one, go to File → Save as Google Sheets)
2. Copy the Sheet ID from the URL: `https://docs.google.com/spreadsheets/d/SHEET_ID_HERE/edit`
3. Go to [console.cloud.google.com](https://console.cloud.google.com):
   - Create or select a project
   - **APIs & Services → Enable APIs** → enable **Google Sheets API**
   - **APIs & Services → Credentials → Create Credentials → Service Account**
   - Click the service account → **Keys** tab → **Add Key → JSON** → save as `credentials.json` in the same folder as `bot.py`
4. Open `credentials.json`, copy the `client_email` value
5. In your Google Sheet, click **Share**, paste that email, set to **Editor**

> `credentials.json` contains sensitive credentials. Never commit it to a public repository. Add it to `.gitignore`.

### 3. Enable Developer Mode in Discord

User Settings → Advanced → **Developer Mode ✓**

This lets you right-click servers, channels, and roles to copy their IDs.

---

## Configuration

Open `bot.py` and fill in every value at the top of the file:

| Variable | What it is | How to get it |
|---|---|---|
| `BOT_TOKEN` | Bot secret token | Developer Portal → Bot → Token |
| `GUILD_ID` | Your server ID | Right-click server icon → Copy Server ID |
| `LOG_CHANNEL_ID` | Channel for strike/appeal logs | Right-click channel → Copy Channel ID |
| `PLACE_CHANNEL_ID` | Channel where placement logs are posted | Right-click channel → Copy Channel ID |
| `PLACEMENT_CMD_CHANNEL` | Channel where `/place` and `/skipplace` can be run | Right-click channel → Copy Channel ID |
| `SHEETS_ALERT_CHANNEL` | Channel for Google Sheets failure alerts | Right-click channel → Copy Channel ID |
| `STRIKE_1_ROLE_ID` | "Strike 1" role | Server Settings → Roles → right-click → Copy Role ID |
| `STRIKE_2_ROLE_ID` | "Strike 2" role | Same as above |
| `CAN_VIEW_ALL_ROLE` | Moderator role (view/appeal strikes, use `/place`, `/skipplace`, `/updatestrikes`) | Same as above |
| `PLACEMENT_STAFF` | Additional role allowed to use `/place` and `/skipplace` | Same as above |
| `DUP_ROLE_1` | Commander Roy Mesular role | Same as above |
| `DUP_ROLE_2` | Captain Major Glow role | Same as above |
| `DUP_ROLE_3` | Captain Liason role | Same as above |
| `GSHEET_CREDS_FILE` | Path to your service account JSON | `"credentials.json"` if it's in the same folder |
| `GSHEET_ID` | Google Sheet ID | From the Sheet URL (see above) |

---

## Running the Bot

```
python bot.py
```

Slash commands sync to your server on first run. If a command doesn't appear immediately, restart your Discord client.

---

## Commands

### `/strike`
**Permission:** `moderate_members` Discord permission

Gives a user a strike. Assigns Strike 1 on first offence; assigns Strike 2 (keeping Strike 1) if they already have Strike 1. Logs to both Google Sheets and `strike_log.xlsx`, and posts an embed to `LOG_CHANNEL_ID`.

| Parameter | Type | Description |
|---|---|---|
| `username` | Member mention | The user to strike |
| `reason` | String | Reason for the strike |

---

### `/viewstrikes`
**Permission:** Anyone (for their own strikes) · `CAN_VIEW_ALL_ROLE` (for others)

Displays a user's full strike history from the sheet. Shows date, issuing moderator, and reason for each strike. Returns N/A for strikes the user doesn't have. Response is ephemeral (only visible to the invoker). Unauthorized attempts to view another user's strikes are logged to `LOG_CHANNEL_ID`.

| Parameter | Type | Description |
|---|---|---|
| `username` | Member mention (optional) | Who to look up — defaults to yourself if omitted |

---

### `/strikeappeal`
**Permission:** `CAN_VIEW_ALL_ROLE`

Removes a strike from the record. Handles four scenarios automatically:

| Situation | Sheet result | Role result |
|---|---|---|
| User has Strike 1 only, appeal Strike 1 | Strike 1 row deleted | Strike 1 role removed |
| User has Strike 1 + Strike 2, appeal Strike 1 | Strike 1 row deleted, Strike 2 row renamed → Strike 1 | Strike 2 role removed, Strike 1 kept |
| User has Strike 2, appeal Strike 2 | Strike 2 row deleted | Strike 2 role removed, Strike 1 kept |
| Strike not found | No changes | No changes |

Logged to `LOG_CHANNEL_ID` with full details of the removed entry and any demotion.

| Parameter | Type | Description |
|---|---|---|
| `username` | Member mention | The user whose strike is being appealed |
| `strike_number` | Integer (1 or 2) | Which strike to remove |

---

### `/place`
**Permission:** `CAN_VIEW_ALL_ROLE` or `PLACEMENT_STAFF` · Must be run in `PLACEMENT_CMD_CHANNEL`

Assigns a role to a user with up to 10 supporting evidence files. Files are **never saved locally** — only their Discord CDN URLs are used. The command response and any warnings are visible to everyone in the channel. Logs to `PLACE_CHANNEL_ID` with clickable file links and an image preview if the first file is an image.

| Parameter | Type | Description |
|---|---|---|
| `username` | Member mention | The user to assign the role to |
| `role` | Role mention | The role to give them |
| `file1` | Attachment (required) | First evidence file |
| `file2`–`file10` | Attachment (optional) | Additional evidence files |

---

### `/skipplace`
**Permission:** `CAN_VIEW_ALL_ROLE` or `PLACEMENT_STAFF` · Must be run in `PLACEMENT_CMD_CHANNEL`

Identical to `/place` but for returning users who were previously placed and have rejoined. Accepts one file. The log embed is clearly marked as a skipped placement and explains why, so the record is unambiguous. Response is visible to all in the channel.

| Parameter | Type | Description |
|---|---|---|
| `username` | Member mention | The returning user |
| `role` | Role mention | The role to restore |
| `file1` | Attachment (required) | Evidence of previous placement |

---

### `/dupcategory`
**Permission:** `DUP_ROLE_1`, `DUP_ROLE_2`, or `DUP_ROLE_3`

Duplicates all channels (text, voice, stage, forum) from an existing category into a new one, preserving channel names, order, and all permission overwrites. Does not copy message history or pins.

| Parameter | Type | Description |
|---|---|---|
| `old_category_id` | String (ID) | ID of the category to copy |
| `new_category_name` | String | Name for the new category |
| `new_category_index` | String (integer) | Position of the new category in the channel list (0 = top) |

> IDs are entered as strings to avoid Discord's integer size limit. Paste the raw number exactly as copied.

---

### `/duprole`
**Permission:** `DUP_ROLE_1`, `DUP_ROLE_2`, or `DUP_ROLE_3`

Creates a new role by copying all permissions from an existing role, then applying a new name, colour, and position. Also copies the hoist and mentionable settings from the source role.

| Parameter | Type | Description |
|---|---|---|
| `old_role_id` | String (ID) | ID of the role to copy permissions from |
| `name` | String | Name for the new role |
| `color_code` | String | Hex colour, e.g. `#FF5733` or `FF5733` |
| `position` | String (integer) | Position in the role hierarchy (1 = bottom) |

> The bot can only place roles below its own highest role. If the requested position is above that, the role is still created and you'll be prompted to set the position manually.

---

### `/updatestrikes`
**Permission:** `CAN_VIEW_ALL_ROLE`

Syncs the local `strike_log.xlsx` backup to Google Sheets. Only needed after a Google Sheets outage — the bot posts an alert to `SHEETS_ALERT_CHANNEL` with a prompt to run this command whenever it falls back to local storage. Clears and rewrites all Google Sheet data rows from the local file, then posts a confirmation embed.

---

## Google Sheets Fallback Behaviour

Every write operation (strike, appeal) attempts **both** Google Sheets and local xlsx simultaneously. If either fails, an alert is posted to `SHEETS_ALERT_CHANNEL` pinging `CAN_VIEW_ALL_ROLE` specifying exactly which destination failed and why. Reads fall back to local silently since no data is modified.

Your data is never lost — at least one destination will have it as long as the bot machine is running.

---

## Files

| File | Purpose |
|---|---|
| `bot.py` | Main bot source |
| `requirements.txt` | Python dependencies |
| `credentials.json` | Google service account credentials **(never share or commit)** |
| `strike_log.xlsx` | Local strike log backup (auto-created on first run) |