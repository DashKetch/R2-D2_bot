# Strike Bot — Setup Guide

## 1. Install dependencies
```
pip install -r requirements.txt
```

## 2. Fill in the placeholders at the top of `bot.py`

| Placeholder | What to put there | Where to find it |
|---|---|---|
| `BOT_TOKEN` | Your bot's secret token | Discord Developer Portal → Your App → Bot → Token |
| `GUILD_ID` | Your server's ID (integer) | Right-click your server icon → Copy Server ID* |
| `LOG_CHANNEL_ID` | Channel where strikes are posted | Right-click the channel → Copy Channel ID* |
| `STRIKE_1_ROLE_ID` | Role ID for "Strike 1" | Server Settings → Roles → right-click role → Copy Role ID* |
| `STRIKE_2_ROLE_ID` | Role ID for "Strike 2" | Same as above |

*Developer Mode must be ON: User Settings → Advanced → Developer Mode ✓

## 3. Bot permissions required
In the Developer Portal, under **OAuth2 → URL Generator**, enable:
- Scopes: `bot`, `applications.commands`
- Bot Permissions: `Manage Roles`, `Send Messages`, `Embed Links`, `Moderate Members`

> ⚠️ The bot's role must be **above** both Strike roles in the role hierarchy, or it won't be able to assign them.

## 4. Enable the Members Intent
Developer Portal → Your App → Bot → Privileged Gateway Intents → **Server Members Intent** ✓

## 5. Run the bot
```
python bot.py
```
The slash command `/strike` will sync to your server on first run.
The spreadsheet `strike_log.xlsx` is created automatically in the same folder.

---

## How `/strike` works

```
/strike username:@SomeUser reason:Spamming in #general
```

| Situation | Result |
|---|---|
| User has no strike | Assigns **Strike 1** role |
| User already has Strike 1 | Removes Strike 1, assigns **Strike 2** role |
| User already has Strike 2 | No action taken, mod gets an ephemeral warning |

Every successful strike is:
- Posted as a styled embed in the log channel
- Appended to `strike_log.xlsx` with columns: Moderator · Date · Username · Strike # · Reason
