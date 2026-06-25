import discord
from discord import app_commands
from discord.ext import commands
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime, timezone
import os
import gspread
from google.oauth2.service_account import Credentials
from fields import *

intents = discord.Intents.default()
intents.members = True

bot = commands.Bot(command_prefix="!", intents=intents)
tree = bot.tree


# ── Shared helpers ────────────────────────────────────────────────────────────

def _safe_send(interaction: discord.Interaction):
    """Return the right send method depending on whether we've already deferred."""
    return interaction.followup.send if interaction.response.is_done() else interaction.response.send_message


async def _post_to_log(guild: discord.Guild, embed: discord.Embed):
    """Safely post an embed to the strike log channel."""
    log_channel = guild.get_channel(LOG_CHANNEL_ID)
    if log_channel:
        try:
            await log_channel.send(embed=embed)
        except discord.Forbidden:
            pass


# ── Spreadsheet helpers ───────────────────────────────────────────────────────

HEADERS = ["Moderator", "Date command used", "Username", "Strike (1 or 2)", "Reason"]

# ── Google Sheets connection ──────────────────────────────────────────────────

def _get_gsheet():
    """
    Return the first worksheet of the configured Google Sheet.
    Raises an exception if credentials are missing or the sheet can't be reached —
    callers should catch and fall back to the local xlsx.
    """
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds  = Credentials.from_service_account_file(GSHEET_CREDS_FILE, scopes=scopes)
    client = gspread.authorize(creds)
    sheet  = client.open_by_key(GSHEET_ID)
    ws     = sheet.get_worksheet(0)
    # Write header row if the sheet is completely empty
    if ws.row_count == 0 or ws.cell(1, 1).value is None:
        ws.update("A1:E1", [HEADERS])
        ws.format("A1:E1", {
            "backgroundColor": {"red": 0.12, "green": 0.31, "blue": 0.47},
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "horizontalAlignment": "CENTER",
        })
    return ws


# ── Local xlsx fallback helpers ───────────────────────────────────────────────

def _cell_style(ws, row: int, col: int, value, strike_num_col: bool = False, strike_num: int = 1):
    fill_color     = "D6E4F0" if (row % 2 == 0) else "FFFFFF"
    cell           = ws.cell(row=row, column=col, value=value)
    cell.fill      = PatternFill("solid", start_color=fill_color, end_color=fill_color)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border    = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin"),
    )
    if strike_num_col:
        cell.font = Font(name="Arial", size=10, bold=True,
                         color="C00000" if strike_num == 2 else "FF6600")
    else:
        cell.font = Font(name="Arial", size=10)
    return cell


def _init_spreadsheet():
    if os.path.exists(SPREADSHEET_PATH):
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Strike Log"
    header_fill  = PatternFill("solid", start_color="1F4E79", end_color="1F4E79")
    header_font  = Font(bold=True, color="FFFFFF", name="Arial", size=11)
    center_align = Alignment(horizontal="center", vertical="center")
    thin_border  = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin"),
    )
    col_widths = [22, 22, 22, 18, 45]
    for col_idx, (header, width) in enumerate(zip(HEADERS, col_widths), start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = center_align
        cell.border    = thin_border
        ws.column_dimensions[cell.column_letter].width = width
    ws.row_dimensions[1].height = 20
    wb.save(SPREADSHEET_PATH)


def _local_append(moderator: str, username: str, strike_number: int, reason: str):
    _init_spreadsheet()
    wb = openpyxl.load_workbook(SPREADSHEET_PATH)
    ws = wb.active
    next_row = ws.max_row + 1
    for r in range(2, ws.max_row + 2):
        if ws.cell(row=r, column=1).value is None:
            next_row = r
            break
    row_data = [moderator, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), username, strike_number, reason]
    for col_idx, value in enumerate(row_data, start=1):
        _cell_style(ws, next_row, col_idx, value,
                    strike_num_col=(col_idx == 4), strike_num=strike_number)
    wb.save(SPREADSHEET_PATH)


# ── Primary API functions (Google Sheets first, local fallback) ───────────────

def _append_strike_row(moderator: str, username: str, strike_number: int, reason: str):
    """Append a strike to Google Sheets. Falls back to local xlsx on any error."""
    row_data = [
        moderator,
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        username,
        strike_number,
        reason,
    ]
    try:
        ws = _get_gsheet()
        ws.append_row(row_data, value_input_option="USER_ENTERED")
        # Colour the Strike column cell based on strike number
        all_vals  = ws.col_values(1)          # col A — find the row we just wrote
        new_row   = len(all_vals)             # last populated row
        strike_color = (
            {"red": 0.75, "green": 0.0, "blue": 0.0}   # red for strike 2
            if strike_number == 2
            else {"red": 1.0, "green": 0.4, "blue": 0.0}  # orange for strike 1
        )
        ws.format(f"D{new_row}", {
            "textFormat": {"bold": True, "foregroundColor": strike_color}
        })
        print(f"[Sheets] Strike row written to Google Sheet (row {new_row})")
    except Exception as e:
        print(f"[Sheets] Google Sheets write failed ({e}), writing to local fallback.")
        _local_append(moderator, username, strike_number, reason)


def _get_strikes_for_user(target: discord.Member) -> dict:
    """Read strike entries for a user from Google Sheets, fallback to local xlsx."""
    target_str = str(target).lower()

    def _parse_rows(rows):
        strike1_rows, strike2_rows = [], []
        for row in rows:
            if len(row) < 5:
                continue
            moderator, date, username, strike_num_raw, reason = row[:5]
            if not username:
                continue
            if str(username).lower() != target_str:
                continue
            try:
                strike_num = int(strike_num_raw)
            except (ValueError, TypeError):
                continue
            entry = {
                "date":      str(date)      if date      else "Unknown",
                "moderator": str(moderator) if moderator else "Unknown",
                "reason":    str(reason)    if reason    else "No reason given",
            }
            if strike_num == 1:
                strike1_rows.append(entry)
            elif strike_num == 2:
                strike2_rows.append(entry)
        return {
            "strike1": strike1_rows if strike1_rows else None,
            "strike2": strike2_rows if strike2_rows else None,
        }

    try:
        ws   = _get_gsheet()
        rows = ws.get_all_values()[1:]   # skip header row
        print(f"[Sheets] Read {len(rows)} rows from Google Sheet")
        return _parse_rows(rows)
    except Exception as e:
        print(f"[Sheets] Google Sheets read failed ({e}), reading from local fallback.")
        _init_spreadsheet()
        wb   = openpyxl.load_workbook(SPREADSHEET_PATH, data_only=True)
        rows = [[cell for cell in row] for row in wb.active.iter_rows(min_row=2, values_only=True)]
        return _parse_rows(rows)


def _appeal_strike(target: discord.Member, appeal_num: int) -> dict:
    """
    Remove a strike from Google Sheets (or local fallback).
    Logic mirrors the original: delete the row, demote Strike 2 → 1 if needed.
    """
    target_str = str(target).lower()

    # ── Try Google Sheets ─────────────────────────────────────────────────────
    try:
        ws       = _get_gsheet()
        all_rows = ws.get_all_values()   # includes header at index 0
        header   = all_rows[0]
        data     = all_rows[1:]          # 0-indexed; sheet row = index + 2

        strike1_rows, strike2_rows = [], []
        for i, row in enumerate(data):
            if len(row) < 5 or not row[2]:
                continue
            if str(row[2]).lower() != target_str:
                continue
            try:
                snum = int(row[3])
            except (ValueError, TypeError):
                continue
            entry = {
                "sheet_row": i + 2,      # 1-indexed sheet row number
                "moderator": row[0] or "Unknown",
                "date":      row[1] or "Unknown",
                "reason":    row[4] or "No reason given",
            }
            if snum == 1:
                strike1_rows.append(entry)
            elif snum == 2:
                strike2_rows.append(entry)

        if appeal_num == 1 and not strike1_rows:
            return {"success": False, "message": f"{target.mention} has no Strike 1 on record.",
                    "deleted_row": None, "demoted_row": None}
        if appeal_num == 2 and not strike2_rows:
            return {"success": False, "message": f"{target.mention} has no Strike 2 on record.",
                    "deleted_row": None, "demoted_row": None}

        deleted_row = demoted_row = None

        if appeal_num == 2:
            target_entry = strike2_rows[-1]
            deleted_row  = target_entry
            ws.delete_rows(target_entry["sheet_row"])
            print(f"[Sheets] Deleted Strike 2 row {target_entry['sheet_row']}")

        else:  # appeal_num == 1
            target_entry = strike1_rows[-1]
            deleted_row  = target_entry
            ws.delete_rows(target_entry["sheet_row"])
            print(f"[Sheets] Deleted Strike 1 row {target_entry['sheet_row']}")

            if strike2_rows:
                s2      = strike2_rows[-1]
                # Row index shifts by -1 after deletion if it was below the deleted row
                s2_row  = s2["sheet_row"] - (1 if s2["sheet_row"] > target_entry["sheet_row"] else 0)
                ws.update_cell(s2_row, 4, 1)   # column D = strike number
                ws.format(f"D{s2_row}", {
                    "textFormat": {
                        "bold": True,
                        "foregroundColor": {"red": 1.0, "green": 0.4, "blue": 0.0}  # orange
                    }
                })
                demoted_row = {"moderator": s2["moderator"], "date": s2["date"], "reason": s2["reason"]}
                print(f"[Sheets] Demoted Strike 2 → Strike 1 at row {s2_row}")

        return {"success": True, "message": "Appeal processed.",
                "deleted_row": deleted_row, "demoted_row": demoted_row}

    except Exception as e:
        print(f"[Sheets] Google Sheets appeal failed ({e}), falling back to local xlsx.")

    # ── Local xlsx fallback ───────────────────────────────────────────────────
    _init_spreadsheet()
    wb         = openpyxl.load_workbook(SPREADSHEET_PATH)
    ws_local   = wb.active
    strike1_rows, strike2_rows = [], []

    for row_idx in range(2, ws_local.max_row + 1):
        vals = [ws_local.cell(row=row_idx, column=c).value for c in range(1, 6)]
        moderator, date, username, strike_num, reason = vals
        if username is None:
            continue
        if str(username).lower() != target_str:
            continue
        entry = {
            "row":       row_idx,
            "moderator": str(moderator) if moderator else "Unknown",
            "date":      str(date)      if date      else "Unknown",
            "reason":    str(reason)    if reason    else "No reason given",
        }
        try:
            snum = int(strike_num)
        except (ValueError, TypeError):
            continue
        if snum == 1:
            strike1_rows.append(entry)
        elif snum == 2:
            strike2_rows.append(entry)

    if appeal_num == 1 and not strike1_rows:
        return {"success": False, "message": f"{target.mention} has no Strike 1 on record.",
                "deleted_row": None, "demoted_row": None}
    if appeal_num == 2 and not strike2_rows:
        return {"success": False, "message": f"{target.mention} has no Strike 2 on record.",
                "deleted_row": None, "demoted_row": None}

    deleted_row = demoted_row = None

    if appeal_num == 2:
        target_row  = strike2_rows[-1]
        deleted_row = target_row
        ws_local.delete_rows(target_row["row"])
    else:
        target_row  = strike1_rows[-1]
        deleted_row = target_row
        ws_local.delete_rows(target_row["row"])
        if strike2_rows:
            s2     = strike2_rows[-1]
            s2_row = s2["row"] - (1 if s2["row"] > target_row["row"] else 0)
            ws_local.cell(row=s2_row, column=4, value=1)
            _cell_style(ws_local, s2_row, 4, 1, strike_num_col=True, strike_num=1)
            demoted_row = {"moderator": s2["moderator"], "date": s2["date"], "reason": s2["reason"]}

    for r in range(2, ws_local.max_row + 1):
        if ws_local.cell(row=r, column=1).value is None:
            break
        for c in range(1, 6):
            fill_color = "D6E4F0" if (r % 2 == 0) else "FFFFFF"
            ws_local.cell(row=r, column=c).fill = PatternFill("solid", start_color=fill_color, end_color=fill_color)

    wb.save(SPREADSHEET_PATH)
    return {"success": True, "message": "Appeal processed.",
            "deleted_row": deleted_row, "demoted_row": demoted_row}


# ── Bot events ────────────────────────────────────────────────────────────────

@bot.event
async def on_ready():
    guild = discord.Object(id=GUILD_ID)
    tree.copy_global_to(guild=guild)
    await tree.sync(guild=guild)
    print(f"Logged in as {bot.user} | Slash commands synced to guild {GUILD_ID}")


# ── /strike ───────────────────────────────────────────────────────────────────

@tree.command(
    name="strike",
    description="Give a user a strike.",
    guild=discord.Object(id=GUILD_ID),
)
@app_commands.describe(username="The server member to strike", reason="Reason for the strike")
@app_commands.checks.has_permissions(moderate_members=True)
async def strike(interaction: discord.Interaction, username: discord.Member, reason: str):
    await interaction.response.defer(ephemeral=True)
    guild        = interaction.guild
    strike1_role = guild.get_role(STRIKE_1_ROLE_ID)
    strike2_role = guild.get_role(STRIKE_2_ROLE_ID)

    if not strike1_role or not strike2_role:
        await interaction.followup.send("❌ Strike roles not found. Check role IDs in `bot.py`.", ephemeral=True)
        return
    if strike2_role in username.roles:
        await interaction.followup.send(f"⚠️ {username.mention} already has **Strike 2**. No action taken.", ephemeral=True)
        return

    if strike1_role in username.roles:
        await username.add_roles(strike2_role, reason=reason)
        strike_number, role_given = 2, strike2_role
    else:
        await username.add_roles(strike1_role, reason=reason)
        strike_number, role_given = 1, strike1_role

    moderator_name = str(interaction.user)
    _append_strike_row(moderator=moderator_name, username=str(username),
                       strike_number=strike_number, reason=reason)

    colour = discord.Color.red() if strike_number == 2 else discord.Color.orange()
    embed  = discord.Embed(title=f"⚠️ Strike {strike_number} Issued", colour=colour,
                           timestamp=datetime.now(timezone.utc))
    embed.add_field(name="User",       value=username.mention,       inline=True)
    embed.add_field(name="Strike",     value=f"**{strike_number}**", inline=True)
    embed.add_field(name="Role Given", value=role_given.mention,     inline=True)
    embed.add_field(name="Reason",     value=reason,                 inline=False)
    embed.set_footer(text=f"Issued by {moderator_name}")

    log_channel = guild.get_channel(LOG_CHANNEL_ID)
    if not log_channel:
        await interaction.followup.send("❌ Log channel not found. Check `LOG_CHANNEL_ID`.", ephemeral=True)
        return
    try:
        await log_channel.send(embed=embed)
    except discord.Forbidden:
        await interaction.followup.send(
            f"✅ Strike {strike_number} applied but I can't post in <#{LOG_CHANNEL_ID}> "
            f"(missing Send Messages / Embed Links).", ephemeral=True)
        return

    await interaction.followup.send(
        f"✅ **Strike {strike_number}** applied to {username.mention} and logged.", ephemeral=True)


@strike.error
async def strike_error(interaction: discord.Interaction, error: app_commands.AppCommandError):
    msg = ("❌ You don't have permission to use `/strike`."
           if isinstance(error, app_commands.MissingPermissions)
           else f"❌ Unexpected error: {error}")
    try:
        await _safe_send(interaction)(msg, ephemeral=True)
    except Exception:
        pass


# ── /viewstrikes ──────────────────────────────────────────────────────────────

@tree.command(
    name="viewstrikes",
    description="View strike history for a user. Leave username blank to check your own.",
    guild=discord.Object(id=GUILD_ID),
)
@app_commands.describe(username="The member to look up (leave blank to check yourself)")
async def viewstrikes(interaction: discord.Interaction, username: discord.Member = None):
    await interaction.response.defer(ephemeral=True)
    invoker       = interaction.user
    guild         = interaction.guild
    can_view_role = guild.get_role(CAN_VIEW_ALL_ROLE)
    target        = username or invoker
    is_self       = (target.id == invoker.id)
    has_view_role = can_view_role in invoker.roles if can_view_role else False

    if not is_self and not has_view_role:
        fail_embed = discord.Embed(title="🚫 Unauthorized /viewstrikes Attempt",
                                   colour=discord.Color.dark_red(),
                                   timestamp=datetime.now(timezone.utc))
        fail_embed.add_field(name="Attempted by", value=invoker.mention, inline=True)
        fail_embed.add_field(name="Target user",  value=target.mention,  inline=True)
        fail_embed.set_footer(text="Failed permission check")
        await _post_to_log(guild, fail_embed)
        await interaction.followup.send(
            "❌ You can only view your own strikes. "
            "You need the moderator role to view another member's strikes.", ephemeral=True)
        return

    data            = _get_strikes_for_user(target)
    strike1_entries = data["strike1"]
    strike2_entries = data["strike2"]

    colour = (discord.Color.red()    if strike2_entries else
              discord.Color.orange() if strike1_entries else
              discord.Color.green())

    embed = discord.Embed(title=f"📋 Strike Record — {target.display_name}",
                          colour=colour, timestamp=datetime.now(timezone.utc))
    embed.set_thumbnail(url=target.display_avatar.url)

    if strike1_entries:
        for i, entry in enumerate(strike1_entries, start=1):
            label = "Strike 1" if len(strike1_entries) == 1 else f"Strike 1 (#{i})"
            embed.add_field(name=f"🟠 {label}",
                            value=f"**Date:** {entry['date']}\n**Moderator:** {entry['moderator']}\n**Reason:** {entry['reason']}",
                            inline=False)
    else:
        embed.add_field(name="🟠 Strike 1", value="N/A", inline=False)

    if strike2_entries:
        for i, entry in enumerate(strike2_entries, start=1):
            label = "Strike 2" if len(strike2_entries) == 1 else f"Strike 2 (#{i})"
            embed.add_field(name=f"🔴 {label}",
                            value=f"**Date:** {entry['date']}\n**Moderator:** {entry['moderator']}\n**Reason:** {entry['reason']}",
                            inline=False)
    else:
        embed.add_field(name="🔴 Strike 2", value="N/A", inline=False)

    embed.set_footer(text=f"Requested by {str(invoker)}")
    await interaction.followup.send(embed=embed, ephemeral=True)


@viewstrikes.error
async def viewstrikes_error(interaction: discord.Interaction, error: app_commands.AppCommandError):
    try:
        await _safe_send(interaction)(f"❌ Unexpected error: {error}", ephemeral=True)
    except Exception:
        pass


# ── /strikeappeal ─────────────────────────────────────────────────────────────

@tree.command(
    name="strikeappeal",
    description="Appeal (remove) a strike from a user. Only available to moderators.",
    guild=discord.Object(id=GUILD_ID),
)
@app_commands.describe(
    username="The member whose strike is being appealed",
    strike_number="Which strike to appeal (1 or 2)",
)
async def strikeappeal(interaction: discord.Interaction,
                       username: discord.Member,
                       strike_number: int):
    await interaction.response.defer(ephemeral=True)
    invoker        = interaction.user
    guild          = interaction.guild
    can_view_role  = guild.get_role(CAN_VIEW_ALL_ROLE)
    has_permission = can_view_role in invoker.roles if can_view_role else False

    if not has_permission:
        fail_embed = discord.Embed(title="🚫 Unauthorized /strikeappeal Attempt",
                                   colour=discord.Color.dark_red(),
                                   timestamp=datetime.now(timezone.utc))
        fail_embed.add_field(name="Attempted by",    value=invoker.mention,    inline=True)
        fail_embed.add_field(name="Target user",     value=username.mention,   inline=True)
        fail_embed.add_field(name="Strike appealed", value=str(strike_number), inline=True)
        fail_embed.set_footer(text="Failed permission check")
        await _post_to_log(guild, fail_embed)
        await interaction.followup.send(
            "❌ You don't have permission to use `/strikeappeal`.", ephemeral=True)
        return

    if strike_number not in (1, 2):
        await interaction.followup.send("❌ `strike_number` must be **1** or **2**.", ephemeral=True)
        return

    result = _appeal_strike(username, strike_number)
    if not result["success"]:
        await interaction.followup.send(f"⚠️ {result['message']}", ephemeral=True)
        return

    strike1_role = guild.get_role(STRIKE_1_ROLE_ID)
    strike2_role = guild.get_role(STRIKE_2_ROLE_ID)

    if strike_number == 2:
        if strike2_role and strike2_role in username.roles:
            await username.remove_roles(strike2_role, reason=f"Strike 2 appealed by {invoker}")
    else:
        if result["demoted_row"]:
            if strike2_role and strike2_role in username.roles:
                await username.remove_roles(strike2_role,
                    reason=f"Strike 1 appealed, Strike 2 demoted to Strike 1 by {invoker}")
        else:
            if strike1_role and strike1_role in username.roles:
                await username.remove_roles(strike1_role, reason=f"Strike 1 appealed by {invoker}")

    deleted = result["deleted_row"]
    demoted = result["demoted_row"]
    embed   = discord.Embed(title=f"✅ Strike {strike_number} Appealed",
                            colour=discord.Color.green(), timestamp=datetime.now(timezone.utc))
    embed.add_field(name="User",            value=username.mention,   inline=True)
    embed.add_field(name="Strike Appealed", value=str(strike_number), inline=True)
    embed.add_field(name="Appealed by",     value=invoker.mention,    inline=True)
    if deleted:
        embed.add_field(name="Removed Entry",
                        value=f"**Date:** {deleted['date']}\n**Originally issued by:** {deleted['moderator']}\n**Reason:** {deleted['reason']}",
                        inline=False)
    if demoted:
        embed.add_field(name="⬇️ Strike 2 → Strike 1 (demoted)",
                        value=f"**Date:** {demoted['date']}\n**Originally issued by:** {demoted['moderator']}\n**Reason:** {demoted['reason']}",
                        inline=False)

    await _post_to_log(guild, embed)
    await interaction.followup.send(
        f"✅ Strike {strike_number} for {username.mention} has been appealed and the log updated.",
        ephemeral=True)


@strikeappeal.error
async def strikeappeal_error(interaction: discord.Interaction, error: app_commands.AppCommandError):
    try:
        await _safe_send(interaction)(f"❌ Unexpected error: {error}", ephemeral=True)
    except Exception:
        pass


# ── /place ────────────────────────────────────────────────────────────────────
#
# Discord slash commands support only ONE attachment parameter natively.
# To accept up to 10 files, we declare file1–file10 as separate optional
# parameters. Discord presents them as individual upload slots in the UI.
# Files are NEVER saved to disk — we only read their .url from Discord's CDN.

@tree.command(
    name="place",
    description="Assign a role to a user with supporting files as evidence. Files are logged, never stored locally.",
    guild=discord.Object(id=GUILD_ID),
)
@app_commands.describe(
    username  = "The member to assign the role to",
    role      = "The role to give the user",
    file1     = "Evidence file 1 (required)",
    file2     = "Evidence file 2",
    file3     = "Evidence file 3",
    file4     = "Evidence file 4",
    file5     = "Evidence file 5",
    file6     = "Evidence file 6",
    file7     = "Evidence file 7",
    file8     = "Evidence file 8",
    file9     = "Evidence file 9",
    file10    = "Evidence file 10",
)
async def place(
    interaction : discord.Interaction,
    username    : discord.Member,
    role        : discord.Role,
    file1       : discord.Attachment,
    file2       : discord.Attachment = None,
    file3       : discord.Attachment = None,
    file4       : discord.Attachment = None,
    file5       : discord.Attachment = None,
    file6       : discord.Attachment = None,
    file7       : discord.Attachment = None,
    file8       : discord.Attachment = None,
    file9       : discord.Attachment = None,
    file10      : discord.Attachment = None,
):
    await interaction.response.defer(ephemeral=False)
    invoker        = interaction.user
    guild          = interaction.guild

    # ── Channel restriction ───────────────────────────────────────────────────
    if interaction.channel_id != PLACEMENT_CMD_CHANNEL:
        await interaction.followup.send(
            f"❌ This command can only be used in <#{PLACEMENT_CMD_CHANNEL}>.",
            ephemeral=True)
        return
    can_view_role  = guild.get_role(CAN_VIEW_ALL_ROLE)
    placement_role = guild.get_role(PLACEMENT_STAFF)
    has_permission = (
        (can_view_role  and can_view_role  in invoker.roles) or
        (placement_role and placement_role in invoker.roles)
    )

    # ── Permission check ──────────────────────────────────────────────────────
    if not has_permission:
        fail_embed = discord.Embed(
            title="🚫 Unauthorized /place Attempt",
            colour=discord.Color.dark_red(),
            timestamp=datetime.now(timezone.utc),
        )
        fail_embed.add_field(name="Attempted by", value=invoker.mention,  inline=True)
        fail_embed.add_field(name="Target user",  value=username.mention, inline=True)
        fail_embed.add_field(name="Role",         value=role.mention,     inline=True)
        fail_embed.set_footer(text="Failed permission check")
        await _post_to_log(guild, fail_embed)
        await interaction.followup.send(
            "❌ You don't have permission to use `/place`.", ephemeral=True)
        return

    # ── Assign role ───────────────────────────────────────────────────────────
    if role in username.roles:
        await interaction.followup.send(
            f"⚠️ {username.mention} already has {role.mention}. No action taken.")
        return

    try:
        await username.add_roles(role, reason=f"/place used by {invoker}")
    except discord.Forbidden:
        await interaction.followup.send(
            f"❌ I don't have permission to assign {role.mention}. "
            f"Make sure my role is above it in the hierarchy.", ephemeral=True)
        return

    # ── Collect attachment URLs (never touch the files themselves) ────────────
    attachments = [a for a in [file1, file2, file3, file4, file5,
                                file6, file7, file8, file9, file10] if a is not None]

    # ── Build the log embed ───────────────────────────────────────────────────
    place_channel = guild.get_channel(PLACE_CHANNEL_ID)
    if not place_channel:
        await interaction.followup.send(
            f"✅ Role {role.mention} assigned to {username.mention}, but "
            f"`PLACE_CHANNEL_ID` is not set or the channel wasn't found.", ephemeral=True)
        return

    embed = discord.Embed(
        title="📁 Role Placement",
        colour=role.colour if role.colour.value != 0 else discord.Color.blurple(),
        timestamp=datetime.now(timezone.utc),
    )
    embed.add_field(name="User",         value=username.mention, inline=True)
    embed.add_field(name="Role Assigned", value=role.mention,    inline=True)
    embed.add_field(name="Issued by",    value=invoker.mention,  inline=True)
    embed.add_field(
        name=f"Evidence ({len(attachments)} file{'s' if len(attachments) != 1 else ''})",
        value="\n".join(
            f"[{a.filename}]({a.url})" for a in attachments
        ),
        inline=False,
    )
    embed.set_footer(text="Files are hosted on Discord's CDN — not stored locally.")

    # ── If the first file is an image, set it as the embed image ─────────────
    image_exts = {".png", ".jpg", ".jpeg", ".gif", ".webp"}
    for a in attachments:
        ext = os.path.splitext(a.filename)[1].lower()
        if ext in image_exts:
            embed.set_image(url=a.url)
            break

    try:
        await place_channel.send(embed=embed)
    except discord.Forbidden:
        await interaction.followup.send(
            f"✅ Role assigned but I can't post in <#{PLACE_CHANNEL_ID}> "
            f"(missing Send Messages / Embed Links).", ephemeral=True)
        return

    await interaction.followup.send(
        f"✅ {role.mention} assigned to {username.mention} and logged with "
        f"{len(attachments)} file{'s' if len(attachments) != 1 else ''}.",
    )


@place.error
async def place_error(interaction: discord.Interaction, error: app_commands.AppCommandError):
    try:
        await _safe_send(interaction)(f"❌ Unexpected error: {error}", ephemeral=True)
    except Exception:
        pass


# ── /dupcategory ──────────────────────────────────────────────────────────────

@tree.command(
    name="dupcategory",
    description="Duplicate all channels from an existing category into a new one.",
    guild=discord.Object(id=GUILD_ID),
)
@app_commands.describe(
    old_category_id    = "The ID of the category to duplicate",
    new_category_name  = "Name for the new category",
    new_category_index = "Position (index) of the new category in the channel list (0 = top)",
)
async def dupcategory(
    interaction        : discord.Interaction,
    old_category_id    : str,
    new_category_name  : str,
    new_category_index : str,
):
    await interaction.response.defer(ephemeral=True)
    invoker = interaction.user
    guild   = interaction.guild

    # ── Validate inputs ───────────────────────────────────────────────────────
    try:
        old_category_id = int(old_category_id)
    except ValueError:
        await interaction.followup.send(
            f"❌ `{old_category_id}` is not a valid ID. Paste the raw number, e.g. `1519168841454850089`.",
            ephemeral=True)
        return

    try:
        new_category_index = int(new_category_index)
    except ValueError:
        await interaction.followup.send(
            f"❌ `{new_category_index}` is not a valid position index. Please enter a whole number.",
            ephemeral=True)
        return

    # ── Permission check ──────────────────────────────────────────────────────
    dup_role_1 = guild.get_role(DUP_ROLE_1)
    dup_role_2 = guild.get_role(DUP_ROLE_2)
    dup_role_3 = guild.get_role(DUP_ROLE_3)
    has_permission = any(
        role and role in invoker.roles
        for role in (dup_role_1, dup_role_2, dup_role_3)
    )

    if not has_permission:
        await interaction.followup.send(
            "❌ You don't have permission to use `/dupcategory`.", ephemeral=True)
        return

    # ── Locate the source category ────────────────────────────────────────────
    source_category = guild.get_channel(old_category_id)

    if source_category is None:
        await interaction.followup.send(
            f"❌ No channel found with ID `{old_category_id}`. "
            f"Make sure you're using the category's ID, not its name.", ephemeral=True)
        return

    if not isinstance(source_category, discord.CategoryChannel):
        await interaction.followup.send(
            f"❌ The channel with ID `{old_category_id}` is not a category.", ephemeral=True)
        return

    # ── Create the new category ───────────────────────────────────────────────
    try:
        new_category = await guild.create_category(
            name     = new_category_name,
            position = new_category_index,
            overwrites = source_category.overwrites,  # copy permission overwrites
            reason   = f"/dupcategory used by {invoker}",
        )
    except discord.Forbidden:
        await interaction.followup.send(
            "❌ I don't have permission to create categories.", ephemeral=True)
        return

    # ── Duplicate each channel in the source category ─────────────────────────
    # Sort by position so order is preserved in the new category
    channels       = sorted(source_category.channels, key=lambda c: c.position)
    created        = []
    failed         = []

    for ch in channels:
        try:
            if isinstance(ch, discord.TextChannel):
                new_ch = await guild.create_text_channel(
                    name       = ch.name,
                    category   = new_category,
                    position   = ch.position,
                    topic      = ch.topic,
                    slowmode_delay = ch.slowmode_delay,
                    nsfw       = ch.nsfw,
                    overwrites = ch.overwrites,
                    reason     = f"/dupcategory — copied from #{ch.name}",
                )
            elif isinstance(ch, discord.VoiceChannel):
                new_ch = await guild.create_voice_channel(
                    name       = ch.name,
                    category   = new_category,
                    position   = ch.position,
                    bitrate    = ch.bitrate,
                    user_limit = ch.user_limit,
                    overwrites = ch.overwrites,
                    reason     = f"/dupcategory — copied from #{ch.name}",
                )
            elif isinstance(ch, discord.StageChannel):
                new_ch = await guild.create_stage_channel(
                    name       = ch.name,
                    category   = new_category,
                    position   = ch.position,
                    overwrites = ch.overwrites,
                    reason     = f"/dupcategory — copied from #{ch.name}",
                )
            elif isinstance(ch, discord.ForumChannel):
                new_ch = await guild.create_forum(
                    name       = ch.name,
                    category   = new_category,
                    position   = ch.position,
                    topic      = ch.topic,
                    slowmode_delay = ch.slowmode_delay,
                    nsfw       = ch.nsfw,
                    overwrites = ch.overwrites,
                    reason     = f"/dupcategory — copied from #{ch.name}",
                )
            else:
                # Announcement / News / unknown — fall back to text channel
                new_ch = await guild.create_text_channel(
                    name       = ch.name,
                    category   = new_category,
                    position   = ch.position,
                    overwrites = ch.overwrites,
                    reason     = f"/dupcategory — copied from #{ch.name} (fallback)",
                )
            created.append(new_ch)

        except Exception as e:
            failed.append((ch.name, str(e)))

    # ── Respond ───────────────────────────────────────────────────────────────
    summary_lines = [
        f"✅ New category **{new_category.name}** created at position {new_category_index}.",
        f"**{len(created)}/{len(channels)}** channel(s) duplicated successfully.",
    ]
    if failed:
        summary_lines.append(
            "⚠️ The following channels could not be duplicated:\n" +
            "\n".join(f"• `#{name}` — {err}" for name, err in failed)
        )

    await interaction.followup.send("\n".join(summary_lines), ephemeral=True)


@dupcategory.error
async def dupcategory_error(interaction: discord.Interaction, error: app_commands.AppCommandError):
    try:
        await _safe_send(interaction)(f"❌ Unexpected error: {error}", ephemeral=True)
    except Exception:
        pass


# ── /skipplace ────────────────────────────────────────────────────────────────

@tree.command(
    name="skipplace",
    description="Assign a role to a returning user who was previously placed and has rejoined.",
    guild=discord.Object(id=GUILD_ID),
)
@app_commands.describe(
    username = "The returning member to assign the role to",
    role     = "The role to give the user",
    file1    = "Evidence file (e.g. screenshot of previous placement)",
)
async def skipplace(
    interaction : discord.Interaction,
    username    : discord.Member,
    role        : discord.Role,
    file1       : discord.Attachment,
):
    await interaction.response.defer(ephemeral=False)
    invoker        = interaction.user
    guild          = interaction.guild

    # ── Channel restriction ───────────────────────────────────────────────────
    if interaction.channel_id != PLACEMENT_CMD_CHANNEL:
        await interaction.followup.send(
            f"❌ This command can only be used in <#{PLACEMENT_CMD_CHANNEL}>.",
            ephemeral=True)
        return
    can_view_role  = guild.get_role(CAN_VIEW_ALL_ROLE)
    placement_role = guild.get_role(PLACEMENT_STAFF)
    has_permission = (
        (can_view_role  and can_view_role  in invoker.roles) or
        (placement_role and placement_role in invoker.roles)
    )

    # ── Permission check ──────────────────────────────────────────────────────
    if not has_permission:
        fail_embed = discord.Embed(
            title="🚫 Unauthorized /skipplace Attempt",
            colour=discord.Color.dark_red(),
            timestamp=datetime.now(timezone.utc),
        )
        fail_embed.add_field(name="Attempted by", value=invoker.mention,  inline=True)
        fail_embed.add_field(name="Target user",  value=username.mention, inline=True)
        fail_embed.add_field(name="Role",         value=role.mention,     inline=True)
        fail_embed.set_footer(text="Failed permission check")
        await _post_to_log(guild, fail_embed)
        await interaction.followup.send(
            "❌ You don't have permission to use `/skipplace`.", ephemeral=True)
        return

    # ── Assign role ───────────────────────────────────────────────────────────
    if role in username.roles:
        await interaction.followup.send(
            f"⚠️ {username.mention} already has {role.mention}. No action taken.")
        return

    try:
        await username.add_roles(role, reason=f"/skipplace (returning user) used by {invoker}")
    except discord.Forbidden:
        await interaction.followup.send(
            f"❌ I don't have permission to assign {role.mention}. "
            f"Make sure my role is above it in the hierarchy.", ephemeral=True)
        return

    # ── Post to place channel ─────────────────────────────────────────────────
    place_channel = guild.get_channel(PLACE_CHANNEL_ID)
    if not place_channel:
        await interaction.followup.send(
            f"✅ Role {role.mention} assigned to {username.mention}, but "
            f"`PLACE_CHANNEL_ID` is not set or the channel wasn't found.", ephemeral=True)
        return

    embed = discord.Embed(
        title="⏭️ Placement Skipped — Returning User",
        colour=discord.Color.gold(),
        timestamp=datetime.now(timezone.utc),
    )
    embed.add_field(name="User",          value=username.mention, inline=True)
    embed.add_field(name="Role Assigned", value=role.mention,     inline=True)
    embed.add_field(name="Issued by",     value=invoker.mention,  inline=True)
    embed.add_field(
        name="ℹ️ Reason for Skip",
        value="This user has been placed before and has since rejoined the server. "
              "Placement was skipped and their role was restored directly.",
        inline=False,
    )
    embed.add_field(
        name="Evidence",
        value=f"[{file1.filename}]({file1.url})",
        inline=False,
    )

    # Embed image preview if the file is an image
    image_exts = {".png", ".jpg", ".jpeg", ".gif", ".webp"}
    if os.path.splitext(file1.filename)[1].lower() in image_exts:
        embed.set_image(url=file1.url)

    embed.set_footer(text="Files are hosted on Discord's CDN — not stored locally.")

    try:
        await place_channel.send(embed=embed)
    except discord.Forbidden:
        await interaction.followup.send(
            f"✅ Role assigned but I can't post in <#{PLACE_CHANNEL_ID}> "
            f"(missing Send Messages / Embed Links).", ephemeral=True)
        return

    await interaction.followup.send(
        f"✅ {role.mention} restored to returning user {username.mention} and logged.",
    )


@skipplace.error
async def skipplace_error(interaction: discord.Interaction, error: app_commands.AppCommandError):
    try:
        await _safe_send(interaction)(f"❌ Unexpected error: {error}", ephemeral=True)
    except Exception:
        pass


# ── Run ───────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    bot.run(BOT_TOKEN)