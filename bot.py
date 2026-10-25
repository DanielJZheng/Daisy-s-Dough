from __future__ import annotations

import asyncio
import logging
from dataclasses import dataclass, replace
from datetime import datetime

import discord
from discord.ext import commands, tasks

from config import Settings, load_settings
from google_errors import format_google_sheets_error
from parser import (
    FORMAT_HELP,
    ParsedEntry,
    SudoTarget,
    looks_like_entry,
    parse_entries,
    parse_recover_request,
    parse_refresh_request,
    parse_section_add,
)
from sheet_queue import SheetWriteQueue
from sheets import SHOPPING_WORD, SheetManager, resolve_keyword

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("daisys-dough")

CONFIRM_EMOJI = "\N{WHITE HEAVY CHECK MARK}"
REJECT_EMOJI = "\N{CROSS MARK}"
WARN_EMOJI = "\N{WARNING SIGN}"
ERROR_EMOJI = "\N{CROSS MARK}"
QUEUE_EMOJI = "\N{HOURGLASS WITH FLOWING SAND}"


def confirmation_logs_entry(emoji: str) -> bool:
    """✅ accepts the suggestion and writes. ❌ declines and does not write."""
    return emoji == CONFIRM_EMOJI


@dataclass
class PendingConfirmation:
    original_message_id: int
    channel_id: int
    user_id: int
    entry: ParsedEntry
    suggested: str
    category: str
    local_dt: datetime
    target_user_id: int
    target_display_name: str


class DaisysDoughBot(commands.Bot):
    def __init__(self, settings: Settings) -> None:
        intents = discord.Intents.default()
        intents.message_content = True
        super().__init__(command_prefix="!", intents=intents)
        self.settings = settings
        self.sheets = SheetManager.from_settings(settings)
        self._sheet_queue = SheetWriteQueue(
            min_interval=settings.sheets_min_interval_sec,
        )
        self.pending: dict[int, PendingConfirmation] = {}

    async def setup_hook(self) -> None:
        await asyncio.to_thread(self.sheets.ensure_foundation)
        self.stamp_open_ledgers.start()
        logger.info("Daisy's Dough is ready as %s", self.user)

    @tasks.loop(seconds=300.0)
    async def stamp_open_ledgers(self) -> None:
        await self._enqueue_stamp_poll()

    @stamp_open_ledgers.before_loop
    async def _before_stamp_loop(self) -> None:
        await self.wait_until_ready()

    async def _enqueue_stamp_poll(self) -> None:
        if not self.sheets._ledger_cache:
            return

        async def on_success() -> None:
            return

        async def on_failure(error: Exception) -> None:
            if "429" in str(error):
                logger.warning(
                    "Sheets read quota (429) — skipped timestamp poll; will retry later"
                )
                return
            logger.exception("Manual ledger sync failed")

        await self._sheet_queue.submit(
            self.sheets.backfill_open_ledger_tabs,
            on_success=on_success,
            on_failure=on_failure,
        )

    async def on_message(self, message: discord.Message) -> None:
        if message.author.bot:
            return
        if message.channel.id != self.settings.discord_channel_id:
            return

        if parse_refresh_request(message.content):
            await self._refresh_ledger(message)
            return

        if parse_recover_request(message.content):
            await self._recover_ledger(message)
            return

        section_name = parse_section_add(message.content)
        if section_name is not None:
            await self._add_section(message, section_name)
            return

        sudo, entries = parse_entries(message.content)
        if not entries:
            if looks_like_entry(message.content):
                await self._send_help(message)
            return

        if sudo is not None:
            target = await self._target_for(message, sudo)
            if target is None:
                await message.reply(
                    "Mention them so the ledger can open: `sudo @name` then one transaction per line",
                    mention_author=False,
                )
                return
            user_id, display_name = target
        else:
            user_id = message.author.id
            display_name = message.author.display_name or message.author.name

        if len(entries) > 1:
            await self._write_many(message, entries, user_id, display_name)
            return

        entry = entries[0]
        local_dt = self.sheets.localize(message.created_at)
        if (
            entry.description.lower() == SHOPPING_WORD
            or entry.entry_type == "income"
        ):
            await self._write_entry(
                message, entry, local_dt, user_id, display_name
            )
            return

        try:
            match = await asyncio.to_thread(self._resolve_keyword, entry.description)
        except Exception as error:
            logger.exception("Keyword lookup failed for message %s", message.id)
            if self.settings.react_on_error:
                await message.add_reaction(ERROR_EMOJI)
            await message.reply(
                format_google_sheets_error(error),
                mention_author=False,
            )
            return

        if match.kind == "exact" and match.category == "Shopping":
            await self._write_entry(
                message,
                replace(entry, description=SHOPPING_WORD),
                local_dt,
                user_id,
                display_name,
            )
            return

        if match.kind == "fuzzy" and match.keyword and match.category:
            await self._prompt_confirmation(
                message,
                entry,
                match.keyword,
                match.category,
                local_dt,
                user_id,
                display_name,
            )
            return

        await self._write_entry(message, entry, local_dt, user_id, display_name)

    async def _write_many(
        self,
        message: discord.Message,
        entries: list[ParsedEntry],
        user_id: int,
        display_name: str,
    ) -> None:
        local_dt = self.sheets.localize(message.created_at)
        count = len(entries)

        def _sync_write() -> None:
            worksheet = self.sheets.get_or_create_user_worksheet(
                display_name,
                user_id,
                local_dt,
            )
            for index, entry in enumerate(entries):
                self.sheets.append_entry(
                    worksheet,
                    local_dt,
                    entry,
                    message.id * 100 + index,
                    defer_refresh=True,
                )
            self.sheets.flush_refresh(worksheet)

        async def on_success() -> None:
            await self._clear_queue_reaction(message)
            await message.add_reaction(CONFIRM_EMOJI)
            await message.reply(
                f"Logged {count} for **{display_name}**.",
                mention_author=False,
            )

        await self._enqueue_sheet(message, _sync_write, on_success=on_success)

    async def _send_help(self, message: discord.Message) -> None:
        if self.settings.react_on_error:
            await message.add_reaction(ERROR_EMOJI)
        await message.reply(FORMAT_HELP, mention_author=False)

    async def _target_for(
        self,
        message: discord.Message,
        sudo: SudoTarget,
    ) -> tuple[int, str] | None:
        if sudo.user_id is not None:
            mentioned = next(
                (user for user in message.mentions if user.id == sudo.user_id),
                None,
            )
            if mentioned is None:
                return sudo.user_id, str(sudo.user_id)
            return mentioned.id, mentioned.display_name or mentioned.name
        if sudo.name is None:
            return None
        if message.guild is not None:
            member = message.guild.get_member_named(sudo.name)
            if member is not None:
                return member.id, member.display_name or member.name
        return await asyncio.to_thread(self.sheets.find_owner_by_name, sudo.name)

    def _resolve_keyword(self, description: str):
        rows = list(self.sheets.keywords())
        if not any(keyword.lower() == SHOPPING_WORD for keyword, _ in rows):
            rows.append((SHOPPING_WORD, "Shopping"))
        return resolve_keyword(
            description,
            rows,
            self.settings.fuzzy_match_cutoff,
        )

    async def _prompt_confirmation(
        self,
        message: discord.Message,
        entry: ParsedEntry,
        suggested: str,
        category: str,
        local_dt: datetime,
        target_user_id: int,
        target_display_name: str,
    ) -> None:
        await message.add_reaction(WARN_EMOJI)
        reply = await message.reply(
            f"Did you mean **{suggested}**? (you typed: {entry.description})"
        )
        self.pending[reply.id] = PendingConfirmation(
            original_message_id=message.id,
            channel_id=message.channel.id,
            user_id=message.author.id,
            entry=entry,
            suggested=suggested,
            category=category,
            local_dt=local_dt,
            target_user_id=target_user_id,
            target_display_name=target_display_name,
        )
        await reply.add_reaction(CONFIRM_EMOJI)
        await reply.add_reaction(REJECT_EMOJI)

    async def _write_entry(
        self,
        message: discord.Message,
        entry: ParsedEntry,
        local_dt: datetime,
        user_id: int,
        display_name: str,
    ) -> None:
        sudo_reply = user_id != message.author.id

        def _sync_write() -> None:
            worksheet = self.sheets.get_or_create_user_worksheet(
                display_name,
                user_id,
                local_dt,
            )
            self.sheets.append_entry(
                worksheet,
                local_dt,
                entry,
                message.id,
            )

        async def on_success() -> None:
            await self._clear_queue_reaction(message)
            await message.add_reaction(CONFIRM_EMOJI)
            if sudo_reply:
                await message.reply(
                    f"Logged for **{display_name}**.",
                    mention_author=False,
                )

        await self._enqueue_sheet(message, _sync_write, on_success=on_success)

    async def _recover_ledger(self, message: discord.Message) -> None:
        local_dt = self.sheets.localize(message.created_at)
        display_name = message.author.display_name or message.author.name
        outcome: dict[str, int] = {"restored": 0}

        def _sync_recover() -> None:
            worksheet = self.sheets.get_or_create_user_worksheet(
                display_name,
                message.author.id,
                local_dt,
            )
            outcome["restored"] = self.sheets.restore_visible_from_raw(worksheet)
            self.sheets._refresh_ledger_views(worksheet)

        async def on_success() -> None:
            await self._clear_queue_reaction(message)
            await message.add_reaction(CONFIRM_EMOJI)
            count = outcome["restored"]
            if count:
                text = f"Restored **{count}** transaction(s) into Recurring / One-off."
            else:
                text = (
                    "No RAW backup rows found to restore. "
                    "Try Google Sheets **File → Version history** for older edits."
                )
            await message.reply(text, mention_author=False)

        await self._enqueue_sheet(message, _sync_recover, on_success=on_success)

    async def _refresh_ledger(self, message: discord.Message) -> None:
        local_dt = self.sheets.localize(message.created_at)
        display_name = message.author.display_name or message.author.name

        outcome: dict[str, int] = {"restored": 0}

        def _sync_refresh() -> None:
            _worksheet, restored = self.sheets.refresh_user_ledger(
                display_name,
                message.author.id,
                local_dt,
            )
            outcome["restored"] = restored

        async def on_success() -> None:
            await self._clear_queue_reaction(message)
            await message.add_reaction(CONFIRM_EMOJI)
            extra = ""
            if outcome["restored"]:
                extra = (
                    f" Restored **{outcome['restored']}** row(s) from hidden RAW backup."
                )
            await message.reply(
                f"Refreshed **{display_name}**'s ledger tab.{extra}",
                mention_author=False,
            )

        await self._enqueue_sheet(message, _sync_refresh, on_success=on_success)

    async def _add_section(self, message: discord.Message, name: str) -> None:
        local_dt = self.sheets.localize(message.created_at)
        display_name = message.author.display_name or message.author.name

        def _sync_add() -> int:
            worksheet = self.sheets.get_or_create_user_worksheet(
                display_name,
                message.author.id,
                local_dt,
            )
            return self.sheets.add_section(worksheet, name)

        async def on_success() -> None:
            await self._clear_queue_reaction(message)
            await message.add_reaction(CONFIRM_EMOJI)
            await message.reply(
                f"Added ledger section **{name.strip()}**.",
                mention_author=False,
            )

        await self._enqueue_sheet(message, _sync_add, on_success=on_success)

    async def _clear_queue_reaction(self, message: discord.Message) -> None:
        if self.user is None:
            return
        try:
            await message.remove_reaction(QUEUE_EMOJI, self.user)
        except discord.HTTPException:
            pass

    async def _enqueue_sheet(
        self,
        message: discord.Message,
        action,
        *,
        on_success,
    ) -> None:
        try:
            await message.add_reaction(QUEUE_EMOJI)
        except discord.HTTPException:
            pass

        async def on_failure(error: Exception) -> None:
            await self._clear_queue_reaction(message)
            logger.exception("Failed sheet write for message %s", message.id)
            if self.settings.react_on_error:
                await message.add_reaction(ERROR_EMOJI)
            await message.reply(
                format_google_sheets_error(error),
                mention_author=False,
            )

        await self._sheet_queue.submit(
            action,
            on_success=on_success,
            on_failure=on_failure,
        )

    async def on_raw_reaction_add(self, payload: discord.RawReactionActionEvent) -> None:
        if self.user is None or payload.user_id == self.user.id:
            return
        pending = self.pending.get(payload.message_id)
        if pending is None or payload.user_id != pending.user_id:
            return

        emoji = str(payload.emoji)
        if emoji not in {CONFIRM_EMOJI, REJECT_EMOJI}:
            return

        channel = self.get_channel(payload.channel_id)
        if not isinstance(channel, (discord.TextChannel, discord.Thread)):
            return

        try:
            original = await channel.fetch_message(pending.original_message_id)
        except discord.NotFound:
            self.pending.pop(payload.message_id, None)
            return

        if not confirmation_logs_entry(emoji):
            self.pending.pop(payload.message_id, None)
            if self.settings.react_on_error:
                try:
                    await original.add_reaction(ERROR_EMOJI)
                except discord.HTTPException:
                    pass
            return

        def _sync_write() -> None:
            entry = pending.entry
            if pending.suggested.lower() == SHOPPING_WORD:
                self.sheets.add_keyword(entry.description, "Shopping")
                entry = replace(entry, description=SHOPPING_WORD)
            else:
                self.sheets.add_keyword(entry.description, pending.category)
            worksheet = self.sheets.get_or_create_user_worksheet(
                pending.target_display_name,
                pending.target_user_id,
                pending.local_dt,
            )
            self.sheets.append_entry(
                worksheet,
                pending.local_dt,
                entry,
                original.id,
            )

        pending_id = payload.message_id

        async def on_success() -> None:
            await self._clear_queue_reaction(original)
            await original.add_reaction(CONFIRM_EMOJI)
            self.pending.pop(pending_id, None)

        async def on_failure(error: Exception) -> None:
            await self._clear_queue_reaction(original)
            logger.exception("Failed to confirm entry for message %s", original.id)
            if self.settings.react_on_error:
                await original.add_reaction(ERROR_EMOJI)
            await original.reply(
                format_google_sheets_error(error),
                mention_author=False,
            )
            self.pending.pop(pending_id, None)

        try:
            await original.add_reaction(QUEUE_EMOJI)
        except discord.HTTPException:
            pass
        await self._sheet_queue.submit(
            _sync_write,
            on_success=on_success,
            on_failure=on_failure,
        )


def main() -> None:
    settings = load_settings()
    bot = DaisysDoughBot(settings)
    bot.run(settings.discord_token)


if __name__ == "__main__":
    main()
