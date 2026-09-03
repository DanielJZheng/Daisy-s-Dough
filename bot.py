from __future__ import annotations

import logging
from dataclasses import dataclass
from datetime import datetime

import discord
from discord.ext import commands

from categories import CategoryRegistry
from config import Settings, load_settings
from parser import ParsedEntry, parse_message
from sheets import SheetManager

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("daisys-dough")

CONFIRM_EMOJI = "\N{WHITE HEAVY CHECK MARK}"
REJECT_EMOJI = "\N{CROSS MARK}"
WARN_EMOJI = "\N{WARNING SIGN}"
ERROR_EMOJI = "\N{CROSS MARK}"


@dataclass
class PendingConfirmation:
    original_message_id: int
    channel_id: int
    user_id: int
    entry: ParsedEntry
    raw_category: str
    suggested: str
    local_dt: datetime


class DaisysDoughBot(commands.Bot):
    def __init__(self, settings: Settings) -> None:
        intents = discord.Intents.default()
        intents.message_content = True
        intents.reactions = True
        intents.members = True
        super().__init__(command_prefix="!", intents=intents)
        self.settings = settings
        self.sheets = SheetManager(
            settings.google_sheet_id,
            str(settings.google_oauth_client_path),
            str(settings.google_token_path),
            settings.timezone,
        )
        self.categories = CategoryRegistry(
            self.sheets._spreadsheet,
            settings.fuzzy_match_cutoff,
        )
        self.pending: dict[int, PendingConfirmation] = {}

    async def setup_hook(self) -> None:
        logger.info("Daisy's Dough is ready as %s", self.user)

    async def on_message(self, message: discord.Message) -> None:
        if message.author.bot:
            return
        if message.channel.id != self.settings.discord_channel_id:
            return

        entry = parse_message(message.content)
        if entry is None:
            if self.settings.react_on_error:
                await message.add_reaction(ERROR_EMOJI)
            return

        local_dt = self.sheets.localize(message.created_at)
        display_name = message.author.display_name or message.author.name
        match = self.categories.resolve(entry.category, entry.entry_type)

        if match.kind == "exact" and match.canonical:
            await self._write_entry(message, entry, match.canonical, local_dt)
            return

        if match.kind == "fuzzy" and match.suggested:
            await self._prompt_confirmation(message, entry, match.suggested, local_dt)
            return

        canonical = self.categories.add_category(
            entry.category,
            entry.entry_type,
            first_seen=local_dt.date(),
        )
        await self._write_entry(message, entry, canonical, local_dt)

    async def _prompt_confirmation(
        self,
        message: discord.Message,
        entry: ParsedEntry,
        suggested: str,
        local_dt: datetime,
    ) -> None:
        await message.add_reaction(WARN_EMOJI)
        reply = await message.reply(
            f"Did you mean **{suggested}**? (you typed: {entry.category})"
        )
        await reply.add_reaction(CONFIRM_EMOJI)
        await reply.add_reaction(REJECT_EMOJI)
        self.pending[reply.id] = PendingConfirmation(
            original_message_id=message.id,
            channel_id=message.channel.id,
            user_id=message.author.id,
            entry=entry,
            raw_category=entry.category,
            suggested=suggested,
            local_dt=local_dt,
        )

    async def _write_entry(
        self,
        message: discord.Message,
        entry: ParsedEntry,
        category: str,
        local_dt: datetime,
    ) -> None:
        display_name = message.author.display_name or message.author.name
        worksheet = self.sheets.get_or_create_user_worksheet(
            local_dt,
            display_name,
            message.author.id,
        )
        self.sheets.upsert_entry(
            worksheet,
            local_dt,
            entry,
            category,
            message.id,
        )
        await message.add_reaction(CONFIRM_EMOJI)

    async def on_raw_reaction_add(self, payload: discord.RawReactionActionEvent) -> None:
        if payload.user_id == self.user.id:
            return
        pending = self.pending.get(payload.message_id)
        if pending is None:
            return
        if payload.user_id != pending.user_id:
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

        if emoji == CONFIRM_EMOJI:
            canonical = self.categories.add_category(
                pending.suggested,
                pending.entry.entry_type,
                first_seen=pending.local_dt.date(),
            )
            self.categories.add_alias(
                canonical,
                pending.raw_category,
                pending.entry.entry_type,
            )
            category = canonical
        else:
            category = self.categories.add_category(
                pending.raw_category,
                pending.entry.entry_type,
                first_seen=pending.local_dt.date(),
            )

        display_name = original.author.display_name or original.author.name
        worksheet = self.sheets.get_or_create_user_worksheet(
            pending.local_dt,
            display_name,
            original.author.id,
        )
        self.sheets.upsert_entry(
            worksheet,
            pending.local_dt,
            pending.entry,
            category,
            original.id,
        )
        await original.add_reaction(CONFIRM_EMOJI)
        self.pending.pop(payload.message_id, None)


def main() -> None:
    settings = load_settings()
    bot = DaisysDoughBot(settings)
    bot.run(settings.discord_token)


if __name__ == "__main__":
    main()
