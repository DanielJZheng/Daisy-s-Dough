from __future__ import annotations

import difflib
import re
from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

import gspread
from gspread.exceptions import APIError, WorksheetNotFound
from gspread.utils import ValueInputOption, ValueRenderOption, rowcol_to_a1

from google_auth import authorize_google
from parser import ParsedEntry
from sheets_retry import with_sheets_retry

import ledger_v6

LEDGER_PREFIX = "Ledger - "
CATEGORIES_SHEET = "Categories"
LEGACY_CATEGORIES_SHEET = "Categories (legacy)"
SUMMARY_SHEET = "Summary"

HEADER_ROW = 1
FIRST_DATA_ROW = 2
INITIAL_TOTALS_ROW = 4

TX_TS = 1
TX_AMT = 2
TX_CAT = 3
TX_NOTE = 4
GAP_COL = 5
ONEOFF_TS = 6
ONEOFF_AMT = 7
ONEOFF_CAT = 8
ONEOFF_NOTE = 9
GAP_COL_2 = 10
SHOP_TS = 11
SHOP_AMT = 12
SHOP_NOTE = 13
GAP_COL_3 = 14
INC_TS = 15
INC_AMT = 16
INC_CAT = 17
INC_NOTE = 18
GAP_COL_4 = 19

# v4 sheets stored these 5 columns to the left, before the one-off insert.
OLD_OWNER_COL = 80
OLD_LAYOUT_COL = 82
META_OWNER = 85
META_TOTALS = 86
META_LAYOUT = 87
TX_TYPE_COL = 88
TX_MSG_COL = 89
SHOP_MSG_COL = 90
INC_MSG_COL = 91
ONEOFF_MSG_COL = 92
RAW_TS = 94
RAW_AMT = 95
RAW_CAT = 96
RAW_NOTE = 97
RAW_MSG = 98

LAYOUT_VERSION = "ledger-v6"
LAYOUT_V5 = "ledger-v5"
LAYOUT_V4 = "ledger-v4"
TOTALS_START = 15
SHOPPING_WORD = "shopping"
TOTAL_SHOPPING = "Shopping"
TOTAL_INCOME = "Income"
TOTAL_NET = "Net"
TAIL_LABELS = (TOTAL_SHOPPING, TOTAL_INCOME, TOTAL_NET)

LEDGER_HEADER_BG = "#7D5A3C"
LEDGER_ZONE_EXPENSE = "#FFF8F0"
LEDGER_ZONE_ONEOFF = "#E6D5C4"
LEDGER_ZONE_SHOP = "#F5EBDD"
LEDGER_ZONE_INCOME = "#E6F0E4"
LEDGER_TOTALS_BG = "#D9CFC0"
LEDGER_SUMMARY_BG = "#EDE4D8"
LEDGER_TOTALS_COL_WIDTH = 150
LEDGER_FORMAT_LAST_ROW = 1000
LEDGER_META_FORMAT_END = META_OWNER - 1
SUMMARY_FOOTER_OFFSET = 3
LEDGER_SUMMARY_SCRUB_TAIL = 12
LEDGER_SUMMARY_FORMAT_END = 24

TX_HEADERS = ["Timestamp", "Amount", "Category", "Note"]
ONEOFF_HEADERS = ["Timestamp", "Amount", "Category", "Note"]
SHOP_HEADERS = ["Timestamp", "Amount", "Note"]
INC_HEADERS = ["Timestamp", "Amount", "Category", "Note"]
MOVED_ONEOFFS = frozenset({"first day tw", "shopee", "indo", "mom"})

CATEGORY_HEADERS = ["Keyword", "Category"]
CATEGORY_SEEDS = [
    ["seven", "Convenience"],
    ["burger", "Food"],
    ["water", "Food"],
    ["coco", "Food"],
]

SUMMARY_BLOCK_WIDTH = 3
SUMMARY_QUERY_ROW = 7

INVALID_TAB_CHARS = re.compile(r"[\[\]\\/?*:]")
TIMESTAMP_LIKE = re.compile(r"^\d{4}-\d{2}-\d{2}")


@dataclass(frozen=True)
class KeywordMatch:
    kind: str
    keyword: str | None = None
    category: str | None = None


def resolve_keyword(
    raw: str,
    rows: list[tuple[str, str]],
    cutoff: float,
) -> KeywordMatch:
    text = raw.strip()
    if not text:
        return KeywordMatch(kind="none")

    by_lower: dict[str, tuple[str, str]] = {}
    for keyword, category in rows:
        name = keyword.strip()
        if name:
            by_lower[name.lower()] = (name, category.strip())

    direct = by_lower.get(text.lower())
    if direct:
        return KeywordMatch(kind="exact", keyword=direct[0], category=direct[1])

    if not by_lower:
        return KeywordMatch(kind="none")

    suggestions = difflib.get_close_matches(
        text.lower(),
        list(by_lower),
        n=1,
        cutoff=cutoff,
    )
    if not suggestions:
        return KeywordMatch(kind="none")

    keyword, category = by_lower[suggestions[0]]
    return KeywordMatch(kind="fuzzy", keyword=keyword, category=category)


def _sanitize_tab_part(value: str) -> str:
    return INVALID_TAB_CHARS.sub("", value).strip()


def tab_name(display_name: str, user_id: int) -> str:
    user = _sanitize_tab_part(display_name) or str(user_id)
    suffix = str(user_id)[-4:]
    name = f"{LEDGER_PREFIX}{user} ({suffix})"
    if len(name) > 100:
        name = name[:100]
    return name


# October stays on the current ledger. November and after get a new tab.
MONTHLY_LEDGER_START = (2026, 11)
LEDGER_MONTH_SUFFIX = re.compile(r" - (\d{4}-\d{2})$")


def ledger_period(when: datetime) -> str | None:
    if (when.year, when.month) < MONTHLY_LEDGER_START:
        return None
    return f"{when.year:04d}-{when.month:02d}"


def ledger_title_period(title: str) -> str | None:
    match = LEDGER_MONTH_SUFFIX.search(title)
    return match.group(1) if match else None


def category_stem(name: str) -> str:
    key = re.sub(r"\s+", " ", name).strip().casefold()
    if len(key) > 4 and key.endswith("ies"):
        return key[:-3] + "y"
    if len(key) > 3 and key.endswith("es") and not key.endswith("ses"):
        return key[:-2]
    if len(key) > 3 and key.endswith("s") and not key.endswith("ss"):
        return key[:-1]
    return key


def preferred_category(labels: list[str]) -> str:
    cleaned = [re.sub(r"\s+", " ", label).strip() for label in labels if label and label.strip()]
    if not cleaned:
        return ""
    counts = Counter(cleaned)

    def rank(label: str) -> tuple[bool, int, int, bool]:
        key = label.casefold()
        plural = key.endswith("s") and not key.endswith("ss")
        titled = any(char.isupper() for char in label)
        return (plural, len(key), counts[label], titled)

    return max(cleaned, key=rank)


def _sheet_ref(title: str) -> str:
    escaped = title.replace("'", "''")
    return f"'{escaped}'"


def _owner_token(user_id: int) -> str:
    return f"owner:{user_id}"


def _as_int(value: Any) -> int | None:
    if value in (None, ""):
        return None
    try:
        return int(float(str(value).strip()))
    except (TypeError, ValueError):
        return None


def _parse_amount(value: Any) -> float | None:
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace("$", "").replace(",", "").strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _message_id(value: Any, fallback: int) -> int:
    digits = "".join(char for char in str(value or "") if char.isdigit())
    if not digits:
        return fallback
    return int(digits)


def _summary_math(ref: str) -> tuple[str, str]:
    expense_amt = _col_letters(TX_AMT)
    shop_amt = _col_letters(SHOP_AMT)
    income_amt = _col_letters(INC_AMT)
    income = f"=SUM({ref}!{income_amt}:{income_amt})"
    expenses = (
        f"=SUM({ref}!{expense_amt}:{expense_amt})"
        f"+SUM({ref}!{shop_amt}:{shop_amt})"
    )
    return income, expenses


def _col_letters(col: int) -> str:
    return rowcol_to_a1(1, col)[:-1]


def _section_band_for_ts(timestamp_col: int) -> tuple[int, int] | None:
    bands = {
        TX_TS: (TX_TS, TX_NOTE),
        ONEOFF_TS: (ONEOFF_TS, ONEOFF_NOTE),
        SHOP_TS: (SHOP_TS, SHOP_NOTE),
        INC_TS: (INC_TS, INC_NOTE),
    }
    return bands.get(timestamp_col)


def _sheet_rgb(hex_color: str) -> dict[str, float]:
    text = hex_color.lstrip("#")
    return {
        "red": int(text[0:2], 16) / 255,
        "green": int(text[2:4], 16) / 255,
        "blue": int(text[4:6], 16) / 255,
    }


def _ledger_repeat_cell(
    sheet_id: int,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
    cell_format: dict[str, Any],
    fields: str,
) -> dict[str, Any]:
    return {
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": start_row - 1,
                "endRowIndex": end_row,
                "startColumnIndex": start_col - 1,
                "endColumnIndex": end_col,
            },
            "cell": {"userEnteredFormat": cell_format},
            "fields": fields,
        }
    }


def total_slot(index: int) -> int:
    return TOTALS_START + index


def _category_display_columns() -> list[int]:
    columns: list[int] = []
    for start, end in (
        (TX_TS, TX_NOTE),
        (ONEOFF_TS, ONEOFF_NOTE),
        (SHOP_TS, SHOP_NOTE),
        (INC_TS, INC_NOTE),
    ):
        columns.extend(range(start, end + 1))
    return columns


def _ledger_column_segments(start_col: int, end_col: int) -> list[tuple[int, int]]:
    gaps = {GAP_COL, GAP_COL_2, GAP_COL_3, GAP_COL_4}
    segments: list[tuple[int, int]] = []
    segment_start: int | None = None
    for column in range(start_col, end_col + 1):
        if column in gaps:
            if segment_start is not None:
                segments.append((segment_start, column - 1))
                segment_start = None
            continue
        if segment_start is None:
            segment_start = column
    if segment_start is not None:
        segments.append((segment_start, end_col))
    return segments


class SheetManager:
    def __init__(
        self,
        sheet_id: str,
        oauth_client_path: str,
        token_path: str,
        timezone: str,
        *,
        service_account_path: str | None = None,
        service_account_info: dict | None = None,
    ) -> None:
        sa_path = Path(service_account_path) if service_account_path else None
        client = authorize_google(
            Path(oauth_client_path),
            Path(token_path),
            service_account_path=sa_path,
            service_account_info=service_account_info,
        )
        self._spreadsheet = client.open_by_key(sheet_id)
        self._tz = ZoneInfo(timezone)
        self._keyword_cache: list[tuple[str, str]] | None = None
        self._refresh_pending: set[int] = set()
        self._refresh_pending_row: dict[int, int] = {}
        self._foundation_ready = False
        self._logged_message_ids: set[tuple[int, str]] = set()
        self._summary_linked: set[str] = set()
        self._ledger_cache: dict[tuple[int, str | None], gspread.Worksheet] = {}
        self._stamp_backfill_at: dict[int, float] = {}

    @classmethod
    def from_settings(cls, settings: object) -> SheetManager:
        return cls(
            settings.google_sheet_id,
            str(settings.google_oauth_client_path),
            str(settings.google_token_path),
            settings.timezone,
            service_account_path=(
                str(settings.google_service_account_path)
                if settings.google_service_account_path
                else None
            ),
            service_account_info=settings.google_service_account_info,
        )

    def backfill_manual_timestamps(
        self,
        worksheet: gspread.Worksheet,
        when: datetime | None = None,
        *,
        force: bool = False,
    ) -> int:
        if self._cell_value(worksheet, rowcol_to_a1(1, META_LAYOUT)) != LAYOUT_VERSION:
            return 0
        import time

        now = time.monotonic()
        if not force and now - self._stamp_backfill_at.get(worksheet.id, 0.0) < 120.0:
            return 0
        moment = when or datetime.now(self._tz)
        if moment.tzinfo is None:
            moment = moment.replace(tzinfo=self._tz)
        else:
            moment = moment.astimezone(self._tz)
        count = ledger_v6.backfill_missing_timestamps(
            self, worksheet, moment=moment
        )
        self._stamp_backfill_at[worksheet.id] = now
        return count

    def backfill_open_ledger_tabs(self) -> None:
        """Stamp empty times and rebalance Recurring/One-off on tabs this process already opened."""
        if not self._ledger_cache:
            return
        seen: set[int] = set()
        for worksheet in self._ledger_cache.values():
            if worksheet.id in seen:
                continue
            seen.add(worksheet.id)
            try:
                self._refresh_ledger_views(worksheet, format_sheet=False)
            except APIError as error:
                if "429" in str(error):
                    raise
                continue
            except Exception:
                continue

    def localize(self, dt: datetime) -> datetime:
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=ZoneInfo("UTC"))
        return dt.astimezone(self._tz)

    def ensure_foundation(self) -> None:
        if self._foundation_ready:
            return
        self._ensure_categories()
        self._ensure_summary_sheet()
        self._foundation_ready = True

    def keywords(self) -> list[tuple[str, str]]:
        if self._keyword_cache is not None:
            return self._keyword_cache

        self.ensure_foundation()
        worksheet = self._spreadsheet.worksheet(CATEGORIES_SHEET)
        values = with_sheets_retry(lambda: worksheet.get_all_values())
        rows: list[tuple[str, str]] = []
        for row in values[1:]:
            if not row or not row[0].strip():
                continue
            keyword = row[0].strip()
            category = row[1].strip() if len(row) > 1 else ""
            rows.append((keyword, category))
        self._keyword_cache = rows
        return rows

    def add_keyword(self, keyword: str, category: str) -> None:
        keyword = keyword.strip()
        category = category.strip()
        if not keyword or not category:
            return
        if any(existing.lower() == keyword.lower() for existing, _ in self.keywords()):
            return

        worksheet = self._spreadsheet.worksheet(CATEGORIES_SHEET)
        with_sheets_retry(lambda: worksheet.append_row([keyword, category]))
        self._keyword_cache = None

    def get_or_create_user_worksheet(
        self,
        display_name: str,
        user_id: int,
        when: datetime | None = None,
    ) -> gspread.Worksheet:
        self.ensure_foundation()
        moment = when or datetime.now(self._tz)
        if moment.tzinfo is None:
            moment = moment.replace(tzinfo=self._tz)
        else:
            moment = moment.astimezone(self._tz)
        period = ledger_period(moment)
        cache_key = (user_id, period)
        worksheet = self._ledger_cache.get(cache_key)
        if worksheet is None:
            worksheet = self._find_ledger(user_id, period)
            if worksheet is None:
                worksheet = self._open_ledger_by_title(
                    display_name, user_id, period
                )
            if worksheet is not None:
                self._ledger_cache[cache_key] = worksheet
        if worksheet is None:
            worksheet = self._create_ledger(display_name, user_id, period)
            self._ledger_cache[cache_key] = worksheet
        self._ensure_summary_block(worksheet.title, display_name)
        return worksheet

    def is_message_logged(self, worksheet: gspread.Worksheet, message_id: int) -> bool:
        token = str(message_id)
        key = (worksheet.id, token)
        if key in self._logged_message_ids:
            return True
        totals_row = _as_int(self._cell_value(worksheet, rowcol_to_a1(1, META_TOTALS)))
        end_row = (totals_row - 2) if totals_row else LEDGER_FORMAT_LAST_ROW
        end_row = max(end_row, FIRST_DATA_ROW)
        block = with_sheets_retry(
            lambda: worksheet.get(
                f"{rowcol_to_a1(FIRST_DATA_ROW, SHOP_MSG_COL)}:"
                f"{rowcol_to_a1(end_row, RAW_MSG)}",
            )
        )
        for row in block or []:
            for cell in row:
                if str(cell).strip() == token:
                    self._logged_message_ids.add(key)
                    return True
        return False

    def _mark_message_logged(self, worksheet: gspread.Worksheet, message_id: int) -> None:
        self._logged_message_ids.add((worksheet.id, str(message_id)))

    def flush_refresh(self, worksheet: gspread.Worksheet) -> None:
        if worksheet.id not in self._refresh_pending:
            return
        self._refresh_pending.discard(worksheet.id)
        hint = self._refresh_pending_row.pop(worksheet.id, None)
        self._refresh_ledger_views(worksheet, data_row_hint=hint)

    def refresh_user_ledger(
        self,
        display_name: str,
        user_id: int,
        when: datetime | None = None,
    ) -> tuple[gspread.Worksheet, int]:
        worksheet = self.get_or_create_user_worksheet(display_name, user_id, when)
        restored = self.restore_visible_from_raw(worksheet)
        self.backfill_manual_timestamps(worksheet, when=when, force=True)
        self._refresh_ledger_views(worksheet)
        return worksheet, restored

    def _expense_fingerprint(
        self,
        stamp: str,
        amount: float,
        category: str,
        note: str,
    ) -> tuple[float, str, str, str]:
        stamp_key = stamp.strip()[:19]
        return (
            round(amount, 2),
            category.casefold(),
            note.casefold(),
            stamp_key,
        )

    def _visible_expense_counts(
        self, worksheet: gspread.Worksheet
    ) -> Counter[tuple[float, str, str, str]]:
        counts: Counter[tuple[float, str, str, str]] = Counter()
        totals = self._totals_row(worksheet)
        last_row = max(totals - 1, ledger_v6.FIRST_DATA_ROW)
        block = with_sheets_retry(
            lambda: worksheet.get(
                f"{rowcol_to_a1(ledger_v6.FIRST_DATA_ROW, TX_TS)}:"
                f"{rowcol_to_a1(last_row, ONEOFF_NOTE)}"
            )
        )
        for cells in block or []:
            padded = list(cells) + [""] * (ONEOFF_NOTE - TX_TS + 1 - len(cells))
            for base in (TX_TS, ONEOFF_TS):
                offset = base - TX_TS
                stamp = str(padded[offset] or "").strip()
                amount = _parse_amount(padded[offset + 1])
                category = str(padded[offset + 2] or "").strip()
                note = str(padded[offset + 3] or "").strip()
                if amount is None or not category:
                    continue
                if stamp.casefold() == "timestamp" or category.casefold() == "category":
                    continue
                counts[self._expense_fingerprint(stamp, amount, category, note)] += 1
        return counts

    def _logged_message_ids_on_ledger(
        self, worksheet: gspread.Worksheet
    ) -> set[str]:
        ids: set[str] = set()
        for col in (TX_MSG_COL, ONEOFF_MSG_COL):
            values = with_sheets_retry(lambda c=col: worksheet.col_values(c))
            for text in values:
                token = str(text or "").strip()
                if token.isdigit():
                    ids.add(token)
        return ids

    def _last_raw_data_row(self, worksheet: gspread.Worksheet) -> int:
        last = ledger_v6.FIRST_DATA_ROW - 1
        for col in (RAW_TS, RAW_AMT, RAW_CAT):
            values = with_sheets_retry(lambda c=col: worksheet.col_values(c))
            for index in range(len(values) - 1, 0, -1):
                if str(values[index] or "").strip():
                    last = max(last, index + 1)
                    break
        return last

    def _iter_raw_expense_rows(
        self, worksheet: gspread.Worksheet
    ) -> list[tuple[int, str, float, str, str, str]]:
        last_row = self._last_raw_data_row(worksheet)
        if last_row < ledger_v6.FIRST_DATA_ROW:
            return []
        block = with_sheets_retry(
            lambda: worksheet.get(
                f"{rowcol_to_a1(ledger_v6.FIRST_DATA_ROW, RAW_TS)}:"
                f"{rowcol_to_a1(last_row, RAW_MSG)}"
            )
        )
        rows: list[tuple[int, str, float, str, str, str]] = []
        for offset, cells in enumerate(block or []):
            row_num = ledger_v6.FIRST_DATA_ROW + offset
            padded = list(cells) + [""] * (RAW_MSG - RAW_TS + 1 - len(cells))
            stamp = ledger_v6.format_stamp(padded[0])
            amount = _parse_amount(padded[RAW_AMT - RAW_TS])
            category = str(padded[RAW_CAT - RAW_TS] or "").strip()
            note = str(padded[RAW_NOTE - RAW_TS] or "").strip()
            message_id = str(padded[RAW_MSG - RAW_TS] or "").strip()
            if stamp.casefold() == "timestamp" or category.casefold() == "category":
                continue
            if amount is None or not category:
                continue
            rows.append((row_num, stamp, amount, category, note, message_id))
        return rows

    def restore_visible_from_raw(self, worksheet: gspread.Worksheet) -> int:
        """Copy hidden RAW rows back into Recurring / One-off when a refresh wiped them."""
        layout = self._cell_value(worksheet, rowcol_to_a1(1, META_LAYOUT))
        if layout != LAYOUT_VERSION:
            return 0
        visible = self._visible_expense_counts(worksheet)
        rec_values: list[list[Any]] = []
        rec_msgs: list[list[str]] = []
        one_values: list[list[Any]] = []
        one_msgs: list[list[str]] = []
        stem_count: dict[str, int] = {}
        for col in (TX_CAT, ONEOFF_CAT):
            values = with_sheets_retry(lambda c=col: worksheet.col_values(c))
            for index, value in enumerate(values):
                if index == 0:
                    continue
                current = re.sub(r"\s+", " ", str(value or "")).strip()
                if not current:
                    continue
                stem = category_stem(current)
                stem_count[stem] = stem_count.get(stem, 0) + 1
        for _row, stamp, amount, category, note, message_id in self._iter_raw_expense_rows(
            worksheet
        ):
            fingerprint = self._expense_fingerprint(stamp, amount, category, note)
            if visible[fingerprint] > 0:
                visible[fingerprint] -= 1
                continue
            stem = category_stem(category)
            count = stem_count.get(stem, 0)
            stem_count[stem] = count + 1
            row_vals = [stamp, amount, category, note]
            msg_vals = [str(message_id)] if message_id else [""]
            if count >= 1:
                rec_values.append(row_vals)
                rec_msgs.append(msg_vals)
            else:
                one_values.append(row_vals)
                one_msgs.append(msg_vals)
            visible[fingerprint] += 1
        restored = len(rec_values) + len(one_values)
        if rec_values:
            dest = ledger_v6.next_open_row(self, worksheet, TX_TS, TX_NOTE)
            end = dest + len(rec_values) - 1
            self._write_values(
                worksheet,
                f"{rowcol_to_a1(dest, TX_TS)}:{rowcol_to_a1(end, TX_NOTE)}",
                rec_values,
                user_entered=True,
            )
            self._write_values(
                worksheet,
                f"{rowcol_to_a1(dest, TX_MSG_COL)}:{rowcol_to_a1(end, TX_MSG_COL)}",
                rec_msgs,
            )
        if one_values:
            dest = ledger_v6.next_open_row(self, worksheet, ONEOFF_TS, ONEOFF_NOTE)
            end = dest + len(one_values) - 1
            self._write_values(
                worksheet,
                f"{rowcol_to_a1(dest, ONEOFF_TS)}:{rowcol_to_a1(end, ONEOFF_NOTE)}",
                one_values,
                user_entered=True,
            )
            self._write_values(
                worksheet,
                f"{rowcol_to_a1(dest, ONEOFF_MSG_COL)}:{rowcol_to_a1(end, ONEOFF_MSG_COL)}",
                one_msgs,
            )
        return restored

    def append_entry(
        self,
        worksheet: gspread.Worksheet,
        local_dt: datetime,
        entry: ParsedEntry,
        message_id: int,
        *,
        stamp: str | None = None,
        defer_refresh: bool = False,
    ) -> None:
        if self.is_message_logged(worksheet, message_id):
            return

        stamp = stamp or local_dt.strftime("%Y-%m-%d %H:%M:%S")
        description = entry.description
        if entry.entry_type == "expense" and description.lower() != SHOPPING_WORD:
            description = self._canonical_expense_category(worksheet, description)
            entry = ParsedEntry(entry.entry_type, entry.amount, description, entry.note)
        header_row = (
            ledger_v6.HEADER_ROW
            if self._cell_value(worksheet, rowcol_to_a1(1, META_LAYOUT)) == LAYOUT_VERSION
            else HEADER_ROW
        )
        income_header = self._cell_value(worksheet, rowcol_to_a1(header_row, INC_TS))
        if entry.entry_type == "income" and income_header == TX_HEADERS[0]:
            row = self._next_data_row(worksheet, INC_TS)
            self._write_values(
                worksheet,
                f"{rowcol_to_a1(row, INC_TS)}:{rowcol_to_a1(row, INC_NOTE)}",
                [[stamp, entry.amount, entry.description, entry.note or ""]],
                user_entered=True,
            )
            self._write_values(
                worksheet,
                rowcol_to_a1(row, INC_MSG_COL),
                [[str(message_id)]],
            )
            self._mark_message_logged(worksheet, message_id)
            self._finish_append(worksheet, defer_refresh, data_row=row)
            return

        if entry.description.lower() == SHOPPING_WORD:
            row = self._next_data_row(worksheet, SHOP_TS)
            self._write_values(
                worksheet,
                f"{rowcol_to_a1(row, SHOP_TS)}:{rowcol_to_a1(row, SHOP_NOTE)}",
                [[stamp, entry.amount, entry.note or ""]],
                user_entered=True,
            )
            self._write_values(
                worksheet,
                rowcol_to_a1(row, SHOP_MSG_COL),
                [[str(message_id)]],
            )
            self._mark_message_logged(worksheet, message_id)
            self._finish_append(worksheet, defer_refresh, data_row=row)
            return

        layout = self._cell_value(worksheet, rowcol_to_a1(1, META_LAYOUT))
        if layout == LAYOUT_VERSION:
            ts, end_col, msg_col = self._expense_write_columns(
                worksheet, entry.description
            )
            row = self._next_data_row(worksheet, ts)
            self._write_values(
                worksheet,
                f"{rowcol_to_a1(row, ts)}:{rowcol_to_a1(row, end_col)}",
                [[stamp, entry.amount, entry.description, entry.note or ""]],
                user_entered=True,
            )
            self._write_values(
                worksheet,
                rowcol_to_a1(row, msg_col),
                [[str(message_id)]],
            )
            self._mark_message_logged(worksheet, message_id)
            self._finish_append(worksheet, defer_refresh, data_row=row)
            return
        if layout != LAYOUT_V5:
            row = self._next_data_row(worksheet, TX_TS)
            self._write_values(
                worksheet,
                f"{rowcol_to_a1(row, TX_TS)}:{rowcol_to_a1(row, TX_NOTE)}",
                [[stamp, entry.amount, entry.description, entry.note or ""]],
                user_entered=True,
            )
            self._write_values(
                worksheet,
                rowcol_to_a1(row, TX_MSG_COL),
                [[str(message_id)]],
            )
            self._mark_message_logged(worksheet, message_id)
            return

        row = self._next_data_row(worksheet, RAW_TS)
        self._write_values(
            worksheet,
            f"{rowcol_to_a1(row, RAW_TS)}:{rowcol_to_a1(row, RAW_NOTE)}",
            [[stamp, entry.amount, entry.description, entry.note or ""]],
            user_entered=True,
        )
        self._write_values(
            worksheet,
            rowcol_to_a1(row, RAW_MSG),
            [[str(message_id)]],
        )
        self._mark_message_logged(worksheet, message_id)
        self._finish_append(worksheet, defer_refresh, data_row=row)

    def _finish_append(
        self,
        worksheet: gspread.Worksheet,
        defer_refresh: bool,
        *,
        data_row: int | None = None,
    ) -> None:
        if defer_refresh:
            self._refresh_pending.add(worksheet.id)
            if data_row is not None:
                self._refresh_pending_row[worksheet.id] = max(
                    data_row,
                    self._refresh_pending_row.get(worksheet.id, FIRST_DATA_ROW),
                )
            return
        self._refresh_ledger_views(worksheet, data_row_hint=data_row)

    def _refresh_ledger_views(
        self,
        worksheet: gspread.Worksheet,
        *,
        data_row_hint: int | None = None,
        format_sheet: bool = True,
    ) -> None:
        layout = self._cell_value(worksheet, rowcol_to_a1(1, META_LAYOUT))
        if layout not in (LAYOUT_VERSION, LAYOUT_V5):
            return
        if layout == LAYOUT_V5:
            ledger_v6.migrate_to_v6(self, worksheet)
        end_row = self._last_entry_row_across_zones(worksheet, floor=data_row_hint)
        self.backfill_manual_timestamps(worksheet, force=True)
        scanned = self._last_entry_row_across_zones(worksheet)
        ledger_v6.rebalance_expense_bands(self, worksheet, max(end_row, scanned))
        end_row = self._last_entry_row_across_zones(worksheet)
        ledger_v6.clear_summary_tail(self, worksheet, end_row)
        ledger_v6.refresh_footer(self, worksheet, end_row)
        totals_row = self._totals_row(worksheet)
        self._clear_legacy_category_formulas(worksheet, totals_row)
        self._write_category_totals(worksheet, totals_row, end_row)
        if format_sheet:
            self._format_ledger_v6(worksheet)

    def _clear_raw_row(self, worksheet: gspread.Worksheet, row: int) -> None:
        blank = ["", "", "", "", ""]
        self._write_values(
            worksheet,
            f"{rowcol_to_a1(row, RAW_TS)}:{rowcol_to_a1(row, RAW_MSG)}",
            [blank],
        )

    def _move_raw_row_to_income(
        self,
        worksheet: gspread.Worksheet,
        raw_row: int,
        *,
        amount: float | None = None,
    ) -> int:
        stamp = self._cell_value(worksheet, rowcol_to_a1(raw_row, RAW_TS)) or ""
        parsed = amount if amount is not None else _parse_amount(
            self._cell_value(worksheet, rowcol_to_a1(raw_row, RAW_AMT))
        )
        category = self._cell_value(worksheet, rowcol_to_a1(raw_row, RAW_CAT)) or ""
        note = self._cell_value(worksheet, rowcol_to_a1(raw_row, RAW_NOTE)) or ""
        message_id = self._cell_value(worksheet, rowcol_to_a1(raw_row, RAW_MSG))
        if parsed is None or not category.strip():
            raise ValueError(f"Row {raw_row} is not a movable expense row")

        income_row = self._next_data_row(worksheet, INC_TS)
        self._write_values(
            worksheet,
            f"{rowcol_to_a1(income_row, INC_TS)}:{rowcol_to_a1(income_row, INC_NOTE)}",
            [[stamp, parsed, category, note]],
            user_entered=True,
        )
        if message_id:
            self._write_values(
                worksheet,
                rowcol_to_a1(income_row, INC_MSG_COL),
                [[str(message_id)]],
            )
        self._clear_raw_row(worksheet, raw_row)
        return income_row

    def _find_ledger(self, user_id: int, period: str | None = None) -> gspread.Worksheet | None:
        suffix = f"({str(user_id)[-4:]})"
        token = _owner_token(user_id)
        owned: list[gspread.Worksheet] = []
        for worksheet in self._spreadsheet.worksheets():
            title = worksheet.title
            if not title.startswith(LEDGER_PREFIX) or suffix not in title:
                continue
            if ledger_title_period(title) != period and not title.endswith(suffix):
                continue
            owner = self._cell_value(worksheet, rowcol_to_a1(1, META_OWNER))
            old_owner = self._cell_value(worksheet, rowcol_to_a1(1, OLD_OWNER_COL))
            legacy = self._cell_value(worksheet, "Z1")
            owner = owner or old_owner
            if owner == token or legacy == token:
                owned.append(worksheet)
        for worksheet in owned:
            if ledger_title_period(worksheet.title) != period:
                continue
            self._ensure_grid(worksheet)
            self._ensure_ledger_layout(worksheet, user_id)
            return worksheet
        return None

    def _ensure_owner_meta(self, worksheet: gspread.Worksheet, user_id: int) -> None:
        token = _owner_token(user_id)
        current = self._cell_value(worksheet, rowcol_to_a1(1, META_OWNER))
        if current != token:
            self._write_values(
                worksheet,
                rowcol_to_a1(1, META_OWNER),
                [[token]],
            )

    def _open_ledger_by_title(
        self,
        display_name: str,
        user_id: int,
        period: str | None,
    ) -> gspread.Worksheet | None:
        """Open an existing tab by name when owner meta was never set."""
        titles = {ws.title: ws for ws in self._spreadsheet.worksheets()}
        candidate = tab_name(display_name, user_id)
        if period:
            candidate = f"{candidate} - {period}"
        for title in (candidate, f"{LEDGER_PREFIX}{user_id}"[:100]):
            worksheet = titles.get(title)
            if worksheet is None:
                continue
            self._ensure_grid(worksheet)
            self._ensure_owner_meta(worksheet, user_id)
            self._ensure_ledger_layout(worksheet, user_id)
            return worksheet
        return None

    def _create_ledger(
        self,
        display_name: str,
        user_id: int,
        period: str | None = None,
    ) -> gspread.Worksheet:
        if period:
            title = f"{tab_name(display_name, user_id)} - {period}"
        else:
            title = self._unique_ledger_title(display_name, user_id)
        try:
            worksheet = self._spreadsheet.add_worksheet(title=title, rows=1000, cols=100)
        except APIError as error:
            if "already exists" not in str(error).lower():
                raise
            worksheet = self._spreadsheet.worksheet(title)
            self._ensure_ledger_layout(worksheet, user_id)
            return worksheet

        self._bootstrap_ledger(worksheet, user_id)
        return worksheet

    def find_owner_by_name(self, name: str) -> tuple[int, str] | None:
        wanted = _sanitize_tab_part(name).casefold()
        if not wanted:
            return None
        for worksheet in self._spreadsheet.worksheets():
            title = worksheet.title
            if not title.startswith(LEDGER_PREFIX):
                continue
            body = title[len(LEDGER_PREFIX):]
            display, sep, _suffix = body.rpartition(" (")
            if not sep:
                display = body
            if _sanitize_tab_part(display).casefold() != wanted:
                continue
            raw = (
                self._cell_value(worksheet, rowcol_to_a1(1, META_OWNER))
                or self._cell_value(worksheet, rowcol_to_a1(1, OLD_OWNER_COL))
                or self._cell_value(worksheet, "Z1")
            )
            if not raw or not str(raw).startswith("owner:"):
                continue
            try:
                user_id = int(str(raw).split(":", 1)[1])
            except ValueError:
                continue
            return user_id, display
        return None

    def _snapshot_rows(
        self, worksheet: gspread.Worksheet
    ) -> list[tuple[str, ParsedEntry, int]]:
        main_len = len(with_sheets_retry(lambda: worksheet.col_values(TX_TS)))
        shop_len = len(with_sheets_retry(lambda: worksheet.col_values(SHOP_TS)))
        income_len = len(with_sheets_retry(lambda: worksheet.col_values(INC_TS)))
        last = max(main_len, shop_len, income_len)
        raw_totals = self._cell_value(worksheet, rowcol_to_a1(1, META_TOTALS))
        parsed = _as_int(raw_totals)
        if parsed is not None:
            last = max(last, parsed - 2)
        if last < FIRST_DATA_ROW:
            return []

        def _load(a1: str) -> list[list[str]]:
            values = with_sheets_retry(
                lambda: worksheet.get(
                    a1,
                    value_render_option=ValueRenderOption.formatted,
                )
            )
            return values or []

        main = _load(f"A{FIRST_DATA_ROW}:D{last}")
        types = _load(
            f"{rowcol_to_a1(FIRST_DATA_ROW, TX_TYPE_COL)}:"
            f"{rowcol_to_a1(last, TX_TYPE_COL)}"
        )
        messages = _load(
            f"{rowcol_to_a1(FIRST_DATA_ROW, TX_MSG_COL)}:"
            f"{rowcol_to_a1(last, TX_MSG_COL)}"
        )
        shopping = _load(f"F{FIRST_DATA_ROW}:H{last}")
        shop_messages = _load(
            f"{rowcol_to_a1(FIRST_DATA_ROW, SHOP_MSG_COL)}:"
            f"{rowcol_to_a1(last, SHOP_MSG_COL)}"
        )
        saved: list[tuple[str, ParsedEntry, int]] = []
        count = last - FIRST_DATA_ROW + 1
        for index in range(count):
            sheet_row = FIRST_DATA_ROW + index
            main_row = main[index] if index < len(main) else []
            amount = _parse_amount(main_row[1] if len(main_row) > 1 else "")
            description = main_row[2].strip() if len(main_row) > 2 else ""
            if amount is not None and description:
                type_row = types[index] if index < len(types) else []
                entry_type = (
                    "income"
                    if type_row and str(type_row[0]).strip().lower() == "income"
                    else "expense"
                )
                note = main_row[3].strip() if len(main_row) > 3 else ""
                msg_row = messages[index] if index < len(messages) else []
                saved.append(
                    (
                        main_row[0] if main_row else "",
                        ParsedEntry(
                            entry_type,
                            amount,
                            description,
                            note or None,
                        ),
                        _message_id(msg_row[0] if msg_row else "", sheet_row),
                    )
                )
            shop_row = shopping[index] if index < len(shopping) else []
            shop_amount = _parse_amount(shop_row[1] if len(shop_row) > 1 else "")
            if shop_amount is not None:
                note = shop_row[2].strip() if len(shop_row) > 2 else ""
                msg_row = shop_messages[index] if index < len(shop_messages) else []
                saved.append(
                    (
                        shop_row[0] if shop_row else "",
                        ParsedEntry("expense", shop_amount, SHOPPING_WORD, note or None),
                        _message_id(
                            msg_row[0] if msg_row else "",
                            10_000_000_000_000 + sheet_row,
                        ),
                    )
                )
        return saved

    def _cell_value(self, worksheet: gspread.Worksheet, cell: str) -> str | None:
        try:
            return with_sheets_retry(lambda: worksheet.acell(cell).value)
        except APIError as error:
            if "exceeds grid limits" in str(error).lower():
                return None
            raise

    def _ensure_grid(self, worksheet: gspread.Worksheet) -> None:
        if worksheet.col_count >= RAW_MSG and worksheet.row_count >= 1000:
            return
        with_sheets_retry(
            lambda: worksheet.resize(
                rows=max(worksheet.row_count, 1000),
                cols=max(worksheet.col_count, RAW_MSG),
            )
        )

    def add_section(self, worksheet: gspread.Worksheet, name: str) -> int:
        layout = self._cell_value(worksheet, rowcol_to_a1(1, META_LAYOUT))
        if layout != LAYOUT_VERSION:
            raise ValueError("Ledger must be on the latest layout before adding sections")
        return ledger_v6.add_custom_section(self, worksheet, name)

    def _data_start_row(self, worksheet: gspread.Worksheet) -> int:
        if self._cell_value(worksheet, rowcol_to_a1(1, META_LAYOUT)) == LAYOUT_VERSION:
            return ledger_v6.FIRST_DATA_ROW
        return FIRST_DATA_ROW

    def _expense_write_columns(
        self, worksheet: gspread.Worksheet, category: str
    ) -> tuple[int, int, int]:
        stem = category_stem(category)
        recurring = False
        count = 0
        for col in (TX_CAT, ONEOFF_CAT, RAW_CAT):
            values = with_sheets_retry(lambda c=col: worksheet.col_values(c))
            for index, value in enumerate(values):
                if index == 0:
                    continue
                current = re.sub(r"\s+", " ", value).strip()
                if not current or category_stem(current) != stem:
                    continue
                count += 1
                if col == TX_CAT:
                    recurring = True
        if recurring or count >= 1:
            return TX_TS, TX_NOTE, TX_MSG_COL
        return ONEOFF_TS, ONEOFF_NOTE, ONEOFF_MSG_COL

    def _ensure_ledger_layout(self, worksheet: gspread.Worksheet, user_id: int) -> None:
        self._ensure_grid(worksheet)
        marker = self._cell_value(worksheet, rowcol_to_a1(1, META_LAYOUT))
        old_marker = self._cell_value(worksheet, rowcol_to_a1(1, OLD_LAYOUT_COL))
        if marker == LAYOUT_VERSION:
            return
        if marker == LAYOUT_V5:
            ledger_v6.migrate_to_v6(self, worksheet)
            return
        if old_marker == LAYOUT_V4 or marker == LAYOUT_V4:
            self._upgrade_v4(worksheet, inserted=old_marker != LAYOUT_V4)
            return
        existing = with_sheets_retry(lambda: worksheet.col_values(TX_TS))
        if len(existing) > 1:
            self._upgrade_legacy_to_v5(worksheet, user_id)
            return
        self._bootstrap_ledger(worksheet, user_id)

    def _looks_like_timestamp(self, value: Any) -> bool:
        return bool(TIMESTAMP_LIKE.match(str(value or "").strip()))

    def _row_looks_like_transaction(
        self, worksheet: gspread.Worksheet, row: int
    ) -> bool:
        stamp = self._cell_value(worksheet, rowcol_to_a1(row, TX_TS))
        amount = _parse_amount(self._cell_value(worksheet, rowcol_to_a1(row, TX_AMT)))
        category = self._cell_value(worksheet, rowcol_to_a1(row, TX_CAT))
        if amount is not None and (self._looks_like_timestamp(stamp) or category):
            return True
        return False

    def _last_entry_row_across_zones(
        self, worksheet: gspread.Worksheet, *, floor: int | None = None
    ) -> int:
        start = self._data_start_row(worksheet)
        if floor is not None and floor >= start:
            return floor
        last = start - 1
        layout = self._cell_value(worksheet, rowcol_to_a1(1, META_LAYOUT))
        if layout == LAYOUT_VERSION:
            for column in (TX_TS, ONEOFF_TS, SHOP_TS, INC_TS):
                values = with_sheets_retry(lambda col=column: worksheet.col_values(col))
                for index in range(len(values) - 1, 0, -1):
                    if ledger_v6.parse_sheet_datetime(values[index]):
                        last = max(last, index + 1)
                        break
            return max(last, start)
        columns = (RAW_TS, RAW_AMT, RAW_CAT, SHOP_TS, SHOP_AMT, INC_TS, INC_AMT)
        for column in columns:
            values = with_sheets_retry(lambda col=column: worksheet.col_values(col))
            for index in range(len(values) - 1, 0, -1):
                if str(values[index]).strip():
                    last = max(last, index + 1)
                    break
        return max(last, start)

    def _upgrade_legacy_to_v5(self, worksheet: gspread.Worksheet, user_id: int) -> None:
        marker = self._cell_value(worksheet, rowcol_to_a1(1, META_LAYOUT))
        if marker in (LAYOUT_VERSION, LAYOUT_V5):
            return

        self._ensure_grid(worksheet)
        self._write_values(
            worksheet,
            f"A{HEADER_ROW}:D{HEADER_ROW}",
            [TX_HEADERS],
        )
        self._write_values(
            worksheet,
            f"{rowcol_to_a1(HEADER_ROW, ONEOFF_TS)}:{rowcol_to_a1(HEADER_ROW, ONEOFF_NOTE)}",
            [ONEOFF_HEADERS],
        )
        self._write_values(
            worksheet,
            f"{rowcol_to_a1(HEADER_ROW, SHOP_TS)}:{rowcol_to_a1(HEADER_ROW, SHOP_NOTE)}",
            [SHOP_HEADERS],
        )
        self._write_values(
            worksheet,
            f"{rowcol_to_a1(HEADER_ROW, INC_TS)}:{rowcol_to_a1(HEADER_ROW, INC_NOTE)}",
            [INC_HEADERS],
        )

        last_row = self._last_entry_row_across_zones(worksheet)
        block = with_sheets_retry(
            lambda: worksheet.get(
                f"A{FIRST_DATA_ROW}:{rowcol_to_a1(last_row, INC_NOTE)}",
                value_render_option=ValueRenderOption.formatted,
            )
        )

        def _at(row_values: list[str], column: int) -> str:
            index = column - 1
            if index >= len(row_values):
                return ""
            return str(row_values[index] or "")

        for offset, row_values in enumerate(block or []):
            row = FIRST_DATA_ROW + offset
            stamp = _at(row_values, TX_TS)
            amount = _parse_amount(_at(row_values, TX_AMT))
            category = _at(row_values, TX_CAT)
            note = _at(row_values, TX_NOTE)
            if amount is not None or category.strip() or self._looks_like_timestamp(stamp):
                self._write_values(
                    worksheet,
                    f"{rowcol_to_a1(row, RAW_TS)}:{rowcol_to_a1(row, RAW_NOTE)}",
                    [[stamp, amount if amount is not None else "", category, note]],
                    user_entered=True,
                )

            shop_stamp = _at(row_values, SHOP_TS)
            shop_amount = _parse_amount(_at(row_values, SHOP_AMT))
            shop_note = _at(row_values, SHOP_NOTE)
            if shop_amount is None:
                shop_stamp = _at(row_values, ONEOFF_TS)
                shop_amount = _parse_amount(_at(row_values, ONEOFF_AMT))
                shop_note = _at(row_values, ONEOFF_NOTE) or _at(row_values, ONEOFF_CAT)
            if shop_amount is not None:
                self._write_values(
                    worksheet,
                    f"{rowcol_to_a1(row, SHOP_TS)}:{rowcol_to_a1(row, SHOP_NOTE)}",
                    [[shop_stamp, shop_amount, shop_note]],
                    user_entered=True,
                )

            income_amount = _parse_amount(_at(row_values, INC_AMT))
            if income_amount is not None:
                self._write_values(
                    worksheet,
                    f"{rowcol_to_a1(row, INC_TS)}:{rowcol_to_a1(row, INC_NOTE)}",
                    [
                        [
                            _at(row_values, INC_TS),
                            income_amount,
                            _at(row_values, INC_CAT),
                            _at(row_values, INC_NOTE),
                        ]
                    ],
                    user_entered=True,
                )

        totals_row = last_row + 2
        self._set_totals_row(worksheet, totals_row)
        owner = (
            self._cell_value(worksheet, rowcol_to_a1(1, META_OWNER))
            or self._cell_value(worksheet, rowcol_to_a1(1, OLD_OWNER_COL))
            or self._cell_value(worksheet, "Z1")
        )
        if not owner:
            self._write_values(
                worksheet,
                rowcol_to_a1(1, META_OWNER),
                [[_owner_token(user_id)]],
            )
        self._write_values(
            worksheet,
            rowcol_to_a1(1, META_LAYOUT),
            [[LAYOUT_V5]],
        )
        self._write_auto_summary(worksheet)
        ledger_v6.migrate_to_v6(self, worksheet)

    def _upgrade_v4(self, worksheet: gspread.Worksheet, *, inserted: bool) -> None:
        moved: list[tuple[int, str, float, str, str, str]] = []
        if not inserted:
            categories = with_sheets_retry(lambda: worksheet.col_values(TX_CAT))
            amounts = with_sheets_retry(lambda: worksheet.col_values(TX_AMT))
            notes = with_sheets_retry(lambda: worksheet.col_values(TX_NOTE))
            stamps = with_sheets_retry(lambda: worksheet.col_values(TX_TS))
            messages = with_sheets_retry(lambda: worksheet.col_values(TX_MSG_COL - 5))
            for index, category in enumerate(categories):
                if index == 0 or category.strip().lower() not in MOVED_ONEOFFS:
                    continue
                amount = _parse_amount(amounts[index] if index < len(amounts) else "")
                if amount is None:
                    continue
                moved.append(
                    (
                        index + 1,
                        stamps[index] if index < len(stamps) else "",
                        amount,
                        category.strip(),
                        notes[index].strip() if index < len(notes) else "",
                        messages[index] if index < len(messages) else "",
                    )
                )
            self._insert_columns(worksheet, GAP_COL + 1, 5)

        self._write_values(
            worksheet,
            f"{rowcol_to_a1(HEADER_ROW, ONEOFF_TS)}:{rowcol_to_a1(HEADER_ROW, ONEOFF_NOTE)}",
            [ONEOFF_HEADERS],
        )
        for row, stamp, amount, category, note, message_id in moved:
            self._write_values(
                worksheet,
                f"{rowcol_to_a1(row, ONEOFF_TS)}:{rowcol_to_a1(row, ONEOFF_NOTE)}",
                [[stamp, amount, category, note]],
                user_entered=True,
            )
            self._write_values(
                worksheet,
                f"{rowcol_to_a1(row, TX_TS)}:{rowcol_to_a1(row, TX_NOTE)}",
                [["", "", "", ""]],
            )
            if message_id:
                self._write_values(worksheet, rowcol_to_a1(row, ONEOFF_MSG_COL), [[message_id]])
                self._write_values(worksheet, rowcol_to_a1(row, TX_MSG_COL), [[""]])
        self._write_values(
            worksheet,
            rowcol_to_a1(1, META_LAYOUT),
            [[LAYOUT_V5]],
        )
        self._write_auto_summary(worksheet)
        ledger_v6.migrate_to_v6(self, worksheet)

    def _insert_columns(self, worksheet: gspread.Worksheet, column: int, count: int) -> None:
        request = {
            "insertDimension": {
                "range": {
                    "sheetId": worksheet.id,
                    "dimension": "COLUMNS",
                    "startIndex": column - 1,
                    "endIndex": column - 1 + count,
                },
                "inheritFromBefore": False,
            }
        }
        with_sheets_retry(
            lambda: worksheet.spreadsheet.batch_update({"requests": [request]})
        )

    def _raw_span(self, end_row: int) -> tuple[str, str, str, str]:
        stamp = _col_letters(RAW_TS)
        amount = _col_letters(RAW_AMT)
        category = _col_letters(RAW_CAT)
        note = _col_letters(RAW_NOTE)
        return (
            f"{stamp}2:{stamp}{end_row}",
            f"{amount}2:{amount}{end_row}",
            f"{category}2:{category}{end_row}",
            f"{note}2:{note}{end_row}",
        )

    def _transactions_formula(self, end_row: int) -> str:
        stamp, _amount, category, note = self._raw_span(end_row)
        block = f"{stamp.split(':')[0]}:{note.split(':')[1]}"
        return (
            f'=IFERROR(FILTER({block},'
            f'({category}<>"")*(COUNTIF({category},{category})>1)),"")'
        )

    def _oneoff_formula(self, end_row: int) -> str:
        stamp, _amount, category, note = self._raw_span(end_row)
        block = f"{stamp.split(':')[0]}:{note.split(':')[1]}"
        return (
            f'=IFERROR(SORT(FILTER({block},'
            f'({category}<>"")*(COUNTIF({category},{category})=1)),2,FALSE),"")'
        )

    def _canonical_expense_category(self, worksheet: gspread.Worksheet, name: str) -> str:
        text = re.sub(r"\s+", " ", name).strip()
        stem = category_stem(text)
        layout = self._cell_value(worksheet, rowcol_to_a1(1, META_LAYOUT))
        cat_columns = (RAW_CAT,)
        if layout == LAYOUT_VERSION:
            cat_columns = (TX_CAT, ONEOFF_CAT, RAW_CAT)
        labels = [text]
        rows: dict[tuple[int, int], str] = {}
        for col in cat_columns:
            values = with_sheets_retry(lambda c=col: worksheet.col_values(c))
            for index, value in enumerate(values):
                if index == 0:
                    continue
                current = re.sub(r"\s+", " ", value).strip()
                if not current or category_stem(current) != stem:
                    continue
                labels.append(current)
                rows[(index + 1, col)] = current
        chosen = preferred_category(labels)
        for (row, col), current in rows.items():
            if current != chosen:
                self._write_values(worksheet, rowcol_to_a1(row, col), [[chosen]])
        return chosen

    def _normalize_raw_categories(self, worksheet: gspread.Worksheet, end_row: int) -> None:
        values = with_sheets_retry(lambda: worksheet.col_values(RAW_CAT))
        stems: dict[str, list[tuple[int, str]]] = {}
        for index, value in enumerate(values):
            if index == 0:
                continue
            row = index + 1
            if row > end_row:
                break
            current = re.sub(r"\s+", " ", value).strip()
            if not current:
                continue
            stem = category_stem(current)
            stems.setdefault(stem, []).append((row, current))
        for entries in stems.values():
            labels = [label for _row, label in entries]
            chosen = preferred_category(labels)
            for row, current in entries:
                if current != chosen:
                    self._write_values(
                        worksheet,
                        rowcol_to_a1(row, RAW_CAT),
                        [[chosen]],
                    )

    def _aggregate_category_totals(
        self, worksheet: gspread.Worksheet, end_row: int
    ) -> list[tuple[str, float]]:
        amounts: dict[str, float] = {}
        counts: dict[str, int] = {}
        labels: dict[str, list[str]] = {}
        layout = self._cell_value(worksheet, rowcol_to_a1(1, META_LAYOUT))
        zone_specs = [(RAW_CAT, RAW_AMT)]
        if layout == LAYOUT_VERSION:
            zone_specs = [
                (TX_CAT, TX_AMT),
                (ONEOFF_CAT, ONEOFF_AMT),
            ]
        for cat_col, amt_col in zone_specs:
            categories = with_sheets_retry(lambda c=cat_col: worksheet.col_values(c))
            amounts_col = with_sheets_retry(lambda a=amt_col: worksheet.col_values(a))
            for index in range(1, min(len(categories), end_row + 1)):
                category = re.sub(r"\s+", " ", categories[index]).strip()
                amount = _parse_amount(
                    amounts_col[index] if index < len(amounts_col) else ""
                )
                if not category or amount is None:
                    continue
                stem = category_stem(category)
                amounts[stem] = amounts.get(stem, 0.0) + amount
                counts[stem] = counts.get(stem, 0) + 1
                labels.setdefault(stem, []).append(category)

        pairs: list[tuple[str, float]] = []
        for stem, total in amounts.items():
            if counts.get(stem, 0) <= 1:
                continue
            pairs.append((preferred_category(labels[stem]), total))

        shop_values = with_sheets_retry(lambda: worksheet.col_values(SHOP_AMT))
        shop_total = 0.0
        for index in range(1, min(len(shop_values), end_row + 1)):
            amount = _parse_amount(shop_values[index])
            if amount is not None:
                shop_total += amount
        if shop_total:
            pairs.append((TOTAL_SHOPPING, shop_total))

        pairs.sort(key=lambda item: item[1], reverse=True)
        return pairs

    def _clear_legacy_category_formulas(
        self, worksheet: gspread.Worksheet, totals_row: int
    ) -> None:
        for column in (TX_TS, TOTALS_START):
            if self._cell_formula(worksheet, totals_row, column):
                self._write_values(
                    worksheet,
                    rowcol_to_a1(totals_row, column),
                    [[""]],
                )

    def _write_category_totals(
        self,
        worksheet: gspread.Worksheet,
        totals_row: int,
        end_row: int,
    ) -> None:
        blank = [""] * INC_NOTE
        self._write_values(
            worksheet,
            f"A{totals_row}:{rowcol_to_a1(totals_row + 1, INC_NOTE)}",
            [blank, blank],
        )
        pairs = self._aggregate_category_totals(worksheet, end_row)
        layout = self._cell_value(worksheet, rowcol_to_a1(1, META_LAYOUT))
        if layout == LAYOUT_VERSION:
            start = max(
                ledger_v6.category_block_row(totals_row),
                end_row + 5,
            )
            clear_end = start + max(len(pairs) + 1, 12)
            self._write_values(
                worksheet,
                f"A{start}:{rowcol_to_a1(clear_end, 2)}",
                [["", ""] for _ in range(clear_end - start + 1)],
            )
            if not pairs:
                return
            rows = [["Category", "Total"], *[[label, amount] for label, amount in pairs]]
            self._write_values(
                worksheet,
                f"A{start}:{rowcol_to_a1(start + len(rows) - 1, 2)}",
                rows,
                user_entered=True,
            )
            return

        slots = _category_display_columns()
        label_row = [""] * INC_NOTE
        value_row = [""] * INC_NOTE
        for index, (label, amount) in enumerate(pairs):
            if index >= len(slots):
                break
            column = slots[index]
            label_row[column - 1] = label
            value_row[column - 1] = amount
        if any(label_row) or any(value_row):
            self._write_values(
                worksheet,
                f"A{totals_row}:{rowcol_to_a1(totals_row, INC_NOTE)}",
                [label_row],
            )
            self._write_values(
                worksheet,
                f"A{totals_row + 1}:{rowcol_to_a1(totals_row + 1, INC_NOTE)}",
                [value_row],
                user_entered=True,
            )

    def _write_auto_summary(
        self, worksheet: gspread.Worksheet, *, end_row: int | None = None
    ) -> None:
        totals_row = self._sync_totals_row(worksheet)
        expense_total_row = totals_row + SUMMARY_FOOTER_OFFSET
        income_row = expense_total_row + 1
        net_row = income_row + 1
        width = 60
        if end_row is None:
            end_row = self._last_entry_row_across_zones(worksheet)
        self._normalize_raw_categories(worksheet, end_row)
        self._write_values(
            worksheet,
            f"{rowcol_to_a1(HEADER_ROW, RAW_TS)}:{rowcol_to_a1(HEADER_ROW, RAW_NOTE)}",
            [TX_HEADERS],
        )
        scrub_start = max(FIRST_DATA_ROW, end_row + 1)
        scrub_end = max(net_row, totals_row + SUMMARY_FOOTER_OFFSET + 3)
        scrub_end = min(
            scrub_end + LEDGER_SUMMARY_SCRUB_TAIL,
            LEDGER_FORMAT_LAST_ROW,
        )
        with_sheets_retry(
            lambda: worksheet.batch_clear(
                [
                    f"A{FIRST_DATA_ROW}:D{end_row}",
                    f"F{FIRST_DATA_ROW}:I{end_row}",
                    f"A{scrub_start}:{rowcol_to_a1(scrub_end, width)}",
                ]
            )
        )
        income_amt = _col_letters(INC_AMT)
        raw_amt = _col_letters(RAW_AMT)
        shop_amt = _col_letters(SHOP_AMT)
        income_cell = rowcol_to_a1(income_row, 2)
        expense_cell = rowcol_to_a1(expense_total_row, 2)
        updates = [
            {
                "range": rowcol_to_a1(FIRST_DATA_ROW, TX_TS),
                "values": [[self._transactions_formula(end_row)]],
            },
            {
                "range": rowcol_to_a1(FIRST_DATA_ROW, ONEOFF_TS),
                "values": [[self._oneoff_formula(end_row)]],
            },
            {"range": rowcol_to_a1(expense_total_row, 1), "values": [["Total expense"]]},
            {
                "range": expense_cell,
                "values": [[
                    f"=SUM({raw_amt}2:{raw_amt}{end_row})+SUM({shop_amt}2:{shop_amt}{end_row})"
                ]],
            },
            {"range": rowcol_to_a1(income_row, 1), "values": [["Total income"]]},
            {
                "range": income_cell,
                "values": [[f"=SUM({income_amt}2:{income_amt}{end_row})"]],
            },
            {"range": rowcol_to_a1(net_row, 1), "values": [["Net"]]},
            {
                "range": rowcol_to_a1(net_row, 2),
                "values": [[
                    f"={income_cell}-{expense_cell}"
                ]],
            },
        ]
        with_sheets_retry(
            lambda: worksheet.batch_update(
                updates,
                value_input_option=ValueInputOption.user_entered,
            )
        )
        value_row = totals_row + 1
        value_width = self._category_totals_width(worksheet, totals_row)
        value_start = self._category_totals_start_col(worksheet, totals_row)
        with_sheets_retry(
            lambda: worksheet.format(
                f"{rowcol_to_a1(value_row, value_start)}:"
                f"{rowcol_to_a1(value_row, value_width)}",
                {"numberFormat": {"type": "CURRENCY", "pattern": "$#,##0.00"}},
            )
        )
        with_sheets_retry(
            lambda: worksheet.format(
                f"B{expense_total_row}:B{net_row}",
                {"numberFormat": {"type": "CURRENCY", "pattern": "$#,##0.00"}},
            )
        )

    def _bootstrap_ledger(self, worksheet: gspread.Worksheet, user_id: int) -> None:
        self._ensure_grid(worksheet)
        with_sheets_retry(
            lambda: worksheet.insert_row(
                [""] * worksheet.col_count,
                index=1,
                inherit_from_before=False,
            )
        )
        ledger_v6.write_headers(self, worksheet)
        self._write_values(
            worksheet,
            rowcol_to_a1(1, META_OWNER),
            [[_owner_token(user_id)]],
        )
        totals_row = ledger_v6.FIRST_DATA_ROW + 1
        self._write_values(
            worksheet,
            rowcol_to_a1(1, META_TOTALS),
            [[totals_row]],
        )
        self._write_values(
            worksheet,
            rowcol_to_a1(1, META_LAYOUT),
            [[LAYOUT_VERSION]],
        )
        ledger_v6.refresh_footer(self, worksheet, ledger_v6.FIRST_DATA_ROW)
        self._format_ledger_v6(worksheet)

    def _unique_ledger_title(self, display_name: str, user_id: int) -> str:
        titles = {ws.title for ws in self._spreadsheet.worksheets()}
        candidate = tab_name(display_name, user_id)
        if candidate not in titles:
            return candidate
        fallback = f"{LEDGER_PREFIX}{user_id}"
        if len(fallback) > 100:
            fallback = fallback[:100]
        if fallback not in titles:
            return fallback
        if candidate in titles:
            return candidate
        return fallback

    def _ensure_categories(self) -> gspread.Worksheet:
        try:
            worksheet = self._spreadsheet.worksheet(CATEGORIES_SHEET)
        except WorksheetNotFound:
            return self._create_categories()

        headers = with_sheets_retry(lambda: worksheet.row_values(1))
        if headers[:2] == CATEGORY_HEADERS:
            return worksheet

        legacy_title = LEGACY_CATEGORIES_SHEET
        existing = {ws.title for ws in self._spreadsheet.worksheets()}
        if legacy_title in existing:
            legacy_title = f"{LEGACY_CATEGORIES_SHEET} {worksheet.id}"
        with_sheets_retry(lambda: worksheet.update_title(legacy_title))
        return self._create_categories()

    def _create_categories(self) -> gspread.Worksheet:
        worksheet = self._spreadsheet.add_worksheet(
            title=CATEGORIES_SHEET, rows=200, cols=4
        )
        self._write_values(
            worksheet,
            "A1:B5",
            [CATEGORY_HEADERS, *CATEGORY_SEEDS],
        )
        self._format_header_row(worksheet, 2)
        return worksheet

    def _ensure_summary_sheet(self) -> gspread.Worksheet:
        try:
            return self._spreadsheet.worksheet(SUMMARY_SHEET)
        except WorksheetNotFound:
            worksheet = self._spreadsheet.add_worksheet(
                title=SUMMARY_SHEET, rows=200, cols=26
            )
            return worksheet

    def _ensure_summary_block(self, ledger_title: str, display_name: str) -> None:
        if ledger_title in self._summary_linked:
            return
        summary = self._ensure_summary_sheet()
        values = with_sheets_retry(lambda: summary.get_all_values())
        ref = _sheet_ref(ledger_title)
        for row_index, row in enumerate(values, start=1):
            for col_index, cell in enumerate(row, start=1):
                if cell == ledger_title:
                    self._refresh_summary_formulas(summary, row_index, col_index, ref)
                    self._summary_linked.add(ledger_title)
                    return

        start_col = 1
        if values:
            header = values[0]
            while start_col - 1 < len(header) and header[start_col - 1].strip():
                start_col += SUMMARY_BLOCK_WIDTH

        ref = _sheet_ref(ledger_title)
        income, expenses = _summary_math(ref)
        name_cell = rowcol_to_a1(1, start_col)
        income_label = rowcol_to_a1(2, start_col)
        income_value = rowcol_to_a1(2, start_col + 1)
        expense_label = rowcol_to_a1(3, start_col)
        expense_value = rowcol_to_a1(3, start_col + 1)
        net_label = rowcol_to_a1(4, start_col)
        net_value = rowcol_to_a1(4, start_col + 1)
        tab_label = rowcol_to_a1(5, start_col)
        tab_value = rowcol_to_a1(5, start_col + 1)
        category_header = rowcol_to_a1(6, start_col)
        total_header = rowcol_to_a1(6, start_col + 1)
        query_cell = rowcol_to_a1(SUMMARY_QUERY_ROW, start_col)
        net_formula = f"={income_value}-{expense_value}"
        query = (
            f'=IFERROR(QUERY({ref}!B:C,'
            f'"select C, sum(B) where C <> \'\' group by C label sum(B) \'\'",0),"")'
        )
        label = _sanitize_tab_part(display_name) or ledger_title

        updates = [
            {"range": name_cell, "values": [[label]]},
            {"range": income_label, "values": [["Income"]]},
            {"range": income_value, "values": [[income]]},
            {"range": expense_label, "values": [["Expenses"]]},
            {"range": expense_value, "values": [[expenses]]},
            {"range": net_label, "values": [["Net"]]},
            {"range": net_value, "values": [[net_formula]]},
            {"range": tab_label, "values": [["Tab"]]},
            {"range": tab_value, "values": [[ledger_title]]},
            {"range": category_header, "values": [["Category"]]},
            {"range": total_header, "values": [["Total"]]},
            {"range": query_cell, "values": [[query]]},
        ]
        with_sheets_retry(
            lambda: summary.batch_update(
                updates,
                value_input_option=ValueInputOption.user_entered,
            )
        )
        self._format_summary_block(summary, start_col)
        self._summary_linked.add(ledger_title)

    def _refresh_summary_formulas(
        self,
        summary: gspread.Worksheet,
        title_row: int,
        title_col: int,
        ref: str,
    ) -> None:
        if title_row != 5:
            return
        value_col = title_col
        income, expenses = _summary_math(ref)
        income_cell = rowcol_to_a1(2, value_col)
        expense_cell = rowcol_to_a1(3, value_col)
        query_cell = rowcol_to_a1(SUMMARY_QUERY_ROW, value_col - 1)
        query = (
            f'=IFERROR(QUERY({ref}!B:C,'
            f'"select C, sum(B) where C <> \'\' group by C label sum(B) \'\'",0),"")'
        )
        updates = [
            {"range": income_cell, "values": [[income]]},
            {"range": expense_cell, "values": [[expenses]]},
            {
                "range": rowcol_to_a1(4, value_col),
                "values": [[f"={income_cell}-{expense_cell}"]],
            },
            {"range": query_cell, "values": [[query]]},
        ]
        with_sheets_retry(
            lambda: summary.batch_update(
                updates,
                value_input_option=ValueInputOption.user_entered,
            )
        )

    def _sync_totals_row(self, worksheet: gspread.Worksheet) -> int:
        last_entry = self._last_entry_row_across_zones(worksheet)
        totals_row = max(last_entry + 2, INITIAL_TOTALS_ROW)
        layout = self._cell_value(worksheet, rowcol_to_a1(1, META_LAYOUT))
        if layout in (LAYOUT_VERSION, LAYOUT_V5):
            self._set_totals_row(worksheet, totals_row)
        return totals_row

    def _totals_row(self, worksheet: gspread.Worksheet) -> int:
        layout = self._cell_value(worksheet, rowcol_to_a1(1, META_LAYOUT))
        if layout in (LAYOUT_VERSION, LAYOUT_V5):
            return self._sync_totals_row(worksheet)
        raw = with_sheets_retry(
            lambda: worksheet.acell(rowcol_to_a1(1, META_TOTALS)).value
        )
        parsed = _as_int(raw)
        if parsed is None:
            return INITIAL_TOTALS_ROW
        last_entry = self._last_entry_row_across_zones(worksheet)
        if self._row_looks_like_transaction(worksheet, parsed):
            return last_entry + 2
        if parsed <= last_entry:
            return last_entry + 2
        return parsed

    def _set_totals_row(self, worksheet: gspread.Worksheet, row: int) -> None:
        self._write_values(worksheet, rowcol_to_a1(1, META_TOTALS), [[row]])

    def _next_data_row(self, worksheet: gspread.Worksheet, timestamp_col: int) -> int:
        layout = self._cell_value(worksheet, rowcol_to_a1(1, META_LAYOUT))
        band = _section_band_for_ts(timestamp_col)
        if layout == LAYOUT_VERSION and band is not None:
            return ledger_v6.next_open_row(self, worksheet, band[0], band[1])
        totals_row = self._totals_row(worksheet)
        start = self._data_start_row(worksheet)
        values = with_sheets_retry(lambda: worksheet.col_values(timestamp_col))
        last_data_row = totals_row - 2
        for row in range(start, last_data_row + 1):
            existing = values[row - 1] if row - 1 < len(values) else ""
            if not str(existing).strip():
                return row

        spacer = totals_row - 1
        self._insert_row(worksheet, spacer)
        self._set_totals_row(worksheet, totals_row + 1)
        return spacer

    def _last_data_row(self, worksheet: gspread.Worksheet) -> int:
        return self._last_entry_row_across_zones(worksheet)

    def _summary_footer_start(
        self, worksheet: gspread.Worksheet, totals_row: int
    ) -> int | None:
        for offset in range(2, 10):
            cell = self._cell_value(worksheet, rowcol_to_a1(totals_row + offset, 1))
            if cell and str(cell).strip().casefold() == "total expense":
                return totals_row + offset
        return None

    def _category_totals_start_col(
        self, worksheet: gspread.Worksheet, totals_row: int
    ) -> int:
        slots = _category_display_columns()
        values = with_sheets_retry(lambda: worksheet.row_values(totals_row))
        for column in slots:
            text = values[column - 1].strip() if column - 1 < len(values) else ""
            if text and text not in TAIL_LABELS:
                return column
        return TX_TS

    def _category_totals_width(self, worksheet: gspread.Worksheet, totals_row: int) -> int:
        slots = _category_display_columns()
        values = with_sheets_retry(lambda: worksheet.row_values(totals_row))
        width = TX_TS
        for column in slots:
            text = values[column - 1].strip() if column - 1 < len(values) else ""
            if text and text not in TAIL_LABELS:
                width = max(width, column)
        return width

    def _neutral_summary_band_requests(
        self,
        sheet_id: int,
        first_row: int,
        last_row: int,
    ) -> list[dict[str, Any]]:
        if first_row > last_row:
            return []
        neutral = {"backgroundColor": _sheet_rgb("#FFFFFF")}
        requests: list[dict[str, Any]] = []
        for gap in (GAP_COL, GAP_COL_2, GAP_COL_3, GAP_COL_4):
            requests.append(
                _ledger_repeat_cell(
                    sheet_id,
                    first_row,
                    last_row,
                    gap,
                    gap,
                    neutral,
                    "userEnteredFormat.backgroundColor",
                )
            )
        for column in range(GAP_COL_4 + 1, LEDGER_SUMMARY_FORMAT_END + 1):
            requests.append(
                _ledger_repeat_cell(
                    sheet_id,
                    first_row,
                    last_row,
                    column,
                    column,
                    neutral,
                    "userEnteredFormat.backgroundColor",
                )
            )
        return requests

    def _cell_formula(
        self, worksheet: gspread.Worksheet, row: int, column: int
    ) -> str | None:
        a1 = rowcol_to_a1(row, column)
        values = with_sheets_retry(
            lambda: worksheet.get(
                a1,
                value_render_option=ValueRenderOption.formula,
            )
        )
        if not values or not values[0]:
            return None
        text = str(values[0][0] or "").strip()
        return text or None

    def _reanchor_category_summary(self, worksheet: gspread.Worksheet) -> None:
        totals_row = self._totals_row(worksheet)
        self._clear_legacy_category_formulas(worksheet, totals_row)

    def _ensure_summary_gap(self, worksheet: gspread.Worksheet) -> None:
        totals_row = self._totals_row(worksheet)
        probe = self._cell_value(worksheet, rowcol_to_a1(totals_row + 2, 1))
        if probe and str(probe).strip().casefold() == "total expense":
            self._insert_dimension_row(worksheet, totals_row + 2)

    def _insert_dimension_row(self, worksheet: gspread.Worksheet, row: int) -> None:
        request = {
            "insertDimension": {
                "range": {
                    "sheetId": worksheet.id,
                    "dimension": "ROWS",
                    "startIndex": row - 1,
                    "endIndex": row,
                },
                "inheritFromBefore": False,
            }
        }
        with_sheets_retry(
            lambda: worksheet.spreadsheet.batch_update({"requests": [request]})
        )

    def _insert_row(self, worksheet: gspread.Worksheet, row: int) -> None:
        self._insert_dimension_row(worksheet, row)
        self._format_ledger(worksheet)

    def _category_labels(self, worksheet: gspread.Worksheet) -> list[str]:
        totals_row = self._totals_row(worksheet)
        values = with_sheets_retry(lambda: worksheet.row_values(totals_row))
        labels: list[str] = []
        for value in values:
            text = value.strip()
            if not text or text in TAIL_LABELS:
                continue
            labels.append(text)
        return labels

    def _ensure_category_total(self, worksheet: gspread.Worksheet, category: str) -> None:
        labels = self._category_labels(worksheet)
        if any(label.lower() == category.lower() for label in labels):
            return
        self._rewrite_totals(worksheet, [*labels, category])

    def _total_formula(self, label: str, header_cell: str, income_cell: str) -> str:
        expense_amt = f"${_col_letters(TX_AMT)}:${_col_letters(TX_AMT)}"
        expense_cat = f"${_col_letters(TX_CAT)}:${_col_letters(TX_CAT)}"
        income_amt = f"${_col_letters(INC_AMT)}:${_col_letters(INC_AMT)}"
        income_cat = f"${_col_letters(INC_CAT)}:${_col_letters(INC_CAT)}"
        shopping = f"${_col_letters(SHOP_AMT)}:${_col_letters(SHOP_AMT)}"
        if label == TOTAL_SHOPPING:
            return f"=SUM({shopping})"
        if label == TOTAL_INCOME:
            return f"=SUM({income_amt})"
        if label == TOTAL_NET:
            return f"={income_cell}-(SUM({expense_amt})+SUM({shopping}))"
        return (
            f"=SUMIF({expense_cat},{header_cell},{expense_amt})"
            f"+SUMIF({income_cat},{header_cell},{income_amt})"
        )

    def _rewrite_totals(self, worksheet: gspread.Worksheet, categories: list[str]) -> None:
        totals_row = self._totals_row(worksheet)
        labels = [*categories, *TAIL_LABELS]
        slots = [total_slot(index) for index in range(len(labels))]
        income_col = slots[labels.index(TOTAL_INCOME)]
        income_cell = rowcol_to_a1(totals_row + 1, income_col)
        width = max(slots) if slots else 1
        blank = [""] * width
        self._write_values(
            worksheet,
            f"A{totals_row}:{rowcol_to_a1(totals_row + 1, width)}",
            [blank, blank],
        )
        updates: list[dict[str, Any]] = []
        for label, column in zip(labels, slots, strict=True):
            header_cell = rowcol_to_a1(totals_row, column)
            updates.append({"range": header_cell, "values": [[label]]})
            updates.append(
                {
                    "range": rowcol_to_a1(totals_row + 1, column),
                    "values": [[self._total_formula(label, header_cell, income_cell)]],
                }
            )
        with_sheets_retry(
            lambda: worksheet.batch_update(
                updates,
                value_input_option=ValueInputOption.user_entered,
            )
        )
        self._format_ledger(worksheet)

    def _format_ledger_v6(self, worksheet: gspread.Worksheet) -> None:
        sheet_id = worksheet.id
        totals_row = self._totals_row(worksheet)
        data_end_row = max(self._last_data_row(worksheet), ledger_v6.FIRST_DATA_ROW)
        zone_end = max(data_end_row, totals_row - 1)
        expense_row = ledger_v6.footer_expense_row(totals_row)
        net_row = expense_row + 2
        category_row = ledger_v6.category_block_row(totals_row)
        summary_last = category_row + 16
        title_format = {
            "backgroundColor": _sheet_rgb(LEDGER_HEADER_BG),
            "textFormat": {"bold": True, "foregroundColor": _sheet_rgb("#FFFFFF")},
        }
        header_format = title_format
        currency = {
            "numberFormat": {"type": "CURRENCY", "pattern": "$#,##0.00"}
        }
        stamp_fmt = {
            "numberFormat": {"type": "DATE_TIME", "pattern": "yyyy-mm-dd hh:mm:ss"}
        }
        requests: list[dict[str, Any]] = [
            {
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": sheet_id,
                        "gridProperties": {"frozenRowCount": 2},
                    },
                    "fields": "gridProperties.frozenRowCount",
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": META_OWNER - 1,
                        "endIndex": RAW_MSG,
                    },
                    "properties": {"hiddenByUser": True},
                    "fields": "hiddenByUser",
                }
            },
        ]
        for gap in ledger_v6.GAP_COLUMNS:
            requests.append(
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": gap - 1,
                            "endIndex": gap,
                        },
                        "properties": {"hiddenByUser": True},
                        "fields": "hiddenByUser",
                    }
                }
            )
        for _name, start, end, color in ledger_v6.SECTION_BANDS:
            requests.append(
                _ledger_repeat_cell(
                    sheet_id,
                    ledger_v6.SECTION_TITLE_ROW,
                    ledger_v6.SECTION_TITLE_ROW,
                    start,
                    end,
                    title_format,
                    "userEnteredFormat(backgroundColor,textFormat)",
                )
            )
            requests.append(
                _ledger_repeat_cell(
                    sheet_id,
                    ledger_v6.HEADER_ROW,
                    ledger_v6.HEADER_ROW,
                    start,
                    end,
                    header_format,
                    "userEnteredFormat(backgroundColor,textFormat)",
                )
            )
            if zone_end >= ledger_v6.FIRST_DATA_ROW:
                requests.append(
                    _ledger_repeat_cell(
                        sheet_id,
                        ledger_v6.FIRST_DATA_ROW,
                        zone_end,
                        start,
                        end,
                        {"backgroundColor": _sheet_rgb(color)},
                        "userEnteredFormat.backgroundColor",
                    )
                )
        for _name, start, end, _color in ledger_v6.SECTION_BANDS:
            for column, size in (
                (start, 160),
                (start + 1, 110),
                (start + 2, 140 if end - start >= 3 else 240),
            ):
                if column > end:
                    break
                requests.append(
                    {
                        "updateDimensionProperties": {
                            "range": {
                                "sheetId": sheet_id,
                                "dimension": "COLUMNS",
                                "startIndex": column - 1,
                                "endIndex": column,
                            },
                            "properties": {"pixelSize": size},
                            "fields": "pixelSize",
                        }
                    }
                )
            note_col = end
            requests.append(
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": note_col - 1,
                            "endIndex": note_col,
                        },
                        "properties": {"pixelSize": 240},
                        "fields": "pixelSize",
                    }
                }
            )
        if zone_end >= ledger_v6.FIRST_DATA_ROW:
            for _name, start, end, _color in ledger_v6.SECTION_BANDS:
                for column in range(start, end + 1):
                    if column == start:
                        fmt = stamp_fmt
                    elif column == start + 1:
                        fmt = currency
                    else:
                        continue
                    requests.append(
                        _ledger_repeat_cell(
                            sheet_id,
                            ledger_v6.FIRST_DATA_ROW,
                            zone_end,
                            column,
                            column,
                            fmt,
                            "userEnteredFormat.numberFormat",
                        )
                    )
                requests.append(
                    _ledger_repeat_cell(
                        sheet_id,
                        ledger_v6.FIRST_DATA_ROW,
                        zone_end,
                        end,
                        end,
                        {"wrapStrategy": "WRAP"},
                        "userEnteredFormat.wrapStrategy",
                    )
                )
        if totals_row <= summary_last:
            white = {"backgroundColor": _sheet_rgb("#FFFFFF")}
            for _name, start, end, _color in ledger_v6.SECTION_BANDS:
                requests.append(
                    _ledger_repeat_cell(
                        sheet_id,
                        totals_row,
                        summary_last,
                        start,
                        end,
                        white,
                        "userEnteredFormat.backgroundColor",
                    )
                )
            summary_fill = {
                "backgroundColor": _sheet_rgb(LEDGER_SUMMARY_BG),
                "textFormat": {"bold": True},
            }
            requests.append(
                _ledger_repeat_cell(
                    sheet_id,
                    expense_row,
                    net_row,
                    1,
                    2,
                    summary_fill,
                    "userEnteredFormat(backgroundColor,textFormat)",
                )
            )
            requests.append(
                _ledger_repeat_cell(
                    sheet_id,
                    expense_row,
                    net_row,
                    2,
                    2,
                    {
                        "backgroundColor": _sheet_rgb(LEDGER_SUMMARY_BG),
                        "numberFormat": {
                            "type": "CURRENCY",
                            "pattern": "$#,##0.00",
                        },
                    },
                    "userEnteredFormat(backgroundColor,numberFormat)",
                )
            )
            requests.append(
                _ledger_repeat_cell(
                    sheet_id,
                    category_row,
                    category_row,
                    1,
                    2,
                    summary_fill,
                    "userEnteredFormat(backgroundColor,textFormat)",
                )
            )
            requests.append(
                _ledger_repeat_cell(
                    sheet_id,
                    category_row + 1,
                    summary_last,
                    2,
                    2,
                    {
                        "numberFormat": {
                            "type": "CURRENCY",
                            "pattern": "$#,##0.00",
                        }
                    },
                    "userEnteredFormat.numberFormat",
                )
            )
        with_sheets_retry(
            lambda: worksheet.spreadsheet.batch_update({"requests": requests})
        )

    def _format_ledger(self, worksheet: gspread.Worksheet) -> None:
        if self._cell_value(worksheet, rowcol_to_a1(1, META_LAYOUT)) == LAYOUT_VERSION:
            self._format_ledger_v6(worksheet)
            return
        self._reanchor_category_summary(worksheet)
        self._ensure_summary_gap(worksheet)
        sheet_id = worksheet.id
        totals_row = self._totals_row(worksheet)
        data_end_row = self._last_data_row(worksheet)
        header_format = {
            "backgroundColor": _sheet_rgb(LEDGER_HEADER_BG),
            "textFormat": {
                "bold": True,
                "foregroundColor": _sheet_rgb("#FFFFFF"),
            },
        }
        currency = {
            "numberFormat": {"type": "CURRENCY", "pattern": "$#,##0.00"}
        }
        stamp = {
            "numberFormat": {"type": "DATE_TIME", "pattern": "yyyy-mm-dd hh:mm:ss"}
        }
        requests: list[dict[str, Any]] = [
            {
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": sheet_id,
                        "gridProperties": {"frozenRowCount": 1},
                    },
                    "fields": "gridProperties.frozenRowCount",
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": META_OWNER - 1,
                        "endIndex": RAW_MSG,
                    },
                    "properties": {"hiddenByUser": True},
                    "fields": "hiddenByUser",
                }
            },
        ]
        for start, end in (
            (TX_TS, TX_NOTE),
            (ONEOFF_TS, ONEOFF_NOTE),
            (SHOP_TS, SHOP_NOTE),
            (INC_TS, INC_NOTE),
        ):
            requests.append(
                _ledger_repeat_cell(
                    sheet_id,
                    HEADER_ROW,
                    HEADER_ROW,
                    start,
                    end,
                    header_format,
                    "userEnteredFormat(backgroundColor,textFormat)",
                )
            )
        if data_end_row >= FIRST_DATA_ROW:
            for start, end, color in (
                (TX_TS, TX_NOTE, LEDGER_ZONE_EXPENSE),
                (ONEOFF_TS, ONEOFF_NOTE, LEDGER_ZONE_ONEOFF),
                (SHOP_TS, SHOP_NOTE, LEDGER_ZONE_SHOP),
                (INC_TS, INC_NOTE, LEDGER_ZONE_INCOME),
            ):
                requests.append(
                    _ledger_repeat_cell(
                        sheet_id,
                        FIRST_DATA_ROW,
                        data_end_row,
                        start,
                        end,
                        {"backgroundColor": _sheet_rgb(color)},
                        "userEnteredFormat.backgroundColor",
                    )
                )
        for column, size in (
            (TX_TS, 160),
            (TX_AMT, 110),
            (TX_CAT, 140),
            (TX_NOTE, 240),
            (ONEOFF_TS, 160),
            (ONEOFF_AMT, 110),
            (ONEOFF_CAT, 140),
            (ONEOFF_NOTE, 240),
            (SHOP_TS, 160),
            (SHOP_AMT, 110),
            (SHOP_NOTE, 240),
            (INC_TS, 160),
            (INC_AMT, 110),
            (INC_CAT, 140),
            (INC_NOTE, 240),
            (GAP_COL, 36),
            (GAP_COL_2, 36),
            (GAP_COL_3, 36),
            (GAP_COL_4, 36),
        ):
            requests.append(
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": column - 1,
                            "endIndex": column,
                        },
                        "properties": {"pixelSize": size},
                        "fields": "pixelSize",
                    }
                }
            )
        if data_end_row >= FIRST_DATA_ROW:
            for column in (TX_NOTE, ONEOFF_NOTE, SHOP_NOTE, INC_NOTE):
                requests.append(
                    _ledger_repeat_cell(
                        sheet_id,
                        FIRST_DATA_ROW,
                        data_end_row,
                        column,
                        column,
                        {"wrapStrategy": "WRAP"},
                        "userEnteredFormat.wrapStrategy",
                    )
                )
            for column, fmt in (
                (TX_AMT, currency),
                (ONEOFF_AMT, currency),
                (SHOP_AMT, currency),
                (INC_AMT, currency),
                (TX_TS, stamp),
                (ONEOFF_TS, stamp),
                (SHOP_TS, stamp),
                (INC_TS, stamp),
            ):
                requests.append(
                    _ledger_repeat_cell(
                        sheet_id,
                        FIRST_DATA_ROW,
                        data_end_row,
                        column,
                        column,
                        fmt,
                        "userEnteredFormat.numberFormat",
                    )
                )
        if FIRST_DATA_ROW <= totals_row < LEDGER_FORMAT_LAST_ROW:
            net_row = totals_row + SUMMARY_FOOTER_OFFSET + 2
            band_start = max(FIRST_DATA_ROW, data_end_row + 1)
            band_end = min(
                net_row + LEDGER_SUMMARY_SCRUB_TAIL,
                LEDGER_FORMAT_LAST_ROW,
            )
            requests.extend(
                self._neutral_summary_band_requests(sheet_id, band_start, band_end)
            )
            start_col = self._category_totals_start_col(worksheet, totals_row)
            width = self._category_totals_width(worksheet, totals_row)
            label_fill = {
                "backgroundColor": _sheet_rgb(LEDGER_TOTALS_BG),
                "textFormat": {"bold": True},
            }
            value_fill = {
                "backgroundColor": _sheet_rgb(LEDGER_TOTALS_BG),
                "numberFormat": {"type": "CURRENCY", "pattern": "$#,##0.00"},
            }
            for segment_start, segment_end in _ledger_column_segments(
                start_col, width
            ):
                requests.append(
                    _ledger_repeat_cell(
                        sheet_id,
                        totals_row,
                        totals_row,
                        segment_start,
                        segment_end,
                        label_fill,
                        "userEnteredFormat(backgroundColor,textFormat)",
                    )
                )
                requests.append(
                    _ledger_repeat_cell(
                        sheet_id,
                        totals_row + 1,
                        totals_row + 1,
                        segment_start,
                        segment_end,
                        value_fill,
                        "userEnteredFormat(backgroundColor,numberFormat)",
                    )
                )
            for column in range(start_col, width + 1):
                if column in (GAP_COL, GAP_COL_2, GAP_COL_3, GAP_COL_4):
                    continue
                requests.append(
                    {
                        "updateDimensionProperties": {
                            "range": {
                                "sheetId": sheet_id,
                                "dimension": "COLUMNS",
                                "startIndex": column - 1,
                                "endIndex": column,
                            },
                            "properties": {"pixelSize": LEDGER_TOTALS_COL_WIDTH},
                            "fields": "pixelSize",
                        }
                    }
                )
            gap_row = totals_row + 2
            if gap_row < LEDGER_FORMAT_LAST_ROW:
                requests.append(
                    _ledger_repeat_cell(
                        sheet_id,
                        gap_row,
                        gap_row,
                        1,
                        INC_NOTE,
                        {"backgroundColor": _sheet_rgb("#FFFFFF")},
                        "userEnteredFormat.backgroundColor",
                    )
                )
            footer_start = totals_row + SUMMARY_FOOTER_OFFSET
            if footer_start < LEDGER_FORMAT_LAST_ROW:
                summary_fill = {
                    "backgroundColor": _sheet_rgb(LEDGER_SUMMARY_BG),
                    "textFormat": {"bold": True},
                }
                requests.append(
                    _ledger_repeat_cell(
                        sheet_id,
                        footer_start,
                        footer_start + 2,
                        1,
                        2,
                        summary_fill,
                        "userEnteredFormat(backgroundColor,textFormat)",
                    )
                )
                requests.append(
                    _ledger_repeat_cell(
                        sheet_id,
                        footer_start,
                        footer_start + 2,
                        2,
                        2,
                        {
                            "backgroundColor": _sheet_rgb(LEDGER_SUMMARY_BG),
                            "numberFormat": {
                                "type": "CURRENCY",
                                "pattern": "$#,##0.00",
                            },
                        },
                        "userEnteredFormat(backgroundColor,numberFormat)",
                    )
                )
        with_sheets_retry(
            lambda: worksheet.spreadsheet.batch_update({"requests": requests})
        )

    def _format_header_row(self, worksheet: gspread.Worksheet, columns: int) -> None:
        end = rowcol_to_a1(HEADER_ROW, columns)
        with_sheets_retry(
            lambda: worksheet.format(
                f"A1:{end}",
                {"textFormat": {"bold": True}},
            )
        )

    def _format_summary_block(self, worksheet: gspread.Worksheet, start_col: int) -> None:
        value_range = {
            "sheetId": worksheet.id,
            "startRowIndex": 1,
            "endRowIndex": 4,
            "startColumnIndex": start_col,
            "endColumnIndex": start_col + 1,
        }
        requests: list[dict[str, Any]] = [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": worksheet.id,
                        "startRowIndex": 0,
                        "endRowIndex": 1,
                        "startColumnIndex": start_col - 1,
                        "endColumnIndex": start_col,
                    },
                    "cell": {"userEnteredFormat": {"textFormat": {"bold": True}}},
                    "fields": "userEnteredFormat.textFormat.bold",
                }
            },
            {
                "repeatCell": {
                    "range": value_range,
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {
                                "type": "CURRENCY",
                                "pattern": "$#,##0.00",
                            }
                        }
                    },
                    "fields": "userEnteredFormat.numberFormat",
                }
            },
        ]
        with_sheets_retry(
            lambda: worksheet.spreadsheet.batch_update({"requests": requests})
        )

    def _write_values(
        self,
        worksheet: gspread.Worksheet,
        range_name: str,
        values: list[list[Any]],
        *,
        user_entered: bool = False,
    ) -> None:
        option = ValueInputOption.user_entered if user_entered else None
        with_sheets_retry(
            lambda: worksheet.update(
                range_name,
                values,
                value_input_option=option,
            )
        )
