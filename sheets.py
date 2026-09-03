from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

import gspread
from gspread.utils import rowcol_to_a1

from google_auth import authorize_google
from parser import ParsedEntry

EXPENSE_LABEL = "EXPENSES"
INCOME_LABEL = "INCOME"
TOTAL_LABEL = "TOTAL"
NET_LABEL = "NET (Income − Expenses):"
OWNER_CELL = "Z1"

DEDUPE_COL = 27  # AA
DEDUPE_HEADERS = ["MessageId", "Date", "Type", "Category", "Amount"]

INVALID_TAB_CHARS = re.compile(r"[\[\]\\/?*]")


@dataclass
class GridSection:
    label: str
    header_row: int
    total_row: int


def _sanitize_tab_part(value: str) -> str:
    return INVALID_TAB_CHARS.sub("", value).strip()


def format_date_label(dt: datetime) -> str:
    return f"{dt.strftime('%b')} {dt.day}"


def tab_name(dt: datetime, display_name: str, user_id: int) -> str:
    month = dt.strftime("%B")
    user = _sanitize_tab_part(display_name) or str(user_id)
    name = f"{month} - {user}"
    if len(name) > 100:
        name = name[:100]
    return name


def _parse_amount(value: Any) -> float:
    if value in (None, ""):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "").replace("$", "")
    if not text:
        return 0.0
    return float(text)


class SheetManager:
    def __init__(
        self,
        sheet_id: str,
        oauth_client_path: str,
        token_path: str,
        timezone: str,
    ) -> None:
        client = authorize_google(
            Path(oauth_client_path),
            Path(token_path),
        )
        self._spreadsheet = client.open_by_key(sheet_id)
        self._tz = ZoneInfo(timezone)

    def localize(self, dt: datetime) -> datetime:
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=ZoneInfo("UTC"))
        return dt.astimezone(self._tz)

    def _worksheets_by_title(self) -> dict[str, gspread.Worksheet]:
        return {ws.title: ws for ws in self._spreadsheet.worksheets()}

    def get_or_create_user_worksheet(
        self,
        dt: datetime,
        display_name: str,
        user_id: int,
    ) -> gspread.Worksheet:
        local_dt = self.localize(dt)
        base_name = tab_name(local_dt, display_name, user_id)
        worksheets = self._worksheets_by_title()

        candidate = worksheets.get(base_name)
        if candidate and self._worksheet_owner_id(candidate) == user_id:
            return candidate

        if candidate and self._worksheet_owner_id(candidate) != user_id:
            suffix = str(user_id)[-4:]
            base_name = f"{base_name} ({suffix})"[:100]
            candidate = worksheets.get(base_name)
            if candidate and self._worksheet_owner_id(candidate) == user_id:
                return candidate

        worksheet = self._spreadsheet.add_worksheet(title=base_name, rows=200, cols=40)
        self._bootstrap_worksheet(worksheet, user_id)
        return worksheet

    def _worksheet_owner_id(self, worksheet: gspread.Worksheet) -> int | None:
        value = worksheet.acell(OWNER_CELL).value
        if not value or not value.startswith("owner:"):
            return None
        try:
            return int(value.split(":", 1)[1])
        except ValueError:
            return None

    def _bootstrap_worksheet(self, worksheet: gspread.Worksheet, user_id: int) -> None:
        worksheet.update(
            "A1:B2",
            [
                [EXPENSE_LABEL, TOTAL_LABEL],
                [TOTAL_LABEL, 0],
            ],
        )
        worksheet.update(
            "A4:B5",
            [
                [INCOME_LABEL, TOTAL_LABEL],
                [TOTAL_LABEL, 0],
            ],
        )
        worksheet.update("A7:B7", [[NET_LABEL, 0]])
        worksheet.update(OWNER_CELL, [[f"owner:{user_id}"]])
        self._ensure_dedupe_headers(worksheet)

    def _ensure_dedupe_headers(self, worksheet: gspread.Worksheet) -> None:
        start = rowcol_to_a1(1, DEDUPE_COL)
        end = rowcol_to_a1(1, DEDUPE_COL + len(DEDUPE_HEADERS) - 1)
        worksheet.update(f"{start}:{end}", [DEDUPE_HEADERS])

    def is_message_logged(self, worksheet: gspread.Worksheet, message_id: int) -> bool:
        self._ensure_dedupe_headers(worksheet)
        values = worksheet.col_values(DEDUPE_COL)
        target = str(message_id)
        return any(cell == target for cell in values[1:])

    def log_message(
        self,
        worksheet: gspread.Worksheet,
        message_id: int,
        entry_date: date,
        entry_type: str,
        category: str,
        amount: float,
    ) -> None:
        self._ensure_dedupe_headers(worksheet)
        next_row = len(worksheet.col_values(DEDUPE_COL)) + 1
        row = [
            str(message_id),
            entry_date.isoformat(),
            entry_type,
            category,
            amount,
        ]
        start = rowcol_to_a1(next_row, DEDUPE_COL)
        end = rowcol_to_a1(next_row, DEDUPE_COL + len(row) - 1)
        worksheet.update(f"{start}:{end}", [row])

    def upsert_entry(
        self,
        worksheet: gspread.Worksheet,
        local_dt: datetime,
        entry: ParsedEntry,
        category: str,
        message_id: int,
    ) -> None:
        if self.is_message_logged(worksheet, message_id):
            return

        section = (
            self._expense_section(worksheet)
            if entry.entry_type == "expense"
            else self._income_section(worksheet)
        )
        category_col = self._ensure_category_column(worksheet, section, category)
        date_row = self._ensure_date_row(worksheet, section, format_date_label(local_dt))

        current_value = worksheet.cell(date_row, category_col).value
        new_amount = _parse_amount(current_value) + entry.amount
        worksheet.update_cell(date_row, category_col, round(new_amount, 2))

        if entry.note:
            cell_label = rowcol_to_a1(date_row, category_col)
            existing_note = ""
            try:
                existing_note = worksheet.get_note(cell_label) or ""
            except Exception:
                existing_note = ""
            merged = (
                f"{existing_note}; {entry.note}"
                if existing_note
                else entry.note
            )
            worksheet.update_note(cell_label, merged)

        section = self._refresh_section(worksheet, section)
        self._rebuild_grid_totals(worksheet, section)
        self._rebuild_net(worksheet)

        self.log_message(
            worksheet,
            message_id,
            local_dt.date(),
            entry.entry_type,
            category,
            entry.amount,
        )

    def _find_section(self, worksheet: gspread.Worksheet, label: str) -> GridSection:
        header_row: int | None = None
        for row in range(1, worksheet.row_count + 1):
            cell = worksheet.cell(row, 1).value
            if cell == label:
                header_row = row
            elif header_row is not None and cell == TOTAL_LABEL:
                return GridSection(label, header_row, row)
        if label == EXPENSE_LABEL:
            return GridSection(EXPENSE_LABEL, 1, 2)
        return GridSection(INCOME_LABEL, 4, 5)

    def _expense_section(self, worksheet: gspread.Worksheet) -> GridSection:
        return self._find_section(worksheet, EXPENSE_LABEL)

    def _income_section(self, worksheet: gspread.Worksheet) -> GridSection:
        return self._find_section(worksheet, INCOME_LABEL)

    def _refresh_section(
        self, worksheet: gspread.Worksheet, section: GridSection
    ) -> GridSection:
        if section.label == EXPENSE_LABEL:
            return self._expense_section(worksheet)
        return self._income_section(worksheet)

    def _header_values(self, worksheet: gspread.Worksheet, section: GridSection) -> list[str]:
        width = self._section_width(worksheet, section)
        return worksheet.row_values(section.header_row)[:width]

    def _section_width(self, worksheet: gspread.Worksheet, section: GridSection) -> int:
        headers = worksheet.row_values(section.header_row)
        if TOTAL_LABEL not in headers:
            return max(len(headers), 2)
        return headers.index(TOTAL_LABEL) + 1

    def _category_columns(self, headers: list[str]) -> dict[str, int]:
        mapping: dict[str, int] = {}
        for index, header in enumerate(headers, start=1):
            if index == 1 or header == TOTAL_LABEL:
                continue
            if header:
                mapping[header.lower()] = index
        return mapping

    def _ensure_category_column(
        self,
        worksheet: gspread.Worksheet,
        section: GridSection,
        category: str,
    ) -> int:
        headers = self._header_values(worksheet, section)
        columns = self._category_columns(headers)
        existing = columns.get(category.lower())
        if existing:
            return existing

        total_col = headers.index(TOTAL_LABEL) + 1
        num_rows = max(worksheet.row_count, section.total_row)
        column_values = [""] * num_rows
        column_values[section.header_row - 1] = category
        worksheet.insert_cols([column_values], col=total_col, inherit_from_before=False)
        headers = self._header_values(worksheet, section)
        return self._category_columns(headers)[category.lower()]

    def _ensure_date_row(
        self,
        worksheet: gspread.Worksheet,
        section: GridSection,
        date_label: str,
    ) -> int:
        for row in range(section.header_row + 1, section.total_row):
            if worksheet.cell(row, 1).value == date_label:
                return row

        width = self._section_width(worksheet, section)
        blank_row = [""] * width
        blank_row[0] = date_label
        worksheet.insert_row(blank_row, index=section.total_row)
        section.total_row += 1
        return section.total_row - 1

    def _rebuild_grid_totals(
        self, worksheet: gspread.Worksheet, section: GridSection
    ) -> None:
        section = self._refresh_section(worksheet, section)
        headers = self._header_values(worksheet, section)
        width = len(headers)
        category_cols = [
            index
            for index, header in enumerate(headers, start=1)
            if index > 1 and header and header != TOTAL_LABEL
        ]
        total_col = headers.index(TOTAL_LABEL) + 1

        column_sums = {col: 0.0 for col in category_cols}
        for row in range(section.header_row + 1, section.total_row):
            row_sum = 0.0
            for col in category_cols:
                amount = _parse_amount(worksheet.cell(row, col).value)
                column_sums[col] += amount
                row_sum += amount
            worksheet.update_cell(row, total_col, round(row_sum, 2) if row_sum else "")

        for col in category_cols:
            total = column_sums[col]
            worksheet.update_cell(section.total_row, col, round(total, 2) if total else "")

        grand_total = sum(column_sums.values())
        worksheet.update_cell(
            section.total_row,
            total_col,
            round(grand_total, 2) if grand_total else 0,
        )

    def _expense_grand_total(self, worksheet: gspread.Worksheet) -> float:
        section = self._expense_section(worksheet)
        headers = self._header_values(worksheet, section)
        total_col = headers.index(TOTAL_LABEL) + 1
        return _parse_amount(worksheet.cell(section.total_row, total_col).value)

    def _income_grand_total(self, worksheet: gspread.Worksheet) -> float:
        section = self._income_section(worksheet)
        headers = self._header_values(worksheet, section)
        total_col = headers.index(TOTAL_LABEL) + 1
        return _parse_amount(worksheet.cell(section.total_row, total_col).value)

    def _rebuild_net(self, worksheet: gspread.Worksheet) -> None:
        net = self._income_grand_total(worksheet) - self._expense_grand_total(worksheet)
        net_row = self._find_net_row(worksheet)
        worksheet.update(f"A{net_row}:B{net_row}", [[NET_LABEL, round(net, 2)]])

    def _find_net_row(self, worksheet: gspread.Worksheet) -> int:
        for row in range(1, worksheet.row_count + 1):
            value = worksheet.cell(row, 1).value or ""
            if value.startswith("NET (Income"):
                return row
        income = self._income_section(worksheet)
        return income.total_row + 2
