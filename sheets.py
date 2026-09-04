from __future__ import annotations

import calendar
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

import gspread
from gspread.exceptions import APIError
from gspread.utils import GridRangeType, ValueInputOption, ValueRenderOption, rowcol_to_a1

from google_auth import authorize_google
from parser import ParsedEntry
from sheets_retry import with_sheets_retry

DAY_HEADER = "Day"
NOTES_HEADER = "Notes"
FIXED_HEADER = "Fixed"
INCOME_HEADER = "Income"
AMOUNT_HEADER = "Amount"
HEADER_ROW = 1
FIRST_DAY_ROW = 2
FIRST_EXPENSE_COL = 2
SIDE_PANEL_START_ROW = 2
SIDE_PANEL_END_ROW = 13
OWNER_CELL = "Z1"
FORMAT_MARKER_CELL = "Z2"
LAYOUT_MARKER_CELL = "Z3"

DEDUPE_COL = 30
DEDUPE_HEADERS = ["MessageId", "Date", "Type", "Category", "Amount"]

SUMMARY_LABELS = ("Total:", "Total Spending:", "Total Income:", "Net Total:")
FIXED_CATEGORIES = frozenset({"rent"})

INVALID_TAB_CHARS = re.compile(r"[\[\]\\/?*]")
LAYOUT_VERSION = "v5"
FORMAT_VERSION = "v8"
GAP_COLUMN_WIDTH = 32

# Fixed column anchors (1-based). Expense categories fill B..H, then insert before gap.
GAP_BEFORE_NOTES_COL = 9
NOTES_COL = 10
GAP_BEFORE_FIXED_COL = 11
FIXED_LABEL_COL = 12
FIXED_AMOUNT_COL = 13
GAP_BEFORE_INCOME_COL = 14
INCOME_LABEL_COL = 15
INCOME_AMOUNT_COL = 16
TOTAL_VALUE_COL = 2


def _rgb(red: float, green: float, blue: float) -> dict[str, float]:
    return {"red": red, "green": green, "blue": blue}


HEADER_FILL = _rgb(0.95, 0.78, 0.58)
SUMMARY_FILL = _rgb(0.94, 0.90, 0.80)
WHITE_FILL = _rgb(1.0, 1.0, 1.0)

HEADER_STYLE = {
    "backgroundColor": HEADER_FILL,
    "horizontalAlignment": "CENTER",
    "verticalAlignment": "MIDDLE",
    "textFormat": {"bold": True, "fontSize": 10},
    "wrapStrategy": "CLIP",
}

CURRENCY_STYLE = {
    "backgroundColor": WHITE_FILL,
    "numberFormat": {"type": "CURRENCY", "pattern": "$#,##0.00"},
    "horizontalAlignment": "RIGHT",
    "verticalAlignment": "MIDDLE",
}

DAY_STYLE = {
    "backgroundColor": WHITE_FILL,
    "horizontalAlignment": "CENTER",
    "verticalAlignment": "MIDDLE",
    "textFormat": {"bold": True},
}

NOTES_STYLE = {
    "backgroundColor": WHITE_FILL,
    "wrapStrategy": "WRAP",
    "horizontalAlignment": "LEFT",
    "verticalAlignment": "TOP",
}

LABEL_STYLE = {
    "backgroundColor": WHITE_FILL,
    "horizontalAlignment": "LEFT",
    "verticalAlignment": "MIDDLE",
}

SUMMARY_LABEL_STYLE = {
    "backgroundColor": SUMMARY_FILL,
    "horizontalAlignment": "RIGHT",
    "verticalAlignment": "MIDDLE",
    "textFormat": {"bold": True},
}

SUMMARY_VALUE_STYLE = {
    "backgroundColor": SUMMARY_FILL,
    "numberFormat": {"type": "CURRENCY", "pattern": "$#,##0.00"},
    "horizontalAlignment": "RIGHT",
    "verticalAlignment": "MIDDLE",
    "textFormat": {"bold": True},
}

NET_TOTAL_STYLE = {
    "backgroundColor": SUMMARY_FILL,
    "numberFormat": {"type": "CURRENCY", "pattern": "$#,##0.00"},
    "horizontalAlignment": "RIGHT",
    "verticalAlignment": "MIDDLE",
    "textFormat": {"bold": True, "fontSize": 11},
}

WHITE_STYLE = {"backgroundColor": WHITE_FILL}
GAP_COLUMN_STYLE = {"backgroundColor": WHITE_FILL}


@dataclass(frozen=True)
class MonthLayout:
    year: int
    month: int
    days: int
    summary_row: int

    @classmethod
    def for_date(cls, dt: datetime) -> MonthLayout:
        days = calendar.monthrange(dt.year, dt.month)[1]
        return cls(
            year=dt.year,
            month=dt.month,
            days=days,
            summary_row=FIRST_DAY_ROW + days + 1,
        )


@dataclass(frozen=True)
class SheetColumns:
    expense: dict[str, int]
    gap_before_notes: int
    notes_col: int
    gap_before_fixed: int
    fixed_label_col: int
    fixed_amount_col: int
    gap_before_income: int
    income_label_col: int
    income_amount_col: int


def _sanitize_tab_part(value: str) -> str:
    return INVALID_TAB_CHARS.sub("", value).strip()


def format_date_label(dt: datetime) -> str:
    return str(dt.day)


def tab_name(dt: datetime, display_name: str, user_id: int) -> str:
    month = dt.strftime("%B")
    user = _sanitize_tab_part(display_name) or str(user_id)
    suffix = str(user_id)[-4:]
    name = f"{month} - {user} ({suffix})"
    if len(name) > 100:
        name = name[:100]
    return name


def _parse_amount(value: Any) -> float:
    if value in (None, ""):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if text.startswith("="):
        return 0.0
    text = text.replace(",", "").replace("$", "")
    if not text:
        return 0.0
    return float(text)


def _is_fixed_category(category: str) -> bool:
    return category.lower() in FIXED_CATEGORIES


def append_amount_formula(current: str | None, amount: float) -> str:
    amount_text = f"{amount:.2f}".rstrip("0").rstrip(".")
    if not current or not str(current).strip():
        return f"={amount_text}"
    text = str(current).strip()
    if text.startswith("="):
        return f"{text}+{amount_text}"
    return f"={text}+{amount_text}"


def category_total_formula(col: int, layout: MonthLayout) -> str:
    start = rowcol_to_a1(FIRST_DAY_ROW, col)
    end = rowcol_to_a1(FIRST_DAY_ROW + layout.days - 1, col)
    return f"=SUM({start}:{end})"


def spending_total_formula(
    columns: SheetColumns, layout: MonthLayout, expense_cols: list[int]
) -> str:
    parts: list[str] = []
    if expense_cols:
        first = rowcol_to_a1(layout.summary_row, expense_cols[0])
        last = rowcol_to_a1(layout.summary_row, expense_cols[-1])
        parts.append(f"SUM({first}:{last})")
    fixed_start = rowcol_to_a1(SIDE_PANEL_START_ROW, columns.fixed_amount_col)
    fixed_end = rowcol_to_a1(SIDE_PANEL_END_ROW, columns.fixed_amount_col)
    parts.append(f"SUM({fixed_start}:{fixed_end})")
    return "=" + "+".join(parts)


def income_total_formula(columns: SheetColumns) -> str:
    start = rowcol_to_a1(SIDE_PANEL_START_ROW, columns.income_amount_col)
    end = rowcol_to_a1(SIDE_PANEL_END_ROW, columns.income_amount_col)
    return f"=SUM({start}:{end})"


def net_total_formula(layout: MonthLayout) -> str:
    income_cell = rowcol_to_a1(layout.summary_row + 2, TOTAL_VALUE_COL)
    spending_cell = rowcol_to_a1(layout.summary_row + 1, TOTAL_VALUE_COL)
    return f"={income_cell}-{spending_cell}"


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
        worksheets = self._worksheets_by_title()
        month = local_dt.strftime("%B")
        user = _sanitize_tab_part(display_name) or str(user_id)

        candidates = [
            tab_name(local_dt, display_name, user_id),
            f"{month} - {user}"[:100],
        ]
        seen: set[str] = set()
        for title in candidates:
            if title in seen:
                continue
            seen.add(title)
            worksheet = worksheets.get(title)
            if worksheet is None:
                continue
            owner = self._worksheet_owner_id(worksheet)
            if owner in (None, user_id):
                if owner is None:
                    self._write_values(worksheet, OWNER_CELL, [[f"owner:{user_id}"]])
                return worksheet

        for worksheet in self._spreadsheet.worksheets():
            if not worksheet.title.startswith(f"{month} - {user}"):
                continue
            owner = self._worksheet_owner_id(worksheet)
            if owner in (None, user_id):
                if owner is None:
                    self._write_values(worksheet, OWNER_CELL, [[f"owner:{user_id}"]])
                return worksheet

        title = tab_name(local_dt, display_name, user_id)
        try:
            worksheet = self._spreadsheet.add_worksheet(title=title, rows=200, cols=40)
        except APIError as error:
            if "already exists" in str(error).lower():
                worksheet = self._spreadsheet.worksheet(title)
                owner = self._worksheet_owner_id(worksheet)
                if owner in (None, user_id):
                    if owner is None:
                        self._write_values(worksheet, OWNER_CELL, [[f"owner:{user_id}"]])
                    return worksheet
            raise

        self._bootstrap_worksheet(worksheet, user_id, local_dt)
        return worksheet

    def _worksheet_owner_id(self, worksheet: gspread.Worksheet) -> int | None:
        value = with_sheets_retry(lambda: worksheet.acell(OWNER_CELL).value)
        if not value or not value.startswith("owner:"):
            return None
        try:
            return int(value.split(":", 1)[1])
        except ValueError:
            return None

    def _bootstrap_header(self) -> list[str]:
        header = [""] * INCOME_AMOUNT_COL
        header[0] = DAY_HEADER
        header[NOTES_COL - 1] = NOTES_HEADER
        header[FIXED_LABEL_COL - 1] = FIXED_HEADER
        header[FIXED_AMOUNT_COL - 1] = AMOUNT_HEADER
        header[INCOME_LABEL_COL - 1] = INCOME_HEADER
        header[INCOME_AMOUNT_COL - 1] = AMOUNT_HEADER
        return header

    def _bootstrap_worksheet(
        self,
        worksheet: gspread.Worksheet,
        user_id: int,
        local_dt: datetime,
    ) -> None:
        layout = MonthLayout.for_date(local_dt)
        header = self._bootstrap_header()
        self._write_values(
            worksheet,
            f"A{HEADER_ROW}:{rowcol_to_a1(HEADER_ROW, INCOME_AMOUNT_COL)}",
            [header],
        )

        day_values = [[day] for day in range(1, layout.days + 1)]
        self._write_values(
            worksheet,
            f"A{FIRST_DAY_ROW}:A{FIRST_DAY_ROW + layout.days - 1}",
            day_values,
        )

        self._write_values(
            worksheet,
            f"A{layout.summary_row}:A{layout.summary_row + len(SUMMARY_LABELS) - 1}",
            [[label] for label in SUMMARY_LABELS],
        )

        self._write_values(worksheet, OWNER_CELL, [[f"owner:{user_id}"]])
        self._write_values(worksheet, LAYOUT_MARKER_CELL, [[LAYOUT_VERSION]])
        self._ensure_dedupe_headers(worksheet)
        columns = self._read_columns(worksheet)
        self._apply_formatting(worksheet, layout, columns)
        self._write_values(worksheet, FORMAT_MARKER_CELL, [[FORMAT_VERSION]])
        self._rebuild_summary(worksheet, layout)

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

    def _write_formula(
        self,
        worksheet: gspread.Worksheet,
        row: int,
        col: int,
        formula: str,
    ) -> None:
        self._write_values(
            worksheet,
            rowcol_to_a1(row, col),
            [[formula]],
            user_entered=True,
        )

    def _batch_format_ranges(
        self,
        worksheet: gspread.Worksheet,
        specs: list[dict[str, Any]],
    ) -> None:
        if not specs:
            return
        with_sheets_retry(lambda: worksheet.batch_format(specs))

    def _format_range(
        self, worksheet: gspread.Worksheet, range_name: str, style: dict[str, Any]
    ) -> None:
        self._batch_format_ranges(
            worksheet, [{"range": range_name, "format": style}]
        )

    def _apply_sheet_setup(
        self,
        worksheet: gspread.Worksheet,
        widths: dict[int, int],
        heights: dict[int, int],
        *,
        unhide_columns: list[int] | None = None,
    ) -> None:
        requests: list[dict[str, Any]] = [
            {
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": worksheet.id,
                        "gridProperties": {
                            "frozenRowCount": 1,
                            "frozenColumnCount": 1,
                            "hideGridlines": False,
                        },
                    },
                    "fields": (
                        "gridProperties.frozenRowCount,"
                        "gridProperties.frozenColumnCount,"
                        "gridProperties.hideGridlines"
                    ),
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": worksheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": DEDUPE_COL - 1,
                        "endIndex": 60,
                    },
                    "properties": {"hiddenByUser": True},
                    "fields": "hiddenByUser",
                }
            },
        ]
        for column in unhide_columns or []:
            requests.append(
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": worksheet.id,
                            "dimension": "COLUMNS",
                            "startIndex": column - 1,
                            "endIndex": column,
                        },
                        "properties": {"hiddenByUser": False},
                        "fields": "hiddenByUser",
                    }
                }
            )
        for column, width in widths.items():
            requests.append(
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": worksheet.id,
                            "dimension": "COLUMNS",
                            "startIndex": column - 1,
                            "endIndex": column,
                        },
                        "properties": {"pixelSize": width},
                        "fields": "pixelSize",
                    }
                }
            )
        for row, height in heights.items():
            requests.append(
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": worksheet.id,
                            "dimension": "ROWS",
                            "startIndex": row - 1,
                            "endIndex": row,
                        },
                        "properties": {"pixelSize": height},
                        "fields": "pixelSize",
                    }
                }
            )
        with_sheets_retry(
            lambda: worksheet.spreadsheet.batch_update({"requests": requests})
        )

        with_sheets_retry(
            lambda: worksheet.spreadsheet.batch_update({"requests": requests})
        )

    def _gap_columns(self, columns: SheetColumns) -> list[int]:
        return [
            columns.gap_before_notes,
            columns.gap_before_fixed,
            columns.gap_before_income,
        ]

    def _resize_columns_for_content(
        self, worksheet: gspread.Worksheet, columns: SheetColumns
    ) -> None:
        with_sheets_retry(
            lambda: worksheet.columns_auto_resize(0, columns.income_amount_col)
        )
        requests = [
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": worksheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": column - 1,
                        "endIndex": column,
                    },
                    "properties": {"pixelSize": GAP_COLUMN_WIDTH},
                    "fields": "pixelSize",
                }
            }
            for column in self._gap_columns(columns)
        ]
        if requests:
            with_sheets_retry(
                lambda: worksheet.spreadsheet.batch_update({"requests": requests})
            )

    def _apply_formatting(
        self,
        worksheet: gspread.Worksheet,
        layout: MonthLayout,
        columns: SheetColumns,
    ) -> None:
        day_end_row = FIRST_DAY_ROW + layout.days - 1
        summary_end_row = layout.summary_row + len(SUMMARY_LABELS) - 1
        expense_cols = sorted(columns.expense.values())
        first_expense_col = expense_cols[0] if expense_cols else FIRST_EXPENSE_COL
        last_expense_col = expense_cols[-1] if expense_cols else FIRST_EXPENSE_COL
        table_end_col = columns.income_amount_col
        gap_columns = self._gap_columns(columns)

        gap_widths = {column: GAP_COLUMN_WIDTH for column in gap_columns}
        self._apply_sheet_setup(
            worksheet,
            gap_widths,
            {HEADER_ROW: 30},
            unhide_columns=gap_columns,
        )

        specs: list[dict[str, Any]] = [
            {
                "range": (
                    f"A{HEADER_ROW}:"
                    f"{rowcol_to_a1(summary_end_row, table_end_col)}"
                ),
                "format": WHITE_STYLE,
            },
        ]

        for gap_col in gap_columns:
            specs.append(
                {
                    "range": (
                        f"{rowcol_to_a1(HEADER_ROW, gap_col)}:"
                        f"{rowcol_to_a1(summary_end_row, gap_col)}"
                    ),
                    "format": GAP_COLUMN_STYLE,
                }
            )

        header_cells = ["A1"]
        header_cells.extend(rowcol_to_a1(HEADER_ROW, col) for col in expense_cols)
        header_cells.extend(
            [
                rowcol_to_a1(HEADER_ROW, columns.notes_col),
                rowcol_to_a1(HEADER_ROW, columns.fixed_label_col),
                rowcol_to_a1(HEADER_ROW, columns.fixed_amount_col),
                rowcol_to_a1(HEADER_ROW, columns.income_label_col),
                rowcol_to_a1(HEADER_ROW, columns.income_amount_col),
            ]
        )
        for cell in header_cells:
            specs.append({"range": cell, "format": HEADER_STYLE})

        if last_expense_col < columns.gap_before_notes - 1:
            specs.append(
                {
                    "range": (
                        f"{rowcol_to_a1(HEADER_ROW, last_expense_col + 1)}:"
                        f"{rowcol_to_a1(day_end_row, columns.gap_before_notes - 1)}"
                    ),
                    "format": WHITE_STYLE,
                }
            )

        specs.extend(
            [
                {
                    "range": f"A{FIRST_DAY_ROW}:A{day_end_row}",
                    "format": DAY_STYLE,
                },
                {
                    "range": (
                        f"{rowcol_to_a1(FIRST_DAY_ROW, columns.notes_col)}:"
                        f"{rowcol_to_a1(day_end_row, columns.notes_col)}"
                    ),
                    "format": NOTES_STYLE,
                },
                {
                    "range": (
                        f"{rowcol_to_a1(SIDE_PANEL_START_ROW, columns.fixed_label_col)}:"
                        f"{rowcol_to_a1(SIDE_PANEL_END_ROW, columns.fixed_label_col)}"
                    ),
                    "format": LABEL_STYLE,
                },
                {
                    "range": (
                        f"{rowcol_to_a1(SIDE_PANEL_START_ROW, columns.fixed_amount_col)}:"
                        f"{rowcol_to_a1(SIDE_PANEL_END_ROW, columns.fixed_amount_col)}"
                    ),
                    "format": CURRENCY_STYLE,
                },
                {
                    "range": (
                        f"{rowcol_to_a1(SIDE_PANEL_START_ROW, columns.income_label_col)}:"
                        f"{rowcol_to_a1(SIDE_PANEL_END_ROW, columns.income_label_col)}"
                    ),
                    "format": LABEL_STYLE,
                },
                {
                    "range": (
                        f"{rowcol_to_a1(SIDE_PANEL_START_ROW, columns.income_amount_col)}:"
                        f"{rowcol_to_a1(SIDE_PANEL_END_ROW, columns.income_amount_col)}"
                    ),
                    "format": CURRENCY_STYLE,
                },
                {
                    "range": (
                        f"{rowcol_to_a1(SIDE_PANEL_END_ROW + 1, columns.fixed_label_col)}:"
                        f"{rowcol_to_a1(day_end_row, columns.income_amount_col)}"
                    ),
                    "format": WHITE_STYLE,
                },
                {
                    "range": (
                        f"A{layout.summary_row}:"
                        f"{rowcol_to_a1(summary_end_row, table_end_col)}"
                    ),
                    "format": SUMMARY_LABEL_STYLE,
                },
            ]
        )

        if expense_cols:
            specs.append(
                {
                    "range": (
                        f"{rowcol_to_a1(FIRST_DAY_ROW, first_expense_col)}:"
                        f"{rowcol_to_a1(day_end_row, last_expense_col)}"
                    ),
                    "format": CURRENCY_STYLE,
                }
            )
            specs.append(
                {
                    "range": (
                        f"{rowcol_to_a1(layout.summary_row, first_expense_col)}:"
                        f"{rowcol_to_a1(layout.summary_row, last_expense_col)}"
                    ),
                    "format": SUMMARY_VALUE_STYLE,
                }
            )

        for offset in (1, 2):
            specs.append(
                {
                    "range": rowcol_to_a1(
                        layout.summary_row + offset, TOTAL_VALUE_COL
                    ),
                    "format": SUMMARY_VALUE_STYLE,
                }
            )
        specs.append(
            {
                "range": rowcol_to_a1(layout.summary_row + 3, TOTAL_VALUE_COL),
                "format": NET_TOTAL_STYLE,
            }
        )
        self._batch_format_ranges(worksheet, specs)
        self._resize_columns_for_content(worksheet, columns)

    def _format_expense_column(
        self,
        worksheet: gspread.Worksheet,
        layout: MonthLayout,
        column: int,
    ) -> None:
        day_end_row = FIRST_DAY_ROW + layout.days - 1
        self._batch_format_ranges(
            worksheet,
            [
                {
                    "range": rowcol_to_a1(HEADER_ROW, column),
                    "format": HEADER_STYLE,
                },
                {
                    "range": (
                        f"{rowcol_to_a1(FIRST_DAY_ROW, column)}:"
                        f"{rowcol_to_a1(day_end_row, column)}"
                    ),
                    "format": CURRENCY_STYLE,
                },
                {
                    "range": rowcol_to_a1(layout.summary_row, column),
                    "format": SUMMARY_VALUE_STYLE,
                },
            ],
        )

    def _ensure_sheet_formatting(
        self, worksheet: gspread.Worksheet, layout: MonthLayout
    ) -> None:
        marker = with_sheets_retry(
            lambda: worksheet.acell(FORMAT_MARKER_CELL).value
        )
        if marker == FORMAT_VERSION:
            return
        columns = self._read_columns(worksheet)
        self._apply_formatting(worksheet, layout, columns)
        self._write_values(worksheet, FORMAT_MARKER_CELL, [[FORMAT_VERSION]])

    def _ensure_dedupe_headers(self, worksheet: gspread.Worksheet) -> None:
        start = rowcol_to_a1(1, DEDUPE_COL)
        end = rowcol_to_a1(1, DEDUPE_COL + len(DEDUPE_HEADERS) - 1)
        self._write_values(worksheet, f"{start}:{end}", [DEDUPE_HEADERS])

    def _dedupe_column(self, worksheet: gspread.Worksheet) -> list[str]:
        values = with_sheets_retry(lambda: worksheet.col_values(DEDUPE_COL))
        if not values or values[0] != DEDUPE_HEADERS[0]:
            self._ensure_dedupe_headers(worksheet)
            values = with_sheets_retry(lambda: worksheet.col_values(DEDUPE_COL))
        return values

    def is_message_logged(self, worksheet: gspread.Worksheet, message_id: int) -> bool:
        values = self._dedupe_column(worksheet)
        return str(message_id) in values[1:]

    def log_message(
        self,
        worksheet: gspread.Worksheet,
        message_id: int,
        entry_date: date,
        entry_type: str,
        category: str,
        amount: float,
        *,
        dedupe_values: list[str] | None = None,
    ) -> None:
        values = dedupe_values if dedupe_values is not None else self._dedupe_column(worksheet)
        next_row = len(values) + 1
        row = [
            str(message_id),
            entry_date.isoformat(),
            entry_type,
            category,
            amount,
        ]
        start = rowcol_to_a1(next_row, DEDUPE_COL)
        end = rowcol_to_a1(next_row, DEDUPE_COL + len(row) - 1)
        self._write_values(worksheet, f"{start}:{end}", [row])

    def upsert_entry(
        self,
        worksheet: gspread.Worksheet,
        local_dt: datetime,
        entry: ParsedEntry,
        category: str,
        message_id: int,
    ) -> None:
        dedupe_values = self._dedupe_column(worksheet)
        if str(message_id) in dedupe_values[1:]:
            return

        layout = MonthLayout.for_date(local_dt)
        self._ensure_month_scaffold(worksheet, layout)
        self._ensure_sheet_formatting(worksheet, layout)

        summary_needed = False
        if entry.entry_type == "income":
            self._upsert_income(worksheet, category, entry.amount)
        elif _is_fixed_category(category):
            self._upsert_fixed(worksheet, category, entry.amount)
        else:
            _, summary_needed = self._upsert_daily_expense(
                worksheet,
                layout,
                local_dt.day,
                category,
                entry.amount,
                entry.note,
            )

        if summary_needed:
            self._rebuild_summary(worksheet, layout)
        self.log_message(
            worksheet,
            message_id,
            local_dt.date(),
            entry.entry_type,
            category,
            entry.amount,
            dedupe_values=dedupe_values,
        )
        columns = self._read_columns(worksheet)
        self._resize_columns_for_content(worksheet, columns)

    def _ensure_month_scaffold(
        self, worksheet: gspread.Worksheet, layout: MonthLayout
    ) -> None:
        layout_marker = with_sheets_retry(
            lambda: worksheet.acell(LAYOUT_MARKER_CELL).value
        )
        if layout_marker == LAYOUT_VERSION:
            return

        owner = self._worksheet_owner_id(worksheet) or 0
        worksheet.clear()
        self._bootstrap_worksheet(
            worksheet,
            owner,
            datetime(layout.year, layout.month, 1, tzinfo=self._tz),
        )

    def _read_columns(self, worksheet: gspread.Worksheet) -> SheetColumns:
        row = with_sheets_retry(lambda: worksheet.row_values(HEADER_ROW))
        padded = row + [""] * (INCOME_AMOUNT_COL - len(row))

        notes_col = padded.index(NOTES_HEADER) + 1 if NOTES_HEADER in padded else NOTES_COL
        gap_before_notes = notes_col - 1

        expense: dict[str, int] = {}
        for index in range(FIRST_EXPENSE_COL, gap_before_notes):
            header = padded[index - 1].strip()
            if header:
                expense[header.lower()] = index

        fixed_label_col = notes_col + 2
        fixed_amount_col = fixed_label_col + 1
        gap_before_income = fixed_amount_col + 1
        income_label_col = gap_before_income + 1
        income_amount_col = income_label_col + 1

        return SheetColumns(
            expense=expense,
            gap_before_notes=gap_before_notes,
            notes_col=notes_col,
            gap_before_fixed=notes_col + 1,
            fixed_label_col=fixed_label_col,
            fixed_amount_col=fixed_amount_col,
            gap_before_income=gap_before_income,
            income_label_col=income_label_col,
            income_amount_col=income_amount_col,
        )

    def _ensure_expense_column(
        self,
        worksheet: gspread.Worksheet,
        category: str,
        layout: MonthLayout,
    ) -> tuple[int, bool]:
        columns = self._read_columns(worksheet)
        existing = columns.expense.get(category.lower())
        if existing:
            return existing, False

        next_col = (
            max(columns.expense.values(), default=FIRST_EXPENSE_COL - 1) + 1
        )
        if next_col >= columns.gap_before_notes:
            insert_col = columns.gap_before_notes
            num_rows = max(
                len(with_sheets_retry(lambda: worksheet.col_values(1))),
                layout.summary_row + len(SUMMARY_LABELS),
            )
            column_values = [""] * num_rows
            column_values[HEADER_ROW - 1] = category
            with_sheets_retry(
                lambda: worksheet.insert_cols(
                    [column_values], col=insert_col, inherit_from_before=False
                )
            )
            self._format_expense_column(worksheet, layout, insert_col)
            return self._read_columns(worksheet).expense[category.lower()], True

        with_sheets_retry(
            lambda: worksheet.update_cell(HEADER_ROW, next_col, category)
        )
        self._format_expense_column(worksheet, layout, next_col)
        return self._read_columns(worksheet).expense[category.lower()], True

    def _day_row(self, layout: MonthLayout, day: int) -> int:
        bounded_day = max(1, min(day, layout.days))
        return FIRST_DAY_ROW + bounded_day - 1

    def _upsert_daily_expense(
        self,
        worksheet: gspread.Worksheet,
        layout: MonthLayout,
        day: int,
        category: str,
        amount: float,
        note: str | None,
    ) -> tuple[int, bool]:
        category_col, created = self._ensure_expense_column(
            worksheet, category, layout
        )
        columns = self._read_columns(worksheet)
        row = self._day_row(layout, day)
        cell = rowcol_to_a1(row, category_col)

        current = with_sheets_retry(
            lambda: worksheet.acell(cell, value_render_option=ValueRenderOption.formula).value
        )
        formula = append_amount_formula(current, amount)
        self._write_formula(worksheet, row, category_col, formula)

        if note:
            notes_cell = rowcol_to_a1(row, columns.notes_col)
            existing = (
                with_sheets_retry(lambda: worksheet.acell(notes_cell).value) or ""
            )
            merged = f"{existing}; {note}" if existing else note
            self._write_values(worksheet, notes_cell, [[merged]])

        return category_col, created
    def _upsert_fixed(
        self, worksheet: gspread.Worksheet, category: str, amount: float
    ) -> None:
        columns = self._read_columns(worksheet)
        row = self._find_side_row(
            worksheet,
            columns.fixed_label_col,
            category,
            start_row=SIDE_PANEL_START_ROW,
            end_row=SIDE_PANEL_END_ROW,
        )
        current = with_sheets_retry(
            lambda: worksheet.acell(
                rowcol_to_a1(row, columns.fixed_amount_col),
                value_render_option=ValueRenderOption.formula,
            ).value
        )
        formula = append_amount_formula(current, amount)
        self._write_formula(worksheet, row, columns.fixed_amount_col, formula)

    def _upsert_income(
        self, worksheet: gspread.Worksheet, category: str, amount: float
    ) -> None:
        columns = self._read_columns(worksheet)
        row = self._find_side_row(
            worksheet,
            columns.income_label_col,
            category,
            start_row=SIDE_PANEL_START_ROW,
            end_row=SIDE_PANEL_END_ROW,
        )
        current = with_sheets_retry(
            lambda: worksheet.acell(
                rowcol_to_a1(row, columns.income_amount_col),
                value_render_option=ValueRenderOption.formula,
            ).value
        )
        formula = append_amount_formula(current, amount)
        self._write_formula(worksheet, row, columns.income_amount_col, formula)

    def _find_side_row(
        self,
        worksheet: gspread.Worksheet,
        label_col: int,
        label: str,
        start_row: int,
        end_row: int,
    ) -> int:
        values = with_sheets_retry(lambda: worksheet.col_values(label_col))
        for row in range(start_row, end_row + 1):
            existing = values[row - 1] if row - 1 < len(values) else ""
            if existing == label:
                return row

        for row in range(start_row, end_row + 1):
            existing = values[row - 1] if row - 1 < len(values) else ""
            if not existing:
                with_sheets_retry(
                    lambda r=row: worksheet.update_cell(r, label_col, label)
                )
                return row

        row = end_row + 1
        with_sheets_retry(lambda: worksheet.update_cell(row, label_col, label))
        return row

    def _rebuild_summary(self, worksheet: gspread.Worksheet, layout: MonthLayout) -> None:
        columns = self._read_columns(worksheet)
        expense_cols = sorted(columns.expense.values())
        batch: list[dict[str, Any]] = []

        for col in expense_cols:
            batch.append(
                {
                    "range": rowcol_to_a1(layout.summary_row, col),
                    "values": [[category_total_formula(col, layout)]],
                }
            )

        batch.extend(
            [
                {
                    "range": rowcol_to_a1(layout.summary_row + 1, TOTAL_VALUE_COL),
                    "values": [
                        [spending_total_formula(columns, layout, expense_cols)]
                    ],
                },
                {
                    "range": rowcol_to_a1(layout.summary_row + 2, TOTAL_VALUE_COL),
                    "values": [[income_total_formula(columns)]],
                },
                {
                    "range": rowcol_to_a1(layout.summary_row + 3, TOTAL_VALUE_COL),
                    "values": [[net_total_formula(layout)]],
                },
            ]
        )

        with_sheets_retry(
            lambda: worksheet.batch_update(
                batch,
                value_input_option=ValueInputOption.user_entered,
            )
        )
