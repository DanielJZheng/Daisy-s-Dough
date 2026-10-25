"""Ledger v6: manual-friendly sections, hidden spacers, no formula spill over data."""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import TYPE_CHECKING, Any

from gspread.utils import ValueInputOption, ValueRenderOption, rowcol_to_a1

from sheets_retry import with_sheets_retry

if TYPE_CHECKING:
    import gspread

    from sheets import SheetManager

LAYOUT_V6 = "ledger-v6"
META_LAYOUT_COL = 87
META_CUSTOM_COL = 84
SECTION_TITLE_ROW = 1
HEADER_ROW = 2
FIRST_DATA_ROW = 3

GAP_COLUMNS = (5, 10, 14, 19)
SUMMARY_TAIL_ROWS = 40
TIMESTAMP_LIKE = re.compile(r"^\d{4}-\d{2}-\d{2}")
ISO_STAMP = re.compile(
    r"^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?"
)
US_STAMP = re.compile(
    r"^(\d{1,2})/(\d{1,2})/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?"
)
EXCEL_EPOCH = datetime(1899, 12, 30)

SECTION_BANDS = (
    ("Recurring", 1, 4, "#FFF8F0"),
    ("One-off", 6, 9, "#E6D5C4"),
    ("Shopping", 11, 13, "#FAF3E8"),
    ("Income", 15, 18, "#E6F0E4"),
)
VISIBLE_LAST_COL = max(end for _name, _start, end, _color in SECTION_BANDS)


@dataclass(frozen=True)
class SectionBand:
    name: str
    start_col: int
    end_col: int
    color: str


def is_v6(manager: SheetManager, worksheet: gspread.Worksheet) -> bool:
    return manager._cell_value(worksheet, rowcol_to_a1(1, META_LAYOUT_COL)) == LAYOUT_V6


def row_has_content(
    manager: SheetManager,
    worksheet: gspread.Worksheet,
    row: int,
    start_col: int,
    end_col: int,
) -> bool:
    for col in range(start_col, end_col + 1):
        value = manager._cell_value(worksheet, rowcol_to_a1(row, col))
        if value is None:
            continue
        text = str(value).strip()
        if text and not text.startswith("="):
            return True
    return False


def data_bands(
    manager: SheetManager, worksheet: gspread.Worksheet
) -> list[tuple[int, int]]:
    bands = [(start, end) for _name, start, end, _color in SECTION_BANDS]
    raw = manager._cell_value(worksheet, rowcol_to_a1(1, META_CUSTOM_COL)) or "[]"
    try:
        sections: list[dict[str, Any]] = json.loads(raw)
    except json.JSONDecodeError:
        sections = []
    for section in sections:
        start = int(section["start"])
        width = int(section["width"])
        bands.append((start, start + width - 1))
    return bands


def backfill_missing_timestamps(
    manager: SheetManager,
    worksheet: gspread.Worksheet,
    *,
    moment: datetime,
) -> int:
    """Fill empty timestamp cells when amount/category/note exist in that section."""
    totals = manager._totals_row(worksheet)
    last_row = totals - 1
    if last_row < FIRST_DATA_ROW:
        return 0
    bands = data_bands(manager, worksheet)
    last_col = max(end for _start, end in bands)
    top_left = rowcol_to_a1(FIRST_DATA_ROW, 1)
    bottom_right = rowcol_to_a1(last_row, last_col)
    block = with_sheets_retry(
        lambda: worksheet.get(f"{top_left}:{bottom_right}")
    )
    if not block:
        return 0
    stamp = moment.strftime("%Y-%m-%d %H:%M:%S")
    updates: list[dict[str, Any]] = []
    for row_offset, row_vals in enumerate(block):
        row_num = FIRST_DATA_ROW + row_offset
        padded = list(row_vals) + [""] * max(0, last_col - len(row_vals))
        for start, end in bands:
            ts_idx = start - 1
            ts_val = str(padded[ts_idx] or "").strip()
            if ts_val and not ts_val.startswith("="):
                continue
            has_other = False
            for col in range(start + 1, end + 1):
                text = str(padded[col - 1] or "").strip()
                if text and not text.startswith("="):
                    has_other = True
                    break
            if has_other:
                updates.append(
                    {
                        "range": rowcol_to_a1(row_num, start),
                        "values": [[stamp]],
                    }
                )
    if not updates:
        return 0
    with_sheets_retry(
        lambda: worksheet.batch_update(
            updates, value_input_option=ValueInputOption.user_entered
        )
    )
    return len(updates)


def next_open_row(
    manager: SheetManager,
    worksheet: gspread.Worksheet,
    start_col: int,
    end_col: int,
) -> int:
    totals = manager._totals_row(worksheet)
    for row in range(FIRST_DATA_ROW, totals - 1):
        if not row_has_content(manager, worksheet, row, start_col, end_col):
            return row
    spacer = totals - 1
    manager._insert_row(worksheet, spacer)
    manager._set_totals_row(worksheet, totals + 1)
    return spacer


def write_headers(manager: SheetManager, worksheet: gspread.Worksheet) -> None:
    from sheets import (
        INC_HEADERS,
        INC_TS,
        ONEOFF_HEADERS,
        ONEOFF_TS,
        SHOP_HEADERS,
        SHOP_TS,
        TX_HEADERS,
    )

    titles = [""] * 18
    headers = [""] * 18
    for name, start, end, _color in SECTION_BANDS:
        titles[start - 1] = name
    headers[0:4] = TX_HEADERS
    headers[ONEOFF_TS - 1 : ONEOFF_TS + 3] = ONEOFF_HEADERS
    headers[SHOP_TS - 1 : SHOP_TS + 2] = SHOP_HEADERS
    headers[INC_TS - 1 : INC_TS + 3] = INC_HEADERS
    manager._write_values(worksheet, f"A{SECTION_TITLE_ROW}", [titles])
    manager._write_values(worksheet, f"A{HEADER_ROW}", [headers])


def clear_formula_spill(manager: SheetManager, worksheet: gspread.Worksheet) -> None:
    from sheets import ONEOFF_TS, TX_TS

    manager._write_values(worksheet, rowcol_to_a1(FIRST_DATA_ROW, TX_TS), [[""]])
    manager._write_values(worksheet, rowcol_to_a1(FIRST_DATA_ROW, ONEOFF_TS), [[""]])


def _col_letter(column: int) -> str:
    return rowcol_to_a1(1, column)[:-1]


def expense_total_formula(end_row: int) -> str:
    """Sum amount columns in Recurring, One-off, and Shopping (not timestamps)."""
    start = FIRST_DATA_ROW
    parts: list[str] = []
    for name, band_start, _band_end, _color in SECTION_BANDS:
        if name == "Income":
            continue
        letter = _col_letter(band_start + 1)
        parts.append(f"SUM({letter}{start}:{letter}{end_row})")
    return "=" + "+".join(parts)


def income_total_formula(end_row: int) -> str:
    start = FIRST_DATA_ROW
    for name, band_start, _band_end, _color in SECTION_BANDS:
        if name != "Income":
            continue
        letter = _col_letter(band_start + 1)
        return f"=SUM({letter}{start}:{letter}{end_row})"
    raise ValueError("Income section missing from SECTION_BANDS")


def footer_expense_row(totals_row: int) -> int:
    return totals_row + 3


def category_block_row(totals_row: int) -> int:
    return footer_expense_row(totals_row) + 4


def summary_tail_range(end_row: int, *, rows: int = SUMMARY_TAIL_ROWS) -> tuple[str, int, int]:
    start = max(end_row + 1, FIRST_DATA_ROW)
    stop = start + rows - 1
    return f"A{start}:{_col_letter(VISIBLE_LAST_COL)}{stop}", start, stop


def clear_summary_tail(
    manager: SheetManager,
    worksheet: gspread.Worksheet,
    end_row: int,
    *,
    rows: int = SUMMARY_TAIL_ROWS,
) -> None:
    """Wipe old footer/category cells so a shorter rewrite cannot leave leftovers."""
    a1, _start, _stop = summary_tail_range(end_row, rows=rows)
    with_sheets_retry(lambda: worksheet.batch_clear([a1]))


def _from_serial(serial: float) -> datetime | None:
    if serial < 20000 or serial >= 60000:
        return None
    return EXCEL_EPOCH + timedelta(days=serial)


def parse_sheet_datetime(value: Any) -> datetime | None:
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, (int, float)):
        return _from_serial(float(value))
    text = str(value).strip()
    match = ISO_STAMP.match(text)
    if match:
        year, month, day, hour, minute, second = match.groups()
        return datetime(
            int(year),
            int(month),
            int(day),
            int(hour or 0),
            int(minute or 0),
            int(second or 0),
        )
    match = US_STAMP.match(text)
    if match:
        month, day, year, hour, minute, second = match.groups()
        return datetime(
            int(year),
            int(month),
            int(day),
            int(hour or 0),
            int(minute or 0),
            int(second or 0),
        )
    try:
        return _from_serial(float(text.replace(",", "")))
    except ValueError:
        return None


def format_stamp(value: Any) -> str:
    parsed = parse_sheet_datetime(value)
    if parsed is None:
        return str(value or "").strip()
    return parsed.strftime("%Y-%m-%d %H:%M:%S")


def is_timestamp(value: Any) -> bool:
    return parse_sheet_datetime(value) is not None


def _stamp_sort_key(stamp: Any) -> tuple[int, datetime | str]:
    parsed = parse_sheet_datetime(stamp)
    if parsed is not None:
        return (0, parsed)
    return (1, str(stamp or "").strip().casefold())


def _amount_cell(value: Any) -> Any:
    if value in (None, ""):
        return ""
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace("$", "").replace(",", "").strip()
    try:
        return float(text)
    except ValueError:
        return value


def sort_band_rows(
    rows: list[list[Any]],
    message_ids: list[Any],
    *,
    width: int = 4,
) -> tuple[list[list[Any]], list[Any]]:
    """Pack real rows oldest-first. Drop overlay junk. Keep message IDs aligned."""
    filled: list[tuple[list[Any], Any, int]] = []
    for index, row in enumerate(rows):
        padded = list(row) + [""] * max(0, width - len(row))
        padded = padded[:width]
        stamp = padded[0] if padded else ""
        amount = _amount_cell(padded[1]) if width > 1 else ""
        has_stamp = is_timestamp(stamp)
        has_amount = isinstance(amount, float)
        if not has_stamp and not has_amount:
            continue
        if has_stamp:
            padded[0] = format_stamp(stamp)
        if width > 1 and has_amount:
            padded[1] = amount
        msg = message_ids[index] if index < len(message_ids) else ""
        filled.append((padded, msg, index))
    filled.sort(key=lambda item: (_stamp_sort_key(item[0][0]), item[2]))
    blank = [""] * width
    visible = [item[0] for item in filled]
    msgs = [item[1] for item in filled]
    missing = max(0, len(rows) - len(visible))
    visible.extend([list(blank) for _ in range(missing)])
    msgs.extend([""] * missing)
    return visible, msgs


def rebalance_expense_rows(
    rec_rows: list[list[Any]],
    rec_msgs: list[Any],
    one_rows: list[list[Any]],
    one_msgs: list[Any],
    *,
    width: int = 4,
) -> tuple[list[list[Any]], list[Any], list[list[Any]], list[Any]]:
    """One stem lives in one band. Count>=2 → Recurring. Count==1 → One-off. Normalize names."""
    from collections import defaultdict

    from sheets import category_stem, preferred_category

    packed: list[tuple[list[Any], Any]] = []
    for rows, msgs in ((rec_rows, rec_msgs), (one_rows, one_msgs)):
        visible, out_msgs = sort_band_rows(rows, msgs, width=width)
        for row, msg in zip(visible, out_msgs, strict=True):
            if not any(str(cell).strip() for cell in row):
                continue
            packed.append((list(row), msg))

    labels_by_stem: dict[str, list[str]] = defaultdict(list)
    for row, _msg in packed:
        category = str(row[2] or "").strip() if len(row) > 2 else ""
        if not category:
            continue
        labels_by_stem[category_stem(category)].append(category)
    canonical = {
        stem: preferred_category(labels) for stem, labels in labels_by_stem.items()
    }
    counts = {stem: len(labels) for stem, labels in labels_by_stem.items()}

    rec: list[tuple[list[Any], Any]] = []
    one: list[tuple[list[Any], Any]] = []
    for row, msg in packed:
        category = str(row[2] or "").strip() if len(row) > 2 else ""
        stem = category_stem(category) if category else ""
        if stem and stem in canonical:
            row[2] = canonical[stem]
        if stem and counts.get(stem, 0) >= 2:
            rec.append((row, msg))
        else:
            one.append((row, msg))
    rec.sort(key=lambda item: (_stamp_sort_key(item[0][0]), str(item[0][2]), item[0][1]))
    one.sort(key=lambda item: (_stamp_sort_key(item[0][0]), str(item[0][2]), item[0][1]))
    return (
        [item[0] for item in rec],
        [item[1] for item in rec],
        [item[0] for item in one],
        [item[1] for item in one],
    )


def sort_expense_bands(
    manager: SheetManager,
    worksheet: gspread.Worksheet,
    end_row: int,
) -> None:
    from sheets import (
        ONEOFF_MSG_COL,
        ONEOFF_NOTE,
        ONEOFF_TS,
        TX_MSG_COL,
        TX_NOTE,
        TX_TS,
    )

    totals = manager._totals_row(worksheet)
    last = min(max(end_row, FIRST_DATA_ROW), max(totals - 1, FIRST_DATA_ROW))
    if last < FIRST_DATA_ROW:
        return
    _sort_one_band(manager, worksheet, TX_TS, TX_NOTE, TX_MSG_COL, last)
    _sort_one_band(manager, worksheet, ONEOFF_TS, ONEOFF_NOTE, ONEOFF_MSG_COL, last)


def rebalance_expense_bands(
    manager: SheetManager,
    worksheet: gspread.Worksheet,
    end_row: int,
) -> None:
    from sheets import (
        ONEOFF_MSG_COL,
        ONEOFF_NOTE,
        ONEOFF_TS,
        TX_MSG_COL,
        TX_NOTE,
        TX_TS,
    )

    totals = manager._totals_row(worksheet)
    last = max(end_row, totals - 2, FIRST_DATA_ROW)
    if last < FIRST_DATA_ROW:
        return
    rec_rows, rec_msgs = _read_band(
        manager, worksheet, TX_TS, TX_NOTE, TX_MSG_COL, last
    )
    one_rows, one_msgs = _read_band(
        manager, worksheet, ONEOFF_TS, ONEOFF_NOTE, ONEOFF_MSG_COL, last
    )
    new_rec, new_rec_msgs, new_one, new_one_msgs = rebalance_expense_rows(
        rec_rows, rec_msgs, one_rows, one_msgs
    )
    write_last = max(
        last,
        FIRST_DATA_ROW + len(new_rec) - 1 if new_rec else last,
        FIRST_DATA_ROW + len(new_one) - 1 if new_one else last,
    )
    _write_band(
        manager, worksheet, TX_TS, TX_NOTE, TX_MSG_COL, write_last, new_rec, new_rec_msgs
    )
    _write_band(
        manager,
        worksheet,
        ONEOFF_TS,
        ONEOFF_NOTE,
        ONEOFF_MSG_COL,
        write_last,
        new_one,
        new_one_msgs,
    )


def _read_band(
    manager: SheetManager,
    worksheet: gspread.Worksheet,
    start_col: int,
    end_col: int,
    msg_col: int,
    last_row: int,
) -> tuple[list[list[Any]], list[Any]]:
    vis_range = (
        f"{rowcol_to_a1(FIRST_DATA_ROW, start_col)}:{rowcol_to_a1(last_row, end_col)}"
    )
    msg_range = (
        f"{rowcol_to_a1(FIRST_DATA_ROW, msg_col)}:{rowcol_to_a1(last_row, msg_col)}"
    )
    visible = with_sheets_retry(
        lambda: worksheet.get(
            vis_range,
            value_render_option=ValueRenderOption.formatted,
        )
    ) or []
    raw_msgs = with_sheets_retry(lambda: worksheet.get(msg_range)) or []
    count = last_row - FIRST_DATA_ROW + 1
    rows: list[list[Any]] = []
    message_ids: list[Any] = []
    for index in range(count):
        rows.append(list(visible[index]) if index < len(visible) else [])
        if index < len(raw_msgs) and raw_msgs[index]:
            message_ids.append(raw_msgs[index][0])
        else:
            message_ids.append("")
    return rows, message_ids


def _write_band(
    manager: SheetManager,
    worksheet: gspread.Worksheet,
    start_col: int,
    end_col: int,
    msg_col: int,
    last_row: int,
    rows: list[list[Any]],
    message_ids: list[Any],
) -> None:
    width = end_col - start_col + 1
    count = last_row - FIRST_DATA_ROW + 1
    padded_rows = [list(row) for row in rows]
    padded_rows.extend([""] * width for _ in range(max(0, count - len(padded_rows))))
    padded_rows = padded_rows[:count]
    padded_msgs = list(message_ids) + [""] * max(0, count - len(message_ids))
    padded_msgs = padded_msgs[:count]
    vis_range = (
        f"{rowcol_to_a1(FIRST_DATA_ROW, start_col)}:{rowcol_to_a1(last_row, end_col)}"
    )
    msg_range = (
        f"{rowcol_to_a1(FIRST_DATA_ROW, msg_col)}:{rowcol_to_a1(last_row, msg_col)}"
    )
    manager._write_values(worksheet, vis_range, padded_rows, user_entered=True)
    manager._write_values(worksheet, msg_range, [[msg] for msg in padded_msgs])


def _sort_one_band(
    manager: SheetManager,
    worksheet: gspread.Worksheet,
    start_col: int,
    end_col: int,
    msg_col: int,
    last_row: int,
) -> None:
    width = end_col - start_col + 1
    rows, message_ids = _read_band(
        manager, worksheet, start_col, end_col, msg_col, last_row
    )
    new_rows, new_msgs = sort_band_rows(rows, message_ids, width=width)
    _write_band(
        manager, worksheet, start_col, end_col, msg_col, last_row, new_rows, new_msgs
    )


def refresh_footer(manager: SheetManager, worksheet: gspread.Worksheet, end_row: int) -> None:
    totals_row = manager._sync_totals_row(worksheet)
    expense_row = footer_expense_row(totals_row)
    income_row = expense_row + 1
    net_row = income_row + 1
    updates = [
        {"range": f"A{expense_row}", "values": [["Total expense"]]},
        {
            "range": f"B{expense_row}",
            "values": [[expense_total_formula(end_row)]],
        },
        {"range": f"A{income_row}", "values": [["Total income"]]},
        {
            "range": f"B{income_row}",
            "values": [[income_total_formula(end_row)]],
        },
        {"range": f"A{net_row}", "values": [["Net"]]},
        {"range": f"B{net_row}", "values": [[f"=B{income_row}-B{expense_row}"]]},
    ]
    with_sheets_retry(
        lambda: worksheet.batch_update(
            updates, value_input_option=ValueInputOption.user_entered
        )
    )


def migrate_to_v6(manager: SheetManager, worksheet: gspread.Worksheet) -> None:
    """Insert section title row, stop formula spill, mark v6."""
    with_sheets_retry(
        lambda: worksheet.insert_row([""] * worksheet.col_count, index=1, inherit_from_before=False)
    )
    write_headers(manager, worksheet)
    clear_formula_spill(manager, worksheet)
    manager._write_values(worksheet, rowcol_to_a1(1, META_LAYOUT_COL), [[LAYOUT_V6]])
    manager._format_ledger_v6(worksheet)


def add_custom_section(manager: SheetManager, worksheet: gspread.Worksheet, name: str) -> int:
    clean = re.sub(r"\s+", " ", name).strip()[:30]
    if not clean:
        raise ValueError("Section name required")
    raw = manager._cell_value(worksheet, rowcol_to_a1(1, META_CUSTOM_COL)) or "[]"
    try:
        sections: list[dict[str, Any]] = json.loads(raw)
    except json.JSONDecodeError:
        sections = []
    start = 20
    if sections:
        last = sections[-1]
        start = int(last["start"]) + int(last["width"])
    if start + 4 > 40:
        raise ValueError("No room for more custom sections")
    width = 4
    with_sheets_retry(
        lambda: worksheet.insert_cols(
            [[""] * max(worksheet.row_count, 50) for _ in range(width)],
            col=start,
            inherit_from_before=False,
        )
    )
    sections.append({"name": clean, "start": start, "width": width})
    manager._write_values(
        worksheet, rowcol_to_a1(1, META_CUSTOM_COL), [[json.dumps(sections)]]
    )
    manager._write_values(
        worksheet,
        rowcol_to_a1(SECTION_TITLE_ROW, start),
        [[clean, "", "", ""]],
    )
    manager._write_values(
        worksheet,
        rowcol_to_a1(HEADER_ROW, start),
        [["Timestamp", "Amount", "Category", "Note"]],
    )
    manager._format_ledger_v6(worksheet)
    return start
