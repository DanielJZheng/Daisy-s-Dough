"""Find ledger rows by amount/category and show hidden Discord message ids."""

from __future__ import annotations

import argparse
import re
import sys

from config import load_settings
from gspread.utils import rowcol_to_a1

from sheets import (
    LEDGER_PREFIX,
    RAW_AMT,
    RAW_CAT,
    RAW_MSG,
    RAW_NOTE,
    RAW_TS,
    SheetManager,
    _parse_amount,
)


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit raw ledger rows")
    parser.add_argument("--user", default="cinnamon", help="Display name fragment")
    parser.add_argument("--amount", type=float, default=None)
    parser.add_argument("--category", default=None)
    args = parser.parse_args()

    settings = load_settings()
    sheets = SheetManager.from_settings(settings)
    wanted = args.user.casefold()

    tabs = [
        ws
        for ws in sheets._spreadsheet.worksheets()
        if ws.title.startswith(LEDGER_PREFIX) and wanted in ws.title.casefold()
    ]
    if not tabs:
        print(f"No ledger tab matching {args.user!r}")
        return 1

    for worksheet in tabs:
        print(f"\n=== {worksheet.title} ===")
        values = worksheet.get(
            f"{rowcol_to_a1(1, RAW_TS)}:{rowcol_to_a1(worksheet.row_count, RAW_MSG)}"
        )
        for index, row in enumerate(values or [], start=1):
            if index == 1:
                continue
            padded = row + [""] * (RAW_MSG - RAW_TS + 1 - len(row))
            stamp = str(padded[0] or "").strip()
            amount = _parse_amount(padded[RAW_AMT - RAW_TS])
            category = str(padded[RAW_CAT - RAW_TS] or "").strip()
            note = str(padded[RAW_NOTE - RAW_TS] or "").strip()
            msg_id = str(padded[RAW_MSG - RAW_TS] or "").strip()
            if not stamp and amount is None and not category:
                continue
            if args.amount is not None and amount != args.amount:
                continue
            if args.category and args.category.casefold() not in category.casefold():
                continue
            print(
                f"row {index}: {stamp!r} | ${amount} | {category!r} | note={note!r} | "
                f"discord_msg_id={msg_id or '(none)'}"
            )

    return 0


if __name__ == "__main__":
    sys.exit(main())
