"""Local smoke test before cloud deploy. No Discord required."""

from __future__ import annotations

import argparse
import sys
from datetime import datetime

from bot import CONFIRM_EMOJI, REJECT_EMOJI, confirmation_logs_entry
from config import load_settings
from gspread.exceptions import APIError
from ledger_v6 import (
    expense_total_formula,
    format_stamp,
    income_total_formula,
    parse_sheet_datetime,
    rebalance_expense_rows,
    sort_band_rows,
    summary_tail_range,
)
from parser import parse_entries, parse_message, parse_sudo
from sheets import (
    CATEGORY_HEADERS,
    CATEGORIES_SHEET,
    SheetManager,
    preferred_category,
    resolve_keyword,
    tab_name,
)


def check_config() -> tuple[object, None] | tuple[None, int]:
    try:
        settings = load_settings()
    except Exception as error:
        print(f"[FAIL] config: {error}")
        return None, 1
    print("[ok] config loaded")
    return settings, None


def check_parser() -> int | None:
    samples = [
        ("105 seven", "expense", "seven", 105.0, None),
        ("43 coco", "expense", "coco", 43.0, None),
        ("167 burger oat milk", "expense", "burger", 167.0, "oat milk"),
        ("-42 groceries oat", "expense", "groceries", 42.0, "oat"),
        ("+1200 paycheck bonus", "income", "paycheck", 1200.0, "bonus"),
        ("$25.50 water", "expense", "water", 25.5, None),
    ]
    for text, entry_type, description, amount, note in samples:
        entry = parse_message(text)
        if entry is None:
            print(f"[FAIL] parser returned None for: {text!r}")
            return 1
        if (
            entry.entry_type != entry_type
            or entry.description != description
            or entry.amount != amount
            or entry.note != note
        ):
            print(f"[FAIL] parser mismatch for: {text!r} -> {entry}")
            return 1

    rejected = ["income 2500 salary", "balance 5234.56", "0 seven", "hello"]
    for text in rejected:
        if parse_message(text) is not None:
            print(f"[FAIL] parser should reject: {text!r}")
            return 1
    sudo = parse_sudo("sudo @cinnamon +1200 paycheck bonus")
    if sudo is None or sudo.user_id is not None or sudo.name != "cinnamon":
        print(f"[FAIL] sudo name parse: {sudo}")
        return 1
    mentioned = parse_sudo("sudo <@123456> 43 shopping milk")
    if mentioned is None or mentioned.user_id != 123456 or mentioned.body != "43 shopping milk":
        print(f"[FAIL] sudo mention parse: {mentioned}")
        return 1
    if parse_message(mentioned.body) is None or parse_message(sudo.body) is None:
        print("[FAIL] sudo body did not parse as an entry")
        return 1
    batch = """sudo cinnamon
$288.00 groceries
Income
FT 41309
Week 1 2400
59309
Fixed
Rent 14000
Others
First day Tw 1889
"""
    target, entries = parse_entries(batch)
    expected = [
        ("expense", 288.0, "groceries"),
        ("income", 41309.0, "FT"),
        ("income", 2400.0, "Week 1"),
        ("expense", 14000.0, "Rent"),
        ("expense", 1889.0, "First day Tw"),
    ]
    got = [(entry.entry_type, entry.amount, entry.description) for entry in entries]
    if target is None or target.name != "cinnamon" or got != expected:
        print(f"[FAIL] batch parse: {target} {got}")
        return 1
    print("[ok] parser samples")

    keywords = [
        ("seven", "Convenience"),
        ("burger", "Food"),
        ("water", "Food"),
        ("coco", "Food"),
    ]
    fuzzy_samples = [
        ("burger", "exact", "burger"),
        ("burgur", "fuzzy", "burger"),
        ("Watter", "fuzzy", "water"),
        ("zzzz", "none", None),
    ]
    for raw, kind, keyword in fuzzy_samples:
        match = resolve_keyword(raw, keywords, 0.75)
        if match.kind != kind or match.keyword != keyword:
            print(f"[FAIL] fuzzy mismatch for: {raw!r} -> {match}")
            return 1
    print("[ok] fuzzy samples")
    return None


def check_footer_formulas() -> int | None:
    formula = expense_total_formula(10)
    if "SUM(F3:F10)" in formula or "SUM(K3:K10)" in formula:
        print(f"[FAIL] expense footer still sums timestamp columns: {formula}")
        return 1
    expected = "=SUM(B3:B10)+SUM(G3:G10)+SUM(L3:L10)"
    if formula != expected:
        print(f"[FAIL] expense footer: {formula!r} != {expected!r}")
        return 1
    income = income_total_formula(10)
    if income != "=SUM(P3:P10)":
        print(f"[FAIL] income footer: {income!r}")
        return 1
    print("[ok] footer formulas — amounts B+G+L, income P")
    a1, start, stop = summary_tail_range(33)
    if a1 != "A34:R73" or start != 34 or stop != 73:
        print(f"[FAIL] summary tail range: {a1} {start} {stop}")
        return 1
    print("[ok] summary tail clears leftover footer rows")
    return None


def check_sort_band_rows() -> int | None:
    rows = [
        ["2026-10-11 15:47:52", "$11.98", "TooGoodToGo", "salmon"],
        ["2026-09-27 15:15:15", "$20.00", "entertainment", "cursor"],
        ["Shopping", "dining", "Groceries", "Honey"],
        ["2026-09-28 20:01:53", "$6.00", "pretzels", "bread"],
    ]
    msgs = ["111", "222", "", "333"]
    visible, out_msgs = sort_band_rows(rows, msgs)
    got = [row[2] for row in visible[:3]]
    if got != ["entertainment", "pretzels", "TooGoodToGo"]:
        print(f"[FAIL] sort order: {got}")
        return 1
    if visible[3] != ["", "", "", ""]:
        print(f"[FAIL] junk row should be dropped: {visible[3]}")
        return 1
    if out_msgs[:3] != ["222", "333", "111"]:
        print(f"[FAIL] message ids not aligned: {out_msgs[:3]}")
        return 1
    if visible[0][1] != 20.0:
        print(f"[FAIL] amount not numeric after sort: {visible[0][1]!r}")
        return 1
    serial = parse_sheet_datetime("46293.22895")
    if serial is None or serial.year != 2026:
        print(f"[FAIL] excel serial parse: {serial}")
        return 1
    kept, kept_msgs = sort_band_rows(
        [["46293.22895", "209", "groceries", ""]],
        ["999"],
    )
    if kept[0][2] != "groceries" or kept_msgs[0] != "999":
        print(f"[FAIL] serial row dropped: {kept[0]} {kept_msgs}")
        return 1
    if not format_stamp("46293.22895").startswith("2026-"):
        print(f"[FAIL] serial stamp format: {format_stamp('46293.22895')}")
        return 1
    print("[ok] recurring/one-off sort by timestamp")
    return None


def check_rebalance_and_names() -> int | None:
    if preferred_category(["honey", "Honey"]) != "Honey":
        print("[FAIL] preferred_category should keep the capital spelling on a tie")
        return 1
    if preferred_category(["groceries", "groceries", "Groceries"]) != "groceries":
        print("[FAIL] preferred_category should keep the more common spelling")
        return 1
    rec = [
        ["2026-10-01 10:00:00", "10", "groceries", "a"],
        ["2026-10-03 10:00:00", "20", "Groceries", "c"],
        ["2026-10-05 10:00:00", "7", "honey", ""],
    ]
    one = [
        ["2026-10-02 10:00:00", "15", "groceries", "b"],
        ["2026-10-04 10:00:00", "5", "Rent", ""],
        ["2026-10-06 10:00:00", "8", "Honey", "x"],
    ]
    new_rec, rec_msgs, new_one, one_msgs = rebalance_expense_rows(
        rec, ["1", "3", "5"], one, ["2", "4", "6"]
    )
    rec_cats = [row[2] for row in new_rec]
    one_cats = [row[2] for row in new_one]
    if rec_cats.count("groceries") != 3:
        print(f"[FAIL] groceries should all sit in Recurring: {rec_cats}")
        return 1
    if "groceries" in one_cats or "Groceries" in one_cats:
        print(f"[FAIL] groceries leftover in One-off: {one_cats}")
        return 1
    if rec_cats.count("Honey") != 2 or "honey" in rec_cats:
        print(f"[FAIL] honey/Honey should normalize into Recurring as Honey: {rec_cats}")
        return 1
    if one_cats != ["Rent"]:
        print(f"[FAIL] singleton Rent should stay One-off: {one_cats}")
        return 1
    if rec_msgs != ["1", "2", "3", "5", "6"] or one_msgs != ["4"]:
        print(f"[FAIL] message ids misaligned: {rec_msgs} {one_msgs}")
        return 1
    rec_tithe = [["2026-10-24 18:00:00", "10", "Tithe", "new"]]
    one_tithe = [["2026-10-09 00:43:19", "104.12", "Tithe", "old"]]
    moved_rec, _, moved_one, _ = rebalance_expense_rows(
        rec_tithe, ["n"], one_tithe, ["o"]
    )
    if [row[2] for row in moved_rec] != ["Tithe", "Tithe"] or moved_one:
        print(f"[FAIL] second Tithe should pull the one-off into Recurring: {moved_rec} {moved_one}")
        return 1
    print("[ok] rebalance bands and normalize names")
    return None


def check_confirmation_reaction() -> int | None:
    if not confirmation_logs_entry(CONFIRM_EMOJI):
        print("[FAIL] ✅ should log the suggested entry")
        return 1
    if confirmation_logs_entry(REJECT_EMOJI):
        print("[FAIL] ❌ should not log after declining a suggestion")
        return 1
    print("[ok] fuzzy reject does not write")
    return None


def check_sheets(settings: object) -> tuple[SheetManager | None, int | None]:
    try:
        sheets = SheetManager.from_settings(settings)
        title = sheets._spreadsheet.title
        print(f"[ok] google auth — spreadsheet {title!r}")
    except Exception as error:
        print(f"[FAIL] google sheets: {error}")
        return None, 1
    return sheets, None


def check_categories(sheets: SheetManager) -> int | None:
    try:
        sheets.ensure_foundation()
        worksheet = sheets._spreadsheet.worksheet(CATEGORIES_SHEET)
        headers = worksheet.row_values(1)
    except Exception as error:
        print(f"[FAIL] categories: {error}")
        return 1
    if headers[:2] != CATEGORY_HEADERS:
        print(f"[FAIL] categories headers: {headers[:2]!r}")
        return 1
    print("[ok] categories map — Keyword, Category")
    return None


def check_write(sheets: SheetManager) -> int | None:
    now = datetime.now().astimezone()
    local_dt = sheets.localize(now)
    display_name = "VerifyBot"
    user_id = 9999999994197
    message_id = 9999999999999
    entry = parse_message("0.01 verify-test smoke check")
    if entry is None:
        print("[FAIL] write test: parser failed on verify message")
        return 1

    worksheet = sheets.get_or_create_user_worksheet(display_name, user_id)
    tab = tab_name(display_name, user_id)

    if sheets.is_message_logged(worksheet, message_id):
        print(f"[ok] write test skipped — already logged on tab {worksheet.title!r}")
        return None

    try:
        sheets.append_entry(worksheet, local_dt, entry, message_id)
        if not sheets.is_message_logged(worksheet, message_id):
            print("[FAIL] write test: entry not found in message id column")
            return 1
        print(f"[ok] write test — saved $0.01 to tab {tab!r} (safe to delete)")
    except APIError as error:
        if "429" in str(error):
            print("[WARN] write test: Google rate limit — wait 60s, retry --write-test")
            return None
        print(f"[FAIL] write test: {error}")
        return 1
    except Exception as error:
        print(f"[FAIL] write test: {error}")
        return 1
    return None


def main() -> int:
    parser = argparse.ArgumentParser(description="Verify Daisy's Dough setup")
    parser.add_argument(
        "--write-test",
        action="store_true",
        help="Write a $0.01 test row to the VerifyBot ledger tab",
    )
    args = parser.parse_args()

    print("=== Daisy's Dough verify ===\n")

    settings, code = check_config()
    if code is not None:
        return code

    code = check_parser()
    if code is not None:
        return code

    code = check_footer_formulas()
    if code is not None:
        return code

    code = check_sort_band_rows()
    if code is not None:
        return code

    code = check_rebalance_and_names()
    if code is not None:
        return code

    code = check_confirmation_reaction()
    if code is not None:
        return code

    sheets, code = check_sheets(settings)
    if code is not None:
        return code

    code = check_categories(sheets)
    if code is not None:
        return code

    if args.write_test:
        code = check_write(sheets)
        if code is not None:
            return code

    print("\nAll automated checks passed.")
    if not args.write_test:
        print("Run with --write-test to confirm sheet writes end-to-end.")
    print("Then run python bot.py and complete the Discord checklist.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
