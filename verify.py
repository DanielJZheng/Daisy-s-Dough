"""Local smoke test before cloud deploy. No Discord required."""

from __future__ import annotations

import argparse
import sys
from datetime import datetime

from categories import CategoryRegistry
from config import load_settings
from gspread.exceptions import APIError
from parser import parse_message
from sheets import SheetManager, tab_name


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
        ("6 groceries test", "expense", "groceries", 6.0),
        ("25.50 groceries Costco run", "expense", "groceries", 25.5),
        ("income 2500 salary", "income", "salary", 2500.0),
        ("700 rent", "expense", "rent", 700.0),
    ]
    for text, entry_type, category, amount in samples:
        entry = parse_message(text)
        if entry is None:
            print(f"[FAIL] parser returned None for: {text!r}")
            return 1
        if (
            entry.entry_type != entry_type
            or entry.category != category
            or entry.amount != amount
        ):
            print(f"[FAIL] parser mismatch for: {text!r} -> {entry}")
            return 1
    print("[ok] parser samples")
    return None


def check_sheets(settings: object) -> tuple[SheetManager | None, int | None]:
    try:
        sheets = SheetManager(
            settings.google_sheet_id,
            str(settings.google_oauth_client_path),
            str(settings.google_token_path),
            settings.timezone,
        )
        title = sheets._spreadsheet.title
        print(f"[ok] google auth — spreadsheet {title!r}")
    except Exception as error:
        print(f"[FAIL] google sheets: {error}")
        return None, 1
    return sheets, None


def check_categories(
    sheets: SheetManager, settings: object
) -> tuple[CategoryRegistry | None, int | None]:
    try:
        registry = CategoryRegistry(sheets._spreadsheet, settings.fuzzy_match_cutoff)
        match = registry.resolve("groceries", "expense")
        print(f"[ok] categories read — groceries -> {match.kind}")
    except Exception as error:
        print(f"[FAIL] categories: {error}")
        return None, 1
    return registry, None


def check_write(
    sheets: SheetManager,
    settings: object,
    registry: CategoryRegistry,
) -> int | None:
    now = datetime.now().astimezone()
    local_dt = sheets.localize(now)
    display_name = "VerifyBot"
    user_id = 9999999994197
    message_id = 9999999999999
    entry = parse_message("0.01 verify-test smoke check")
    if entry is None:
        print("[FAIL] write test: parser failed on verify message")
        return 1

    worksheet = sheets.get_or_create_user_worksheet(local_dt, display_name, user_id)
    tab = tab_name(local_dt, display_name, user_id)
    category = registry.add_category("verify-test", "expense", first_seen=local_dt.date())

    if sheets.is_message_logged(worksheet, message_id):
        print(f"[ok] write test skipped — already logged on tab {tab!r}")
        return None

    try:
        sheets.upsert_entry(worksheet, local_dt, entry, category, message_id)
        if not sheets.is_message_logged(worksheet, message_id):
            print("[FAIL] write test: entry not found in dedupe log")
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
        help="Write a $0.01 test row to your monthly tab",
    )
    args = parser.parse_args()

    print("=== Daisy's Dough verify ===\n")

    settings, code = check_config()
    if code is not None:
        return code

    code = check_parser()
    if code is not None:
        return code

    sheets, code = check_sheets(settings)
    if code is not None:
        return code

    registry, code = check_categories(sheets, settings)
    if code is not None:
        return code

    if args.write_test:
        code = check_write(sheets, settings, registry)
        if code is not None:
            return code

    print("\nAll automated checks passed.")
    if not args.write_test:
        print("Run with --write-test to confirm sheet writes end-to-end.")
    print("Then run python bot.py and complete the Discord checklist.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
