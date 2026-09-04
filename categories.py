from __future__ import annotations

import difflib
from dataclasses import dataclass
from datetime import date

import gspread

from sheets_retry import with_sheets_retry

CATEGORIES_SHEET = "Categories"
CATEGORY_HEADERS = ["Category", "Type", "FirstSeen", "Aliases"]
CATEGORY_DATA_RANGE = "A:D"


@dataclass(frozen=True)
class CategoryMatch:
    kind: str
    canonical: str | None = None
    suggested: str | None = None


class CategoryRegistry:
    def __init__(self, spreadsheet: gspread.Spreadsheet, fuzzy_cutoff: float) -> None:
        self._spreadsheet = spreadsheet
        self._fuzzy_cutoff = fuzzy_cutoff
        self._worksheet = self._ensure_sheet()
        self._cached_rows: list[list[str]] | None = None

    def _ensure_sheet(self) -> gspread.Worksheet:
        try:
            worksheet = self._spreadsheet.worksheet(CATEGORIES_SHEET)
        except gspread.WorksheetNotFound:
            worksheet = self._spreadsheet.add_worksheet(
                title=CATEGORIES_SHEET, rows=200, cols=10
            )
            with_sheets_retry(
                lambda: worksheet.update("A1:D1", [CATEGORY_HEADERS])
            )
        else:
            headers = with_sheets_retry(lambda: worksheet.row_values(1))
            if headers != CATEGORY_HEADERS:
                with_sheets_retry(
                    lambda: worksheet.update("A1:D1", [CATEGORY_HEADERS])
                )
        return worksheet

    def _invalidate_cache(self) -> None:
        self._cached_rows = None

    def _rows(self) -> list[list[str]]:
        if self._cached_rows is not None:
            return self._cached_rows

        values = with_sheets_retry(
            lambda: self._worksheet.get_all_values(CATEGORY_DATA_RANGE)
        )
        if len(values) <= 1:
            self._cached_rows = []
        else:
            self._cached_rows = values[1:]
        return self._cached_rows

    def load_categories(self, entry_type: str) -> list[str]:
        categories: list[str] = []
        for row in self._rows():
            if len(row) < 2:
                continue
            name, row_type = row[0].strip(), row[1].strip().lower()
            if name and row_type == entry_type:
                categories.append(name)
        return categories

    def _alias_map(self, entry_type: str) -> dict[str, str]:
        mapping: dict[str, str] = {}
        for row in self._rows():
            if len(row) < 2:
                continue
            name, row_type = row[0].strip(), row[1].strip().lower()
            if not name or row_type != entry_type:
                continue
            mapping[name.lower()] = name
            if len(row) >= 4 and row[3].strip():
                for alias in row[3].split(","):
                    alias = alias.strip()
                    if alias:
                        mapping[alias.lower()] = name
        return mapping

    def resolve(self, raw_category: str, entry_type: str) -> CategoryMatch:
        raw = raw_category.strip()
        if not raw:
            return CategoryMatch(kind="none")

        alias_map = self._alias_map(entry_type)
        direct = alias_map.get(raw.lower())
        if direct:
            return CategoryMatch(kind="exact", canonical=direct)

        known = list({name for name in alias_map.values()})
        if not known:
            return CategoryMatch(kind="none")

        suggestions = difflib.get_close_matches(
            raw, known, n=1, cutoff=self._fuzzy_cutoff
        )
        if suggestions:
            return CategoryMatch(kind="fuzzy", suggested=suggestions[0])

        return CategoryMatch(kind="none")

    def add_category(
        self,
        name: str,
        entry_type: str,
        *,
        aliases: str = "",
        first_seen: date | None = None,
    ) -> str:
        canonical = name.strip()
        existing = {c.lower(): c for c in self.load_categories(entry_type)}
        if canonical.lower() in existing:
            return existing[canonical.lower()]

        seen = (first_seen or date.today()).isoformat()
        with_sheets_retry(
            lambda: self._worksheet.append_row(
                [canonical, entry_type, seen, aliases]
            )
        )
        self._invalidate_cache()
        return canonical

    def add_alias(self, canonical: str, alias: str, entry_type: str) -> None:
        alias = alias.strip()
        if not alias:
            return

        rows = self._rows()
        for index, row in enumerate(rows, start=2):
            if len(row) < 2:
                continue
            name, row_type = row[0].strip(), row[1].strip().lower()
            if name == canonical and row_type == entry_type:
                aliases = row[3].strip() if len(row) >= 4 else ""
                parts = [part.strip() for part in aliases.split(",") if part.strip()]
                if alias not in parts:
                    parts.append(alias)
                    with_sheets_retry(
                        lambda: self._worksheet.update_cell(
                            index, 4, ", ".join(parts)
                        )
                    )
                    self._invalidate_cache()
                return
