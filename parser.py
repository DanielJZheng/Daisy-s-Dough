import re
from dataclasses import dataclass

EXPENSE_PATTERN = re.compile(
    r"^(\d+(?:\.\d{1,2})?)\s+([A-Za-z][A-Za-z0-9_-]*)(?:\s+(.+))?$"
)
INCOME_PATTERN = re.compile(
    r"^income\s+(\d+(?:\.\d{1,2})?)\s+([A-Za-z][A-Za-z0-9_-]*)(?:\s+(.+))?$",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class ParsedEntry:
    entry_type: str
    amount: float
    category: str
    note: str | None


def parse_message(content: str) -> ParsedEntry | None:
    text = content.strip()
    if not text:
        return None

    income_match = INCOME_PATTERN.match(text)
    if income_match:
        amount = float(income_match.group(1))
        category = income_match.group(2)
        note = income_match.group(3)
        return ParsedEntry("income", amount, category, note.strip() if note else None)

    expense_match = EXPENSE_PATTERN.match(text)
    if expense_match:
        amount = float(expense_match.group(1))
        category = expense_match.group(2)
        note = expense_match.group(3)
        return ParsedEntry("expense", amount, category, note.strip() if note else None)

    return None
