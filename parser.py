import re
from dataclasses import dataclass

DESCRIPTION_PATTERN = r"[A-Za-z][A-Za-z0-9_'-]*"
ENTRY_PATTERN = re.compile(
    rf"^(?P<sign>[+-])?\$?(?P<amount>\d+(?:\.\d{{1,2}})?)\s+"
    rf"(?P<description>{DESCRIPTION_PATTERN})(?:\s+(?P<note>.+))?$"
)

SUDO_PATTERN = re.compile(
    r"^sudo\s+(?:<@!?(\d+)>|@?([A-Za-z0-9_.'-]+))\s+(.+)$",
    re.IGNORECASE | re.DOTALL,
)
SUDO_HEADER_PATTERN = re.compile(
    r"^sudo\s+(?:<@!?(\d+)>|@?([A-Za-z0-9_.'-]+))\s*$",
    re.IGNORECASE,
)
AMOUNT_TOKEN = r"(?P<sign>[+-])?\$?(?P<amount>\d{1,3}(?:,\d{3})*|\d+)(?:\.\d{1,2})?"
AMOUNT_FIRST_PATTERN = re.compile(rf"^{AMOUNT_TOKEN}\s+(?P<rest>.+)$")
AMOUNT_LAST_PATTERN = re.compile(rf"^(?P<rest>.+?)\s+{AMOUNT_TOKEN}$")
ONLY_AMOUNT_PATTERN = re.compile(rf"^{AMOUNT_TOKEN}$")
SECTION_TYPES = {"income": "income", "fixed": "expense", "others": "expense"}
SECTION_ADD_PATTERN = re.compile(r"^section\s+add\s+(.+)$", re.IGNORECASE)
REFRESH_PATTERN = re.compile(
    r"^(?:refresh|sync|update)(?:\s+(?:sheet|ledger))?$",
    re.IGNORECASE,
)
RECOVER_PATTERN = re.compile(
    r"^(?:recover|restore)(?:\s+(?:sheet|ledger|transactions?))?$",
    re.IGNORECASE,
)


def parse_refresh_request(content: str) -> bool:
    text = _normalize_message(content)
    return bool(REFRESH_PATTERN.match(text))


def parse_recover_request(content: str) -> bool:
    text = _normalize_message(content)
    return bool(RECOVER_PATTERN.match(text))


def parse_section_add(content: str) -> str | None:
    text = _normalize_message(content)
    match = SECTION_ADD_PATTERN.match(text)
    if not match:
        return None
    name = match.group(1).strip()
    return name or None

FORMAT_HELP = (
    "Format:\n"
    "• `105 seven` or `167 burger oat` — expense. First word is category, rest is note\n"
    "• `43 shopping milk` — shopping table\n"
    "• `+1200 paycheck bonus` — income table\n"
    "• `Week 1 2400` — multi-word category, amount at the end\n"
    "• Several lines in one message. `sudo @name` on the first line logs the rest for them\n"
    "Categories used once land in one-off. Repeats all move to Recurring. Same name, same spelling.\n"
    "Optional `$` and commas. Skip section titles like Income, Fixed, Others.\n"
    "• `section add Vacation` — new column group on your ledger tab\n"
    "• `refresh` — backfill timestamps, totals, category summary on your tab\n"
    "• `recover` — restore Recurring/One-off from hidden RAW backup"
)


def _normalize_message(content: str) -> str:
    return content.strip().replace("\u2019", "'").replace("\u2018", "'")


def looks_like_entry(content: str) -> bool:
    text = _normalize_message(content)
    if not text:
        return False
    return (
        text[0].isdigit()
        or text[0] in "$+-"
        or text.lower().startswith("sudo ")
    )


@dataclass(frozen=True)
class SudoTarget:
    user_id: int | None
    name: str | None
    body: str


def parse_sudo(content: str) -> SudoTarget | None:
    text = _normalize_message(content)
    match = SUDO_PATTERN.match(text)
    if not match:
        return None
    user_id = int(match.group(1)) if match.group(1) else None
    return SudoTarget(user_id, match.group(2), match.group(3).strip())


@dataclass(frozen=True)
class ParsedEntry:
    entry_type: str
    amount: float
    description: str
    note: str | None


def _amount_value(raw: str) -> float | None:
    try:
        amount = float(raw.replace(",", ""))
    except ValueError:
        return None
    if amount <= 0:
        return None
    return amount


def _entry_from_match(match: re.Match[str], *, entry_type: str | None = None) -> ParsedEntry | None:
    amount = _amount_value(match.group("amount"))
    if amount is None:
        return None
    sign = match.group("sign")
    if entry_type is None:
        entry_type = "income" if sign == "+" else "expense"
    rest = match.group("rest").strip()
    if not rest:
        return None
    return ParsedEntry(entry_type, amount, rest, None)


def parse_line(content: str, *, entry_type: str | None = None) -> ParsedEntry | None:
    text = _normalize_message(content).replace("\t", " ")
    text = re.sub(r"\s+", " ", text)
    if not text or ONLY_AMOUNT_PATTERN.match(text):
        return None
    if text.lower().startswith(("balance ", "bank ")):
        return None
    if text.lower().startswith("oneoff "):
        text = text.split(" ", 1)[1].strip()
        entry_type = entry_type or "expense"
    if text.lower() in SECTION_TYPES:
        return None

    strict = ENTRY_PATTERN.match(text)
    if strict and entry_type is None:
        amount = _amount_value(strict.group("amount"))
        if amount is None:
            return None
        sign = strict.group("sign")
        kind = "income" if sign == "+" else "expense"
        note = strict.group("note")
        return ParsedEntry(
            kind,
            amount,
            strict.group("description"),
            note.strip() if note else None,
        )

    first = AMOUNT_FIRST_PATTERN.match(text)
    if first:
        entry = _entry_from_match(first, entry_type=entry_type)
        if entry is not None and " " not in entry.description:
            return entry
        if entry is not None and entry_type is not None:
            return entry

    last = AMOUNT_LAST_PATTERN.match(text)
    if last:
        return _entry_from_match(last, entry_type=entry_type)
    if first:
        return _entry_from_match(first, entry_type=entry_type)
    return None


def parse_entries(content: str) -> tuple[SudoTarget | None, list[ParsedEntry]]:
    text = _normalize_message(content)
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    sudo: SudoTarget | None = None
    if lines:
        header = SUDO_HEADER_PATTERN.match(lines[0].replace("\t", " ").strip())
        whole = parse_sudo(lines[0]) if lines[0].lower().startswith("sudo") else None
        if header:
            user_id = int(header.group(1)) if header.group(1) else None
            sudo = SudoTarget(user_id, header.group(2), "")
            lines = lines[1:]
        elif whole is not None:
            user_id = whole.user_id
            sudo = SudoTarget(user_id, whole.name, "")
            lines = [whole.body, *lines[1:]]

    entries: list[ParsedEntry] = []
    section: str | None = None
    for line in lines:
        key = re.sub(r"\s+", " ", line.replace("\t", " ")).strip().lower()
        if key in SECTION_TYPES:
            section = SECTION_TYPES[key]
            continue
        entry = parse_line(line, entry_type=section)
        if entry is not None:
            entries.append(entry)
    return sudo, entries


def parse_message(content: str) -> ParsedEntry | None:
    text = _normalize_message(content)
    if not text or "\n" in text:
        return None
    _sudo, entries = parse_entries(text)
    if len(entries) != 1:
        return None
    return entries[0]
