from __future__ import annotations

import time
from collections.abc import Callable
from typing import TypeVar

from gspread.exceptions import APIError

T = TypeVar("T")


def with_sheets_retry(action: Callable[[], T], *, attempts: int = 8) -> T:
    delay = 2.0
    last_error: APIError | None = None
    for attempt in range(attempts):
        try:
            return action()
        except APIError as error:
            last_error = error
            if "429" not in str(error) or attempt == attempts - 1:
                raise
            time.sleep(delay)
            delay = min(delay * 2, 30.0)
    if last_error:
        raise last_error
    raise RuntimeError("Sheets retry failed without error")
