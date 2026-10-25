from __future__ import annotations

from google.auth.exceptions import RefreshError


def format_google_sheets_error(error: Exception) -> str:
    if "429" in str(error):
        return (
            "Google Sheets rate limit — could not save after several retries. "
            "Wait ~1 minute and send the same line again (duplicates are skipped)."
        )
    if isinstance(error, RefreshError) or "invalid_grant" in str(error):
        return (
            "Google Sheets sign-in expired or revoked.\n"
            "• OAuth: delete `%USERPROFILE%\\.config\\daisys-dough\\token.json`, "
            "run `python verify.py`, sign in again.\n"
            "• Or use a service account (set GOOGLE_SERVICE_ACCOUNT_PATH in .env) — "
            "no browser login, does not expire like user tokens."
        )
    return str(error)
