from __future__ import annotations

from pathlib import Path

import gspread

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def authorize_google(
    oauth_client_path: Path,
    token_path: Path,
) -> gspread.Client:
    if not oauth_client_path.is_file():
        raise FileNotFoundError(
            f"OAuth client file not found: {oauth_client_path}. "
            "Download a Desktop OAuth client JSON from Google Cloud Console."
        )

    return gspread.oauth(
        scopes=SCOPES,
        credentials_filename=str(oauth_client_path),
        authorized_user_filename=str(token_path),
    )
