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
    *,
    service_account_path: Path | None = None,
    service_account_info: dict | None = None,
) -> gspread.Client:
    if service_account_info is not None:
        return gspread.service_account_from_dict(
            service_account_info,
            scopes=SCOPES,
        )
    if service_account_path is not None:
        if not service_account_path.is_file():
            raise FileNotFoundError(
                f"Service account file not found: {service_account_path}. "
                "Create a service account in Google Cloud, download JSON, "
                "and share your spreadsheet with the service account email."
            )
        return gspread.service_account(
            filename=str(service_account_path),
            scopes=SCOPES,
        )

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
