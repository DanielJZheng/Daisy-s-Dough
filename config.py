import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()


@dataclass(frozen=True)
class Settings:
    discord_token: str
    discord_channel_id: int
    google_sheet_id: str
    google_oauth_client_path: Path
    google_token_path: Path
    timezone: str
    react_on_error: bool
    fuzzy_match_cutoff: float


def _require(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise ValueError(f"Missing required environment variable: {name}")
    return value


def load_settings() -> Settings:
    return Settings(
        discord_token=_require("DISCORD_TOKEN"),
        discord_channel_id=int(_require("DISCORD_CHANNEL_ID")),
        google_sheet_id=_require("GOOGLE_SHEET_ID"),
        google_oauth_client_path=Path(
            os.getenv("GOOGLE_OAUTH_CLIENT_PATH", "oauth_client.json")
        ),
        google_token_path=Path(os.getenv("GOOGLE_TOKEN_PATH", "token.json")),
        timezone=os.getenv("TIMEZONE", "America/Los_Angeles"),
        react_on_error=os.getenv("REACT_ON_ERROR", "true").lower() == "true",
        fuzzy_match_cutoff=float(os.getenv("FUZZY_MATCH_CUTOFF", "0.75")),
    )
