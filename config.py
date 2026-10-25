import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

PROJECT_ROOT = Path(__file__).resolve().parent
CONFIG_DIR = Path.home() / ".config" / "daisys-dough"


@dataclass(frozen=True)
class Settings:
    discord_token: str
    discord_channel_id: int
    google_sheet_id: str
    google_oauth_client_path: Path
    google_token_path: Path
    google_service_account_path: Path | None
    timezone: str
    react_on_error: bool
    fuzzy_match_cutoff: float
    sheets_min_interval_sec: float


def _require(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise ValueError(f"Missing required environment variable: {name}")
    return value


def default_oauth_client_path() -> Path:
    return CONFIG_DIR / "oauth_client.json"


def default_token_path() -> Path:
    return CONFIG_DIR / "token.json"


def resolve_oauth_client_path(configured: Path) -> Path:
    candidates: list[Path] = []
    if configured.is_absolute():
        candidates.append(configured)
    else:
        candidates.append(PROJECT_ROOT / configured)

    candidates.append(default_oauth_client_path())
    candidates.extend(sorted(CONFIG_DIR.glob("client_secret*.json")))
    candidates.append(PROJECT_ROOT / "oauth_client.json")
    candidates.extend(sorted(PROJECT_ROOT.glob("client_secret*.json")))

    seen: set[Path] = set()
    for path in candidates:
        resolved = path.resolve()
        if resolved in seen:
            continue
        seen.add(resolved)
        if resolved.is_file():
            return resolved

    raise FileNotFoundError(
        "OAuth client JSON not found. Run setup-credentials.ps1 or copy Desktop OAuth JSON to "
        f"{default_oauth_client_path()}"
    )


def resolve_token_path(configured: Path) -> Path:
    if configured.is_absolute():
        return configured
    if configured.name == "token.json" and configured.parent == Path("."):
        return default_token_path()
    return PROJECT_ROOT / configured


def resolve_service_account_path() -> Path | None:
    raw = os.getenv("GOOGLE_SERVICE_ACCOUNT_PATH", "").strip()
    if not raw:
        default = CONFIG_DIR / "service_account.json"
        if default.is_file():
            return default
        return None
    path = Path(raw)
    if path.is_absolute():
        return path
    return (PROJECT_ROOT / path).resolve()


def load_settings() -> Settings:
    oauth_path = resolve_oauth_client_path(
        Path(os.getenv("GOOGLE_OAUTH_CLIENT_PATH", str(default_oauth_client_path())))
    )
    return Settings(
        discord_token=_require("DISCORD_TOKEN"),
        discord_channel_id=int(_require("DISCORD_CHANNEL_ID")),
        google_sheet_id=_require("GOOGLE_SHEET_ID"),
        google_oauth_client_path=oauth_path,
        google_token_path=resolve_token_path(
            Path(os.getenv("GOOGLE_TOKEN_PATH", str(default_token_path())))
        ),
        google_service_account_path=resolve_service_account_path(),
        timezone=os.getenv("TIMEZONE", "America/Los_Angeles"),
        react_on_error=os.getenv("REACT_ON_ERROR", "true").lower() == "true",
        fuzzy_match_cutoff=float(os.getenv("FUZZY_MATCH_CUTOFF", "0.75")),
        sheets_min_interval_sec=float(os.getenv("SHEETS_MIN_INTERVAL_SEC", "3.0")),
    )
