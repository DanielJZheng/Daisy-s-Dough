import unittest
from pathlib import Path
from unittest.mock import patch

from google_auth import SCOPES, authorize_google


class GoogleAuthTests(unittest.TestCase):
    def test_authorize_requires_oauth_client_file(self) -> None:
        with self.assertRaises(FileNotFoundError):
            authorize_google(Path("missing-oauth-client.json"), Path("token.json"))

    @patch("google_auth.gspread.oauth")
    def test_authorize_uses_gspread_oauth(self, mock_oauth) -> None:
        client_path = Path("oauth_client.json")
        token_path = Path("token.json")
        client_path.write_text("{}", encoding="utf-8")
        try:
            authorize_google(client_path, token_path)
        finally:
            client_path.unlink(missing_ok=True)

        mock_oauth.assert_called_once_with(
            scopes=SCOPES,
            credentials_filename=str(client_path),
            authorized_user_filename=str(token_path),
        )


if __name__ == "__main__":
    unittest.main()
