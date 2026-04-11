from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

from ag_slide_mcp.config import _config_dir

SCOPES = [
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/drive",
]


def get_credentials() -> Credentials:
    """Get valid Google OAuth 2.0 credentials.

    On first run, opens a browser for consent. Subsequent runs use the cached token.
    """
    config = _config_dir()
    token_path = config / "token.json"
    creds_path = config / "credentials.json"

    creds = None

    # Try loading cached token
    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)

    # Refresh or run consent flow
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    elif not creds or not creds.valid:
        if not creds_path.exists():
            raise FileNotFoundError(
                f"Google OAuth credentials not found at {creds_path}. "
                "To set up:\n"
                "1. Go to https://console.cloud.google.com/\n"
                "2. Create a project and enable Google Slides API + Google Drive API\n"
                "3. Go to APIs & Services > Credentials\n"
                "4. Create an OAuth 2.0 Client ID (Desktop app)\n"
                "5. Download the JSON and save it as:\n"
                f"   {creds_path}\n"
                "6. Add your email as a test user in the OAuth consent screen"
            )
        flow = InstalledAppFlow.from_client_secrets_file(str(creds_path), SCOPES)
        creds = flow.run_local_server(port=0)

    # Cache the token
    with open(token_path, "w") as f:
        f.write(creds.to_json())

    return creds
