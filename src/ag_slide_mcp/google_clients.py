from googleapiclient.discovery import build

from ag_slide_mcp.auth import get_credentials

_slides_service = None
_drive_service = None


def get_slides_service():
    """Get a lazily-initialized Google Slides API service."""
    global _slides_service
    if _slides_service is None:
        creds = get_credentials()
        _slides_service = build("slides", "v1", credentials=creds)
    return _slides_service


def get_drive_service():
    """Get a lazily-initialized Google Drive API service."""
    global _drive_service
    if _drive_service is None:
        creds = get_credentials()
        _drive_service = build("drive", "v3", credentials=creds)
    return _drive_service
