from googleapiclient.errors import HttpError

from ag_slide_mcp.config import get_template_id, set_template_id, load_config, _config_path
from ag_slide_mcp.google_clients import get_drive_service, get_slides_service
from ag_slide_mcp.server import server


@server.tool()
def set_template(template_id: str) -> dict:
    """Set the default Google Slides template ID.

    This template is used every time create_presentation is called.
    The template ID is the long string in the Google Slides URL:
    https://docs.google.com/presentation/d/{TEMPLATE_ID}/edit

    Args:
        template_id: The Google Slides presentation ID to use as the default template.
    """
    set_template_id(template_id)
    return {
        "success": True,
        "template_id": template_id,
        "config_path": str(_config_path()),
    }


@server.tool()
def get_config() -> dict:
    """Get the current AG Slide MCP configuration, including the default template ID."""
    config = load_config()
    template_id = config.get("default_template_id")
    return {
        "default_template_id": template_id,
        "config_path": str(_config_path()),
        "template_configured": template_id is not None,
    }


@server.tool()
def create_presentation(title: str, template_id: str | None = None) -> dict:
    """Create a new Google Slides presentation from the configured template.

    Uses the default template from config. You can override with template_id.
    If no template is configured, returns an error — call set_template first.

    Returns the presentation ID, title, and slide count.
    """
    try:
        # Resolve template: explicit arg > config > error
        effective_template = template_id or get_template_id()
        if not effective_template:
            return {
                "error": "No template configured. Call set_template(template_id) first with your Google Slides template ID. "
                         "The template ID is the long string in the URL: "
                         "https://docs.google.com/presentation/d/{TEMPLATE_ID}/edit",
            }

        drive = get_drive_service()
        copied = drive.files().copy(
            fileId=effective_template,
            body={"name": title},
        ).execute()
        presentation_id = copied["id"]

        slides_svc = get_slides_service()
        pres = slides_svc.presentations().get(presentationId=presentation_id).execute()

        return {
            "presentation_id": pres["presentationId"],
            "title": pres.get("title", title),
            "slides_count": len(pres.get("slides", [])),
            "slide_ids": [s["objectId"] for s in pres.get("slides", [])],
            "template_id": effective_template,
        }
    except HttpError as e:
        return {"error": f"Google API error: {e.reason}", "status": e.resp.status}


@server.tool()
def get_presentation(presentation_id: str) -> dict:
    """Get metadata for a Google Slides presentation.

    Returns title, slide count, slide IDs, and page dimensions.
    """
    try:
        slides_svc = get_slides_service()
        pres = slides_svc.presentations().get(presentationId=presentation_id).execute()

        page_size = pres.get("pageSize", {})
        width = page_size.get("width", {})
        height = page_size.get("height", {})

        return {
            "presentation_id": pres["presentationId"],
            "title": pres.get("title", ""),
            "slides_count": len(pres.get("slides", [])),
            "slide_ids": [s["objectId"] for s in pres.get("slides", [])],
            "width_pt": width.get("magnitude", 0) / 12700 if width.get("unit") == "EMU" else width.get("magnitude", 0),
            "height_pt": height.get("magnitude", 0) / 12700 if height.get("unit") == "EMU" else height.get("magnitude", 0),
        }
    except HttpError as e:
        return {"error": f"Google API error: {e.reason}", "status": e.resp.status}


@server.tool()
def list_presentations(max_results: int = 20, query: str | None = None) -> dict:
    """List Google Slides presentations in the user's Drive.

    Optional query filter (e.g., "Q4 Report") searches by name.
    """
    try:
        drive = get_drive_service()
        q = "mimeType='application/vnd.google-apps.presentation'"
        if query:
            q += f" and name contains '{query}'"

        result = drive.files().list(
            q=q,
            pageSize=max_results,
            fields="files(id, name, modifiedTime, createdTime)",
            orderBy="modifiedTime desc",
        ).execute()

        files = result.get("files", [])
        return {
            "count": len(files),
            "presentations": [
                {
                    "id": f["id"],
                    "name": f["name"],
                    "modified": f.get("modifiedTime", ""),
                    "created": f.get("createdTime", ""),
                }
                for f in files
            ],
        }
    except HttpError as e:
        return {"error": f"Google API error: {e.reason}", "status": e.resp.status}


@server.tool()
def delete_presentation(presentation_id: str) -> dict:
    """Move a Google Slides presentation to trash."""
    try:
        drive = get_drive_service()
        drive.files().update(
            fileId=presentation_id,
            body={"trashed": True},
        ).execute()
        return {"success": True, "presentation_id": presentation_id}
    except HttpError as e:
        return {"error": f"Google API error: {e.reason}", "status": e.resp.status}


@server.tool()
def rename_presentation(presentation_id: str, new_title: str) -> dict:
    """Rename a Google Slides presentation."""
    try:
        drive = get_drive_service()
        drive.files().update(
            fileId=presentation_id,
            body={"name": new_title},
        ).execute()
        return {"success": True, "presentation_id": presentation_id, "new_title": new_title}
    except HttpError as e:
        return {"error": f"Google API error: {e.reason}", "status": e.resp.status}
