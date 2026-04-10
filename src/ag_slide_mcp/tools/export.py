import time

from googleapiclient.errors import HttpError

from ag_slide_mcp.auth import get_credentials
from ag_slide_mcp.google_clients import get_drive_service, get_slides_service
from ag_slide_mcp.server import server
from ag_slide_mcp.utils import download_url, generate_filename, resize_image


@server.tool()
def export_slide_as_image(
    presentation_id: str,
    slide_index: int = 0,
    max_dimension: int = 2000,
) -> str:
    """Export a single slide as a PNG image for visual verification.

    Returns the file path to the saved PNG. Use Claude's Read tool to view it.

    Args:
        presentation_id: The Google Slides presentation ID.
        slide_index: Zero-based index of the slide to export (default: 0).
        max_dimension: Maximum pixel dimension for width or height (default: 2000).
    """
    try:
        slides_svc = get_slides_service()
        pres = slides_svc.presentations().get(presentationId=presentation_id).execute()
        slides_list = pres.get("slides", [])

        if slide_index < 0 or slide_index >= len(slides_list):
            return f"Error: slide_index {slide_index} out of range. Presentation has {len(slides_list)} slides (0-{len(slides_list) - 1})."

        page_id = slides_list[slide_index]["objectId"]

        # Get thumbnail URL
        thumbnail = slides_svc.presentations().pages().getThumbnail(
            presentationId=presentation_id,
            pageObjectId=page_id,
            thumbnailProperties_mimeType="PNG",
            thumbnailProperties_thumbnailSize="LARGE",
        ).execute()

        content_url = thumbnail["contentUrl"]
        filepath = generate_filename("slide", presentation_id, "png", slide_index)

        download_url(content_url, filepath)
        resize_image(filepath, max_dimension)

        return filepath

    except HttpError as e:
        if e.resp.status == 500:
            # Thumbnail may not be ready after recent edits — retry once
            time.sleep(2)
            try:
                thumbnail = slides_svc.presentations().pages().getThumbnail(
                    presentationId=presentation_id,
                    pageObjectId=page_id,
                    thumbnailProperties_mimeType="PNG",
                    thumbnailProperties_thumbnailSize="LARGE",
                ).execute()
                content_url = thumbnail["contentUrl"]
                filepath = generate_filename("slide", presentation_id, "png", slide_index)
                download_url(content_url, filepath)
                resize_image(filepath, max_dimension)
                return filepath
            except HttpError as retry_err:
                return f"Error exporting slide after retry: {retry_err.reason}"
        return f"Error exporting slide: {e.reason}"


@server.tool()
def export_presentation_as_pdf(presentation_id: str) -> str:
    """Export the entire presentation as a PDF file.

    Returns the file path to the saved PDF.
    """
    try:
        drive = get_drive_service()
        pdf_content = drive.files().export(
            fileId=presentation_id,
            mimeType="application/pdf",
        ).execute()

        filepath = generate_filename("presentation", presentation_id, "pdf")

        with open(filepath, "wb") as f:
            f.write(pdf_content)

        return filepath

    except HttpError as e:
        return f"Error exporting presentation as PDF: {e.reason}"
