import re

from googleapiclient.errors import HttpError

from ag_slide_mcp.config import get_template_id
from ag_slide_mcp.google_clients import get_drive_service, get_slides_service
from ag_slide_mcp.server import server

PLACEHOLDER_RE = re.compile(r"\{\{[^}]+\}\}")


def _extract_text_from_element(element: dict) -> str:
    """Extract plain text from a page element."""
    shape = element.get("shape", {})
    text_content = shape.get("text", {})
    text_elements = text_content.get("textElements", [])
    parts = []
    for te in text_elements:
        text_run = te.get("textRun", {})
        if "content" in text_run:
            parts.append(text_run["content"])
    return "".join(parts)


@server.tool()
def list_placeholders(presentation_id: str) -> dict:
    """Scan all slides for placeholder text patterns like {{title}}, {{date}}, etc.

    Returns a mapping of slide index to list of placeholders found on that slide.
    """
    try:
        slides_svc = get_slides_service()
        pres = slides_svc.presentations().get(presentationId=presentation_id).execute()

        result = {}
        all_placeholders = set()

        for i, slide in enumerate(pres.get("slides", [])):
            slide_placeholders = set()
            for element in slide.get("pageElements", []):
                text = _extract_text_from_element(element)
                matches = PLACEHOLDER_RE.findall(text)
                slide_placeholders.update(matches)
            if slide_placeholders:
                result[f"slide_{i}"] = sorted(slide_placeholders)
                all_placeholders.update(slide_placeholders)

        return {
            "presentation_id": presentation_id,
            "total_placeholders": len(all_placeholders),
            "all_placeholders": sorted(all_placeholders),
            "by_slide": result,
        }
    except HttpError as e:
        return {"error": f"Google API error: {e.reason}", "status": e.resp.status}


@server.tool()
def fill_template(presentation_id: str, replacements: dict[str, str]) -> dict:
    """Replace placeholder text throughout a presentation.

    Takes a dict mapping placeholder strings to replacement values.
    Example: {"{{title}}": "Q4 Report", "{{date}}": "2026-04-09"}
    """
    try:
        slides_svc = get_slides_service()

        requests_list = []
        for find_text, replace_text in replacements.items():
            requests_list.append({
                "replaceAllText": {
                    "containsText": {
                        "text": find_text,
                        "matchCase": True,
                    },
                    "replaceText": replace_text,
                }
            })

        if not requests_list:
            return {"error": "No replacements provided."}

        response = slides_svc.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": requests_list},
        ).execute()

        # Count replacements made per key
        counts = {}
        for reply in response.get("replies", []):
            replace_result = reply.get("replaceAllText", {})
            occurrences = replace_result.get("occurrencesChanged", 0)
            # Match reply order to request order
            idx = response["replies"].index(reply)
            if idx < len(replacements):
                key = list(replacements.keys())[idx]
                counts[key] = occurrences

        return {
            "presentation_id": presentation_id,
            "replacements_made": counts,
        }
    except HttpError as e:
        return {"error": f"Google API error: {e.reason}", "status": e.resp.status}


@server.tool()
def fill_template_from_copy(
    title: str,
    replacements: dict[str, str],
    template_id: str | None = None,
) -> dict:
    """Copy the default template and fill in all placeholders in one step.

    Uses the configured default template. You can override with template_id.
    1. Copies the template via Google Drive API
    2. Replaces all {{placeholder}} text with provided values
    3. Returns the new presentation ID

    Args:
        title: Title for the new presentation.
        replacements: Dict mapping placeholder strings to values.
        template_id: Optional override template ID (uses default from config if omitted).
    """
    try:
        effective_template = template_id or get_template_id()
        if not effective_template:
            return {
                "error": "No template configured. Call set_template(template_id) first.",
            }

        # Step 1: Copy the template
        drive = get_drive_service()
        copied = drive.files().copy(
            fileId=effective_template,
            body={"name": title},
        ).execute()
        new_id = copied["id"]

        # Step 2: Fill placeholders
        if replacements:
            fill_result = fill_template(new_id, replacements)
            if "error" in fill_result:
                return {
                    "presentation_id": new_id,
                    "title": title,
                    "warning": f"Copied but fill failed: {fill_result['error']}",
                }
            replacements_made = fill_result.get("replacements_made", {})
        else:
            replacements_made = {}

        # Step 3: Get final state
        slides_svc = get_slides_service()
        pres = slides_svc.presentations().get(presentationId=new_id).execute()

        return {
            "presentation_id": new_id,
            "title": title,
            "slides_count": len(pres.get("slides", [])),
            "slide_ids": [s["objectId"] for s in pres.get("slides", [])],
            "replacements_made": replacements_made,
        }
    except HttpError as e:
        return {"error": f"Google API error: {e.reason}", "status": e.resp.status}
