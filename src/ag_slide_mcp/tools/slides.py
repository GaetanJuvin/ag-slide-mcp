import uuid

from googleapiclient.errors import HttpError

from ag_slide_mcp.google_clients import get_slides_service
from ag_slide_mcp.server import server


def _estimate_text_overflow(text: str, width_pt: float, height_pt: float, font_size: float = 14) -> dict:
    """Estimate whether text will overflow a text box.

    Uses a heuristic: ~1.6 chars per point of width at a given font size,
    and ~1.4× font size per line height. This is approximate but catches
    obvious overflow issues before they hit the rendered slide.
    """
    if not text or width_pt <= 0 or height_pt <= 0:
        return {"overflow_risk": "none"}

    chars_per_line = max(1, int(width_pt / (font_size * 0.6)))
    line_height_pt = font_size * 1.4
    max_lines = max(1, int(height_pt / line_height_pt))

    # Count actual lines (word wrap + explicit newlines)
    lines_needed = 0
    for paragraph in text.split("\n"):
        if not paragraph.strip():
            lines_needed += 1
        else:
            lines_needed += max(1, -(-len(paragraph) // chars_per_line))  # ceiling division

    if lines_needed <= max_lines:
        risk = "none"
    elif lines_needed <= max_lines * 1.2:
        risk = "low"
    elif lines_needed <= max_lines * 1.5:
        risk = "medium"
    else:
        risk = "high"

    return {
        "overflow_risk": risk,
        "lines_needed": lines_needed,
        "max_lines": max_lines,
        "chars_per_line": chars_per_line,
        "text_length": len(text),
    }


def _extract_element_info(element: dict) -> dict:
    """Extract useful info from a page element."""
    info = {
        "id": element.get("objectId", ""),
        "type": "unknown",
    }

    # Position and size from transform (handles scaleX/scaleY)
    transform = element.get("transform", {})
    size = element.get("size", {})
    scale_x = transform.get("scaleX", 1) if transform else 1
    scale_y = transform.get("scaleY", 1) if transform else 1

    if transform:
        info["x_pt"] = round(transform.get("translateX", 0) / 12700, 1)
        info["y_pt"] = round(transform.get("translateY", 0) / 12700, 1)

    width_pt = 0
    height_pt = 0
    if size:
        w = size.get("width", {})
        h = size.get("height", {})
        if w:
            width_pt = round(w.get("magnitude", 0) / 12700 * abs(scale_x), 1)
            info["width_pt"] = width_pt
        if h:
            height_pt = round(h.get("magnitude", 0) / 12700 * abs(scale_y), 1)
            info["height_pt"] = height_pt

    # Shape with text
    if "shape" in element:
        shape = element["shape"]
        info["type"] = shape.get("shapeType", "SHAPE")
        text_content = shape.get("text", {})
        text_elements = text_content.get("textElements", [])
        parts = []
        font_size = 14  # default estimate
        for te in text_elements:
            text_run = te.get("textRun", {})
            if "content" in text_run:
                parts.append(text_run["content"])
            # Try to extract font size from style
            style = text_run.get("style", {})
            fs = style.get("fontSize", {})
            if fs.get("magnitude"):
                font_size = fs["magnitude"]
        text = "".join(parts).strip()
        if text:
            info["text"] = text[:500]
            info["text_length"] = len(text)
            # Estimate overflow risk
            overflow = _estimate_text_overflow(text, width_pt, height_pt, font_size)
            if overflow["overflow_risk"] != "none":
                info["overflow_risk"] = overflow["overflow_risk"]
                info["lines_needed"] = overflow["lines_needed"]
                info["max_lines"] = overflow["max_lines"]

    # Image
    if "image" in element:
        info["type"] = "IMAGE"
        info["source_url"] = element["image"].get("sourceUrl", "")

    # Table
    if "table" in element:
        table = element["table"]
        info["type"] = "TABLE"
        info["rows"] = table.get("rows", 0)
        info["columns"] = table.get("columns", 0)

    return info


@server.tool()
def add_slide(
    presentation_id: str,
    layout: str = "BLANK",
    insertion_index: int | None = None,
) -> dict:
    """Add a new slide to the presentation.

    Layout options: BLANK, CAPTION_ONLY, TITLE, TITLE_AND_BODY, TITLE_AND_TWO_COLUMNS,
    TITLE_ONLY, ONE_COLUMN_TEXT, MAIN_POINT, SECTION_HEADER, BIG_NUMBER.

    Args:
        presentation_id: The presentation ID.
        layout: Predefined layout name (default: 'BLANK').
        insertion_index: Position to insert the slide (default: end).
    """
    try:
        slides_svc = get_slides_service()
        slide_id = f"slide_{uuid.uuid4().hex[:12]}"

        request = {
            "createSlide": {
                "objectId": slide_id,
                "slideLayoutReference": {"predefinedLayout": layout},
            }
        }
        if insertion_index is not None:
            request["createSlide"]["insertionIndex"] = insertion_index

        slides_svc.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": [request]},
        ).execute()

        return {"success": True, "slide_id": slide_id, "layout": layout}
    except HttpError as e:
        return {"error": f"Google API error: {e.reason}", "status": e.resp.status}


@server.tool()
def duplicate_slide(presentation_id: str, slide_id: str) -> dict:
    """Duplicate an existing slide.

    Args:
        presentation_id: The presentation ID.
        slide_id: The object ID of the slide to duplicate.
    """
    try:
        slides_svc = get_slides_service()

        response = slides_svc.presentations().batchUpdate(
            presentationId=presentation_id,
            body={
                "requests": [{
                    "duplicateObject": {
                        "objectId": slide_id,
                    }
                }]
            },
        ).execute()

        new_id = response["replies"][0]["duplicateObject"]["objectId"]
        return {"success": True, "original_slide_id": slide_id, "new_slide_id": new_id}
    except HttpError as e:
        return {"error": f"Google API error: {e.reason}", "status": e.resp.status}


@server.tool()
def delete_slide(presentation_id: str, slide_id: str) -> dict:
    """Delete a slide from the presentation.

    Args:
        presentation_id: The presentation ID.
        slide_id: The object ID of the slide to delete.
    """
    try:
        slides_svc = get_slides_service()

        slides_svc.presentations().batchUpdate(
            presentationId=presentation_id,
            body={
                "requests": [{
                    "deleteObject": {"objectId": slide_id}
                }]
            },
        ).execute()

        return {"success": True, "deleted_slide_id": slide_id}
    except HttpError as e:
        return {"error": f"Google API error: {e.reason}", "status": e.resp.status}


@server.tool()
def reorder_slides(
    presentation_id: str,
    slide_ids: list[str],
    insertion_index: int,
) -> dict:
    """Move slides to a new position in the presentation.

    Args:
        presentation_id: The presentation ID.
        slide_ids: List of slide object IDs to move.
        insertion_index: The zero-based index where the slides should be moved to.
    """
    try:
        slides_svc = get_slides_service()

        slides_svc.presentations().batchUpdate(
            presentationId=presentation_id,
            body={
                "requests": [{
                    "updateSlidesPosition": {
                        "slideObjectIds": slide_ids,
                        "insertionIndex": insertion_index,
                    }
                }]
            },
        ).execute()

        return {"success": True, "moved_slides": slide_ids, "new_index": insertion_index}
    except HttpError as e:
        return {"error": f"Google API error: {e.reason}", "status": e.resp.status}


@server.tool()
def get_slide(presentation_id: str, slide_index: int) -> dict:
    """Get all elements on a specific slide with their IDs, positions, and content.

    Args:
        presentation_id: The presentation ID.
        slide_index: Zero-based index of the slide.
    """
    try:
        slides_svc = get_slides_service()
        pres = slides_svc.presentations().get(presentationId=presentation_id).execute()
        slides_list = pres.get("slides", [])

        if slide_index < 0 or slide_index >= len(slides_list):
            return {"error": f"slide_index {slide_index} out of range. Presentation has {len(slides_list)} slides (0-{len(slides_list) - 1})."}

        slide = slides_list[slide_index]
        elements = [_extract_element_info(e) for e in slide.get("pageElements", [])]

        return {
            "slide_id": slide["objectId"],
            "slide_index": slide_index,
            "elements_count": len(elements),
            "elements": elements,
        }
    except HttpError as e:
        return {"error": f"Google API error: {e.reason}", "status": e.resp.status}
