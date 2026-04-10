import uuid

from googleapiclient.errors import HttpError

from ag_slide_mcp.google_clients import get_slides_service
from ag_slide_mcp.server import server
from ag_slide_mcp.utils import hex_to_rgb, pt_to_emu


def _make_id() -> str:
    return f"obj_{uuid.uuid4().hex[:12]}"


@server.tool()
def add_text_box(
    presentation_id: str,
    slide_id: str,
    text: str,
    x: float,
    y: float,
    width: float,
    height: float,
    font_size: int = 18,
    bold: bool = False,
    color: str | None = None,
) -> dict:
    """Add a text box to a slide.

    Coordinates are in points (1 inch = 72 points).
    Color is a hex string like '#FF5733'.

    Args:
        presentation_id: The presentation ID.
        slide_id: The slide object ID to add the text box to.
        text: The text content.
        x: X position in points from left edge.
        y: Y position in points from top edge.
        width: Width in points.
        height: Height in points.
        font_size: Font size in points (default: 18).
        bold: Whether to bold the text (default: False).
        color: Optional hex color for the text (e.g., '#333333').
    """
    try:
        slides_svc = get_slides_service()
        element_id = _make_id()

        requests_list = [
            {
                "createShape": {
                    "objectId": element_id,
                    "shapeType": "TEXT_BOX",
                    "elementProperties": {
                        "pageObjectId": slide_id,
                        "size": {
                            "width": {"magnitude": pt_to_emu(width), "unit": "EMU"},
                            "height": {"magnitude": pt_to_emu(height), "unit": "EMU"},
                        },
                        "transform": {
                            "scaleX": 1,
                            "scaleY": 1,
                            "translateX": pt_to_emu(x),
                            "translateY": pt_to_emu(y),
                            "unit": "EMU",
                        },
                    },
                }
            },
            {
                "insertText": {
                    "objectId": element_id,
                    "text": text,
                    "insertionIndex": 0,
                }
            },
        ]

        # Style the text
        style = {"fontSize": {"magnitude": font_size, "unit": "PT"}}
        fields = "fontSize"

        if bold:
            style["bold"] = True
            fields += ",bold"

        if color:
            rgb = hex_to_rgb(color)
            style["foregroundColor"] = {"opaqueColor": {"rgbColor": rgb}}
            fields += ",foregroundColor"

        requests_list.append({
            "updateTextStyle": {
                "objectId": element_id,
                "style": style,
                "textRange": {"type": "ALL"},
                "fields": fields,
            }
        })

        slides_svc.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": requests_list},
        ).execute()

        return {"success": True, "element_id": element_id, "slide_id": slide_id}
    except HttpError as e:
        return {"error": f"Google API error: {e.reason}", "status": e.resp.status}


@server.tool()
def update_text(presentation_id: str, shape_id: str, text: str) -> dict:
    """Replace all text in an existing text box or shape.

    Args:
        presentation_id: The presentation ID.
        shape_id: The element/shape object ID.
        text: The new text content.
    """
    try:
        slides_svc = get_slides_service()

        requests_list = [
            {
                "deleteText": {
                    "objectId": shape_id,
                    "textRange": {"type": "ALL"},
                }
            },
            {
                "insertText": {
                    "objectId": shape_id,
                    "text": text,
                    "insertionIndex": 0,
                }
            },
        ]

        slides_svc.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": requests_list},
        ).execute()

        return {"success": True, "shape_id": shape_id}
    except HttpError as e:
        return {"error": f"Google API error: {e.reason}", "status": e.resp.status}


@server.tool()
def replace_all_text(presentation_id: str, find_text: str, replace_text: str) -> dict:
    """Global find-and-replace text across all slides in the presentation.

    Args:
        presentation_id: The presentation ID.
        find_text: The text to search for.
        replace_text: The replacement text.
    """
    try:
        slides_svc = get_slides_service()

        response = slides_svc.presentations().batchUpdate(
            presentationId=presentation_id,
            body={
                "requests": [{
                    "replaceAllText": {
                        "containsText": {"text": find_text, "matchCase": True},
                        "replaceText": replace_text,
                    }
                }]
            },
        ).execute()

        occurrences = 0
        for reply in response.get("replies", []):
            occurrences += reply.get("replaceAllText", {}).get("occurrencesChanged", 0)

        return {"success": True, "occurrences_changed": occurrences}
    except HttpError as e:
        return {"error": f"Google API error: {e.reason}", "status": e.resp.status}


@server.tool()
def add_image(
    presentation_id: str,
    slide_id: str,
    image_url: str,
    x: float,
    y: float,
    width: float,
    height: float,
) -> dict:
    """Insert an image from a URL onto a slide.

    The image URL must be publicly accessible. Coordinates are in points.

    Args:
        presentation_id: The presentation ID.
        slide_id: The slide object ID.
        image_url: Public URL of the image.
        x: X position in points.
        y: Y position in points.
        width: Width in points.
        height: Height in points.
    """
    try:
        slides_svc = get_slides_service()
        element_id = _make_id()

        slides_svc.presentations().batchUpdate(
            presentationId=presentation_id,
            body={
                "requests": [{
                    "createImage": {
                        "objectId": element_id,
                        "url": image_url,
                        "elementProperties": {
                            "pageObjectId": slide_id,
                            "size": {
                                "width": {"magnitude": pt_to_emu(width), "unit": "EMU"},
                                "height": {"magnitude": pt_to_emu(height), "unit": "EMU"},
                            },
                            "transform": {
                                "scaleX": 1,
                                "scaleY": 1,
                                "translateX": pt_to_emu(x),
                                "translateY": pt_to_emu(y),
                                "unit": "EMU",
                            },
                        },
                    }
                }]
            },
        ).execute()

        return {"success": True, "element_id": element_id, "slide_id": slide_id}
    except HttpError as e:
        return {"error": f"Google API error: {e.reason}", "status": e.resp.status}


@server.tool()
def add_shape(
    presentation_id: str,
    slide_id: str,
    shape_type: str,
    x: float,
    y: float,
    width: float,
    height: float,
    fill_color: str | None = None,
) -> dict:
    """Add a shape to a slide.

    Coordinates are in points. shape_type can be: RECTANGLE, ELLIPSE, ROUND_RECTANGLE,
    TRIANGLE, DIAMOND, PENTAGON, HEXAGON, STAR_4, STAR_5, ARROW_EAST, etc.

    Args:
        presentation_id: The presentation ID.
        slide_id: The slide object ID.
        shape_type: The shape type (e.g., 'RECTANGLE', 'ELLIPSE').
        x: X position in points.
        y: Y position in points.
        width: Width in points.
        height: Height in points.
        fill_color: Optional hex fill color (e.g., '#4285F4').
    """
    try:
        slides_svc = get_slides_service()
        element_id = _make_id()

        requests_list = [
            {
                "createShape": {
                    "objectId": element_id,
                    "shapeType": shape_type,
                    "elementProperties": {
                        "pageObjectId": slide_id,
                        "size": {
                            "width": {"magnitude": pt_to_emu(width), "unit": "EMU"},
                            "height": {"magnitude": pt_to_emu(height), "unit": "EMU"},
                        },
                        "transform": {
                            "scaleX": 1,
                            "scaleY": 1,
                            "translateX": pt_to_emu(x),
                            "translateY": pt_to_emu(y),
                            "unit": "EMU",
                        },
                    },
                }
            }
        ]

        if fill_color:
            rgb = hex_to_rgb(fill_color)
            requests_list.append({
                "updateShapeProperties": {
                    "objectId": element_id,
                    "shapeProperties": {
                        "shapeBackgroundFill": {
                            "solidFill": {"color": {"rgbColor": rgb}},
                        }
                    },
                    "fields": "shapeBackgroundFill.solidFill.color",
                }
            })

        slides_svc.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": requests_list},
        ).execute()

        return {"success": True, "element_id": element_id, "slide_id": slide_id}
    except HttpError as e:
        return {"error": f"Google API error: {e.reason}", "status": e.resp.status}


@server.tool()
def update_shape_style(
    presentation_id: str,
    shape_id: str,
    fill_color: str | None = None,
    border_color: str | None = None,
    border_weight: float | None = None,
) -> dict:
    """Update visual properties of an existing shape.

    Args:
        presentation_id: The presentation ID.
        shape_id: The shape object ID.
        fill_color: Optional hex fill color (e.g., '#4285F4').
        border_color: Optional hex border color.
        border_weight: Optional border weight in points.
    """
    try:
        slides_svc = get_slides_service()

        props = {}
        fields = []

        if fill_color:
            rgb = hex_to_rgb(fill_color)
            props["shapeBackgroundFill"] = {
                "solidFill": {"color": {"rgbColor": rgb}},
            }
            fields.append("shapeBackgroundFill.solidFill.color")

        if border_color:
            rgb = hex_to_rgb(border_color)
            outline = props.setdefault("outline", {})
            outline["outlineFill"] = {
                "solidFill": {"color": {"rgbColor": rgb}},
            }
            fields.append("outline.outlineFill.solidFill.color")

        if border_weight is not None:
            outline = props.setdefault("outline", {})
            outline["weight"] = {"magnitude": pt_to_emu(border_weight), "unit": "EMU"}
            fields.append("outline.weight")

        if not fields:
            return {"error": "No style properties provided to update."}

        slides_svc.presentations().batchUpdate(
            presentationId=presentation_id,
            body={
                "requests": [{
                    "updateShapeProperties": {
                        "objectId": shape_id,
                        "shapeProperties": props,
                        "fields": ",".join(fields),
                    }
                }]
            },
        ).execute()

        return {"success": True, "shape_id": shape_id}
    except HttpError as e:
        return {"error": f"Google API error: {e.reason}", "status": e.resp.status}


@server.tool()
def add_table(
    presentation_id: str,
    slide_id: str,
    rows: int,
    cols: int,
    x: float,
    y: float,
    width: float,
    height: float,
) -> dict:
    """Insert a table on a slide.

    Coordinates are in points.

    Args:
        presentation_id: The presentation ID.
        slide_id: The slide object ID.
        rows: Number of rows.
        cols: Number of columns.
        x: X position in points.
        y: Y position in points.
        width: Width in points.
        height: Height in points.
    """
    try:
        slides_svc = get_slides_service()
        element_id = _make_id()

        slides_svc.presentations().batchUpdate(
            presentationId=presentation_id,
            body={
                "requests": [{
                    "createTable": {
                        "objectId": element_id,
                        "elementProperties": {
                            "pageObjectId": slide_id,
                            "size": {
                                "width": {"magnitude": pt_to_emu(width), "unit": "EMU"},
                                "height": {"magnitude": pt_to_emu(height), "unit": "EMU"},
                            },
                            "transform": {
                                "scaleX": 1,
                                "scaleY": 1,
                                "translateX": pt_to_emu(x),
                                "translateY": pt_to_emu(y),
                                "unit": "EMU",
                            },
                        },
                        "rows": rows,
                        "columns": cols,
                    }
                }]
            },
        ).execute()

        return {"success": True, "element_id": element_id, "slide_id": slide_id, "rows": rows, "cols": cols}
    except HttpError as e:
        return {"error": f"Google API error: {e.reason}", "status": e.resp.status}


@server.tool()
def resize_element(
    presentation_id: str,
    element_id: str,
    width: float | None = None,
    height: float | None = None,
    x: float | None = None,
    y: float | None = None,
) -> dict:
    """Resize and/or reposition an existing element (text box, shape, image).

    Use this to fix text overflow — enlarge a text box so content fits.
    Coordinates are in points (1 inch = 72 points).

    Only the provided parameters are updated; others remain unchanged.

    Args:
        presentation_id: The presentation ID.
        element_id: The element object ID.
        width: New width in points (optional).
        height: New height in points (optional).
        x: New X position in points (optional).
        y: New Y position in points (optional).
    """
    try:
        slides_svc = get_slides_service()

        # Fetch current element to get existing transform and size
        pres = slides_svc.presentations().get(presentationId=presentation_id).execute()
        current_element = None
        for slide in pres.get("slides", []):
            for el in slide.get("pageElements", []):
                if el.get("objectId") == element_id:
                    current_element = el
                    break
            if current_element:
                break

        if not current_element:
            return {"error": f"Element {element_id} not found in presentation."}

        current_transform = current_element.get("transform", {})
        current_size = current_element.get("size", {})

        new_transform = {
            "scaleX": current_transform.get("scaleX", 1),
            "scaleY": current_transform.get("scaleY", 1),
            "translateX": current_transform.get("translateX", 0),
            "translateY": current_transform.get("translateY", 0),
            "shearX": current_transform.get("shearX", 0),
            "shearY": current_transform.get("shearY", 0),
            "unit": "EMU",
        }

        new_size = {
            "width": current_size.get("width", {"magnitude": 0, "unit": "EMU"}),
            "height": current_size.get("height", {"magnitude": 0, "unit": "EMU"}),
        }

        if width is not None:
            new_size["width"] = {"magnitude": pt_to_emu(width), "unit": "EMU"}
        if height is not None:
            new_size["height"] = {"magnitude": pt_to_emu(height), "unit": "EMU"}
        if x is not None:
            new_transform["translateX"] = pt_to_emu(x)
        if y is not None:
            new_transform["translateY"] = pt_to_emu(y)

        slides_svc.presentations().batchUpdate(
            presentationId=presentation_id,
            body={
                "requests": [{
                    "updatePageElementTransform": {
                        "objectId": element_id,
                        "applyMode": "ABSOLUTE",
                        "transform": new_transform,
                    }
                }]
            },
        ).execute()

        # Size update requires a separate approach — update via element properties
        # For size, we need to use updatePageElementTransform with scale factors
        if width is not None or height is not None:
            # Calculate new scale factors based on desired size vs original size
            orig_w = current_size.get("width", {}).get("magnitude", 1)
            orig_h = current_size.get("height", {}).get("magnitude", 1)

            if width is not None and orig_w > 0:
                new_transform["scaleX"] = pt_to_emu(width) / orig_w
            if height is not None and orig_h > 0:
                new_transform["scaleY"] = pt_to_emu(height) / orig_h

            slides_svc.presentations().batchUpdate(
                presentationId=presentation_id,
                body={
                    "requests": [{
                        "updatePageElementTransform": {
                            "objectId": element_id,
                            "applyMode": "ABSOLUTE",
                            "transform": new_transform,
                        }
                    }]
                },
            ).execute()

        return {"success": True, "element_id": element_id}
    except HttpError as e:
        return {"error": f"Google API error: {e.reason}", "status": e.resp.status}


@server.tool()
def update_table_cell(
    presentation_id: str,
    table_id: str,
    row: int,
    col: int,
    text: str,
) -> dict:
    """Set the text content of a specific table cell.

    Args:
        presentation_id: The presentation ID.
        table_id: The table element object ID.
        row: Zero-based row index.
        col: Zero-based column index.
        text: The text to set in the cell.
    """
    try:
        slides_svc = get_slides_service()

        slides_svc.presentations().batchUpdate(
            presentationId=presentation_id,
            body={
                "requests": [
                    {
                        "deleteText": {
                            "objectId": table_id,
                            "cellLocation": {"rowIndex": row, "columnIndex": col},
                            "textRange": {"type": "ALL"},
                        }
                    },
                    {
                        "insertText": {
                            "objectId": table_id,
                            "cellLocation": {"rowIndex": row, "columnIndex": col},
                            "text": text,
                            "insertionIndex": 0,
                        }
                    },
                ]
            },
        ).execute()

        return {"success": True, "table_id": table_id, "row": row, "col": col}
    except HttpError as e:
        return {"error": f"Google API error: {e.reason}", "status": e.resp.status}
