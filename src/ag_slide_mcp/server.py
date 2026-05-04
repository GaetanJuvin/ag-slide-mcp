from mcp.server.fastmcp import FastMCP

server = FastMCP(
    "ag-slide-mcp",
    instructions="""You are creating Google Slides presentations. Follow these quality rules:

## Slide Quality Checklist (ALWAYS verify after editing)

1. **Export and visually verify every slide** after making changes. Call `export_slide_as_image` and inspect the PNG. Never assume text fits — always check.

2. **Text overflow is the #1 quality issue.** After calling `get_slide`, check the `overflow_risk` field on text elements:
   - "high" or "medium" → text WILL be clipped. Shorten the text or use `resize_element` to enlarge the text box.
   - The template's text boxes have fixed sizes. Long paragraphs will be cut off at the bottom.

3. **Text length guidelines by slide type:**
   - Title slides: max ~30 chars for main title, ~60 chars for subtitle
   - Section dividers: max ~40 chars
   - Tagline/quote slides: max ~80 chars (2 short lines)
   - Body text slides: max ~300 chars
   - List items (picks layout): max ~100 chars per item (title + description)

4. **After every update_text call**, re-export the slide to verify the text fits visually. If it doesn't:
   - First try shortening the text (preferred — keeps design clean)
   - If meaning would be lost, use `resize_element` to make the text box larger

5. **Visual feedback loop:** Create → Fill → Export → Review → Fix → Re-export. Iterate until every slide looks clean.

6. **Coordinate system:** All positions are in points (1 inch = 72 points). Slide dimensions are typically 720×405 pt (widescreen 16:9).

7. **Respect template fonts.** The template defines theme fonts for titles and body text. Do not override `fontFamily` ad-hoc.
   - Prefer editing existing template text boxes (`update_text`, `replace_all_text`, `fill_template`) — they preserve the original font.
   - When you must create a new text box, call `add_text_box` with `font_role="title"` or `font_role="body"` so it picks up the template's theme font automatically. Use `get_theme_fonts` if you need to inspect them.
   - Only pass an explicit `font_family` when the user asks for a specific font.
""",
)

# Tool modules register themselves via @server.tool() on import
from ag_slide_mcp.tools import presentation  # noqa: E402, F401
from ag_slide_mcp.tools import slides  # noqa: E402, F401
from ag_slide_mcp.tools import content  # noqa: E402, F401
from ag_slide_mcp.tools import export  # noqa: E402, F401
from ag_slide_mcp.tools import template  # noqa: E402, F401
