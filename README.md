# AG Slide MCP

A Model Context Protocol (MCP) server for Google Slides. Create, populate, and visually verify presentations from Claude Code or any MCP client.

The key differentiator is a **visual feedback loop**: the server exports slides as PNG images that Claude can inspect, enabling iterative refinement until slides look polished.

## Tools (26)

| Module | Tools | Purpose |
|--------|-------|---------|
| **Presentation** | `set_template`, `get_config`, `create_presentation`, `get_presentation`, `list_presentations`, `delete_presentation`, `rename_presentation` | Config + CRUD operations |
| **Slides** | `add_slide`, `duplicate_slide`, `delete_slide`, `reorder_slides`, `get_slide` | Slide management with overflow detection |
| **Content** | `add_text_box`, `update_text`, `replace_all_text`, `add_image`, `add_shape`, `update_shape_style`, `resize_element`, `add_table`, `update_table_cell` | Content manipulation |
| **Export** | `export_slide_as_image`, `export_presentation_as_pdf` | Visual feedback loop |
| **Template** | `list_placeholders`, `fill_template`, `fill_template_from_copy` | Template-based workflow |

## Setup

### 1. Google Cloud credentials

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a project and enable **Google Slides API** + **Google Drive API**
3. Go to APIs & Services > Credentials
4. Create an **OAuth 2.0 Client ID** (Desktop app)
5. Download the JSON and save it as `~/.ag_slide_mcp/credentials.json`
6. Add your email as a test user in the OAuth consent screen

### 2. Install

```bash
git clone git@github.com:GaetanJuvin/ag-slide-mcp.git
cd ag-slide-mcp
python3 -m venv .venv
source .venv/bin/activate
pip install -e .
```

### 3. Register with Claude Code

The `.mcp.json` is included. Update the Python path to match your setup:

```json
{
  "mcpServers": {
    "ag-slide-mcp": {
      "command": "/path/to/ag-slide-mcp/.venv/bin/python",
      "args": ["-m", "ag_slide_mcp"],
      "env": {
        "AG_SLIDE_MCP_CONFIG_DIR": "~/.ag_slide_mcp"
      }
    }
  }
}
```

### 4. Configure your template

A default template is **required**. On first use, set it via the MCP tool:

```
set_template("1UQbx4i3OOnm...")  →  saves to ~/.ag_slide_mcp/config.json
```

The template ID is the long string in the Google Slides URL:
`https://docs.google.com/presentation/d/{TEMPLATE_ID}/edit`

Every `create_presentation` call copies this template. You can override per-call with `template_id`.

### 5. First run

On first use, a browser window opens for OAuth consent. Tokens are cached in `~/.ag_slide_mcp/token.json` for subsequent runs.

## How it works

### Template workflow

```
set_template("1UQbx4i3OOnm...")  →  saves default template to config
create_presentation(title)  →  copies template via Drive API
fill_template(presentation_id, {"{{title}}": "Q4 Report"})  →  replaces placeholders
export_slide_as_image(presentation_id, slide_index=0)  →  PNG to /tmp/ag_slide_mcp/
```

### Visual feedback loop

```
1. Create or update slide content
2. export_slide_as_image  →  saves PNG
3. Claude reads the PNG  →  evaluates layout and text
4. If issues found  →  update_text or resize_element to fix
5. Re-export and verify  →  repeat until clean
```

### Quality awareness

The server includes built-in quality guidance:

- **Overflow detection**: `get_slide` reports `overflow_risk` on text elements (none/low/medium/high)
- **Server instructions**: Text length guidelines per slide type are sent to the AI automatically
- **`resize_element`**: Fix layout issues by resizing text boxes programmatically

## Architecture

```
src/ag_slide_mcp/
├── server.py           # FastMCP server + quality instructions
├── config.py           # Template + config management (~/.ag_slide_mcp/config.json)
├── auth.py             # OAuth 2.0 flow with token caching
├── google_clients.py   # Lazy-initialized Slides + Drive API singletons
├── utils.py            # hex_to_rgb, pt_to_emu, resize_image helpers
└── tools/
    ├── presentation.py # create, get, list, delete, rename
    ├── slides.py       # add, duplicate, delete, reorder, get (with overflow detection)
    ├── content.py      # text boxes, shapes, images, tables, resize
    ├── export.py       # slide→PNG, presentation→PDF
    └── template.py     # placeholders, fill, copy+fill
```

## Coordinate system

All positions are in **points** (1 inch = 72 points). Standard widescreen slides are 720 × 405 pt. Colors are hex strings (`#FF5733`). Conversions to EMUs happen internally.
