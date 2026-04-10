import os
import re
import time
from pathlib import Path

import requests
from PIL import Image

TMP_DIR = Path("/tmp/ag_slide_mcp")


def ensure_tmp_dir() -> Path:
    TMP_DIR.mkdir(parents=True, exist_ok=True)
    return TMP_DIR


def hex_to_rgb(hex_color: str) -> dict:
    """Convert '#FF5733' to {'red': 1.0, 'green': 0.34, 'blue': 0.20}."""
    hex_color = hex_color.lstrip("#")
    if not re.match(r"^[0-9a-fA-F]{6}$", hex_color):
        raise ValueError(f"Invalid hex color: #{hex_color}. Expected format: #RRGGBB")
    r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    return {"red": r / 255.0, "green": g / 255.0, "blue": b / 255.0}


def pt_to_emu(points: float) -> int:
    """Convert points to EMUs (English Metric Units). 1 pt = 12700 EMU."""
    return int(points * 12700)


def inches_to_emu(inches: float) -> int:
    """Convert inches to EMUs. 1 inch = 914400 EMU."""
    return int(inches * 914400)


def resize_image(filepath: str, max_dimension: int = 2000) -> str:
    """Resize an image if either dimension exceeds max_dimension. Returns the filepath."""
    with Image.open(filepath) as img:
        if img.width > max_dimension or img.height > max_dimension:
            scale = min(max_dimension / img.width, max_dimension / img.height)
            new_size = (int(img.width * scale), int(img.height * scale))
            img = img.resize(new_size, Image.LANCZOS)
            img.save(filepath)
    return filepath


def download_url(url: str, filepath: str) -> str:
    """Download a URL and save to filepath."""
    resp = requests.get(url, timeout=30)
    resp.raise_for_status()
    with open(filepath, "wb") as f:
        f.write(resp.content)
    return filepath


def generate_filename(prefix: str, presentation_id: str, ext: str, slide_index: int | None = None) -> str:
    """Generate a timestamped filename in the tmp directory."""
    ensure_tmp_dir()
    ts = time.strftime("%Y%m%d_%H%M%S")
    parts = [prefix, presentation_id[:12]]
    if slide_index is not None:
        parts.append(str(slide_index))
    parts.append(ts)
    name = "_".join(parts) + f".{ext}"
    return str(TMP_DIR / name)
