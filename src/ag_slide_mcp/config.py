import json
import os
from pathlib import Path


def _config_dir() -> Path:
    path = Path(os.environ.get("AG_SLIDE_MCP_CONFIG_DIR", "~/.ag_slide_mcp")).expanduser()
    path.mkdir(parents=True, exist_ok=True)
    return path


def _config_path() -> Path:
    return _config_dir() / "config.json"


def load_config() -> dict:
    """Load the config file. Returns empty dict if not found."""
    path = _config_path()
    if path.exists():
        with open(path) as f:
            return json.load(f)
    return {}


def save_config(config: dict) -> None:
    """Save config to disk."""
    path = _config_path()
    with open(path, "w") as f:
        json.dump(config, f, indent=2)


def get_template_id() -> str | None:
    """Get the configured default template ID."""
    return load_config().get("default_template_id")


def set_template_id(template_id: str) -> None:
    """Set the default template ID."""
    config = load_config()
    config["default_template_id"] = template_id
    save_config(config)
