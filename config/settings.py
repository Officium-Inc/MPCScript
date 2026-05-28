import json
import os
from pathlib import Path

APPDATA_DIR = Path(os.environ.get("APPDATA", Path.home())) / "MPCScript"
SETTINGS_FILE = APPDATA_DIR / "settings.json"
ALIASES_FILE  = APPDATA_DIR / "aliases.json"

DEFAULTS: dict = {
    "prepared_by_name": "Gilfred C. Sale",
    "prepared_by_title": "Sports Supervisor",
    "checked_by_name": "Jazmin Montealegre",
    "checked_by_title": "S&A Executive Secretary",
    "validated_by_name": "Maricel Balingbing",
    "validated_by_title": "Billing Assistant",
    "commission_rate": 5.0,
}


def load() -> dict:
    APPDATA_DIR.mkdir(parents=True, exist_ok=True)
    if SETTINGS_FILE.exists():
        try:
            with open(SETTINGS_FILE, encoding="utf-8") as f:
                saved = json.load(f)
            return {**DEFAULTS, **saved}
        except (json.JSONDecodeError, OSError):
            pass
    return DEFAULTS.copy()


def save(settings: dict) -> None:
    APPDATA_DIR.mkdir(parents=True, exist_ok=True)
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=2)


def load_aliases() -> dict:
    """Load saved {pos_name_lower: isp_display_name_or_None} alias mappings."""
    APPDATA_DIR.mkdir(parents=True, exist_ok=True)
    if ALIASES_FILE.exists():
        try:
            with open(ALIASES_FILE, encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            pass
    return {}


def save_aliases(aliases: dict) -> None:
    """Persist alias mappings to disk."""
    APPDATA_DIR.mkdir(parents=True, exist_ok=True)
    with open(ALIASES_FILE, "w", encoding="utf-8") as f:
        json.dump(aliases, f, indent=2)
