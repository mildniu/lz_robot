import json
from pathlib import Path
from typing import Dict

ALIAS_CONFIG_FILE = Path("settings/webhook_aliases.json")


def load_webhook_aliases() -> dict:
    if not ALIAS_CONFIG_FILE.exists():
        return {"aliases": {}, "email_alias": ""}
    try:
        data = json.loads(ALIAS_CONFIG_FILE.read_text(encoding="utf-8"))
        aliases = data.get("aliases", {})
        if not isinstance(aliases, dict):
            aliases = {}
        email_alias = data.get("email_alias", "")
        if not isinstance(email_alias, str):
            email_alias = ""
        return {"aliases": aliases, "email_alias": email_alias}
    except Exception:
        return {"aliases": {}, "email_alias": ""}


def save_webhook_aliases(aliases: Dict[str, str], email_alias: str) -> None:
    ALIAS_CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
    temp_file = ALIAS_CONFIG_FILE.with_suffix(ALIAS_CONFIG_FILE.suffix + ".tmp")
    temp_file.write_text(
        json.dumps(
            {
                "aliases": aliases,
                "email_alias": email_alias,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    temp_file.replace(ALIAS_CONFIG_FILE)


def resolve_webhook_url(alias: str, aliases: Dict[str, str]) -> str:
    if not alias:
        return ""
    return aliases.get(alias, "").strip()
