from __future__ import annotations

import json
from pathlib import Path

ALIAS_CONFIG_FILE = Path("settings/webhook_aliases.json")


def load_webhook_aliases() -> dict:
    if not ALIAS_CONFIG_FILE.exists():
        return {"aliases": {}, "email_alias": ""}
    try:
        data = json.loads(ALIAS_CONFIG_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {"aliases": {}, "email_alias": ""}

    aliases = data.get("aliases", {})
    if not isinstance(aliases, dict):
        aliases = {}

    email_alias = data.get("email_alias", "")
    if not isinstance(email_alias, str):
        email_alias = ""

    return {"aliases": aliases, "email_alias": email_alias}


def save_webhook_aliases(aliases: dict[str, str], email_alias: str) -> None:
    ALIAS_CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
    tmp_file = ALIAS_CONFIG_FILE.with_suffix(ALIAS_CONFIG_FILE.suffix + ".tmp")
    tmp_file.write_text(
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
    tmp_file.replace(ALIAS_CONFIG_FILE)


def resolve_webhook_url(alias: str, aliases: dict[str, str]) -> str:
    if not alias:
        return ""
    return aliases.get(alias, "").strip()
