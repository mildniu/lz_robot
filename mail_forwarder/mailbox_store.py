import json
from pathlib import Path
from typing import Any

MAILBOX_CONFIG_FILE = Path("settings/mailbox_aliases.json")


def _normalize_mailbox(raw_item: dict[str, Any]) -> dict[str, Any] | None:
    alias = str(raw_item.get("alias", "")).strip()
    host = str(raw_item.get("host", "")).strip()
    username = str(raw_item.get("username", "")).strip()
    password = str(raw_item.get("password", ""))
    mailbox = str(raw_item.get("mailbox", "INBOX")).strip() or "INBOX"

    try:
        port = int(raw_item.get("port"))
    except (TypeError, ValueError):
        port = 0

    if not alias or not host or port <= 0 or not username or not password:
        return None

    return {
        "alias": alias,
        "host": host,
        "port": port,
        "username": username,
        "password": password,
        "mailbox": mailbox,
    }


def load_mailbox_configs() -> dict[str, Any]:
    if not MAILBOX_CONFIG_FILE.exists():
        return {"mailboxes": []}

    try:
        raw = json.loads(MAILBOX_CONFIG_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {"mailboxes": []}

    mailboxes = []
    for item in raw.get("mailboxes", []):
        if not isinstance(item, dict):
            continue
        normalized = _normalize_mailbox(item)
        if normalized:
            mailboxes.append(normalized)
    return {"mailboxes": mailboxes}


def save_mailbox_configs(mailboxes: list[dict[str, Any]]) -> None:
    normalized_mailboxes = []
    for item in mailboxes:
        if not isinstance(item, dict):
            continue
        normalized = _normalize_mailbox(item)
        if normalized:
            normalized_mailboxes.append(normalized)

    MAILBOX_CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
    temp_file = MAILBOX_CONFIG_FILE.with_suffix(MAILBOX_CONFIG_FILE.suffix + ".tmp")
    temp_file.write_text(
        json.dumps({"mailboxes": normalized_mailboxes}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    temp_file.replace(MAILBOX_CONFIG_FILE)


def load_mailbox_alias_map() -> dict[str, dict[str, Any]]:
    alias_map: dict[str, dict[str, Any]] = {}
    for item in load_mailbox_configs().get("mailboxes", []):
        alias = str(item.get("alias", "")).strip()
        if alias:
            alias_map[alias] = item
    return alias_map
