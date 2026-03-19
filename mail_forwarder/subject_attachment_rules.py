import json
from pathlib import Path
from typing import Any

from .attachment_service import normalize_attachment_types

RULES_FILE = Path("settings/subject_attachment_rules.json")


def parse_types_input(raw_value: str) -> list[str]:
    value = str(raw_value).strip().replace(";", ",").replace("|", ",")
    if not value:
        return []
    first_item = value.split(",", 1)[0].strip()
    return normalize_attachment_types([first_item] if first_item else [])


def parse_filename_keywords_input(raw_value: str) -> list[str]:
    value = str(raw_value).strip().replace(";", ",").replace("|", ",")
    if not value:
        return []
    first_item = value.split(",", 1)[0].strip()
    return [first_item] if first_item else []


def normalize_trigger_mode(raw_value: Any) -> str:
    value = str(raw_value).strip().lower()
    return value if value in {"periodic", "timed"} else "periodic"


def normalize_schedule_time(raw_value: Any) -> str:
    value = str(raw_value).strip()
    if len(value) != 5 or value[2] != ":":
        return ""
    hour_text, minute_text = value.split(":", 1)
    if not (hour_text.isdigit() and minute_text.isdigit()):
        return ""
    hour = int(hour_text)
    minute = int(minute_text)
    if not (0 <= hour <= 23 and 0 <= minute <= 59):
        return ""
    return f"{hour:02d}:{minute:02d}"


def _normalize_rule(raw_rule: dict[str, Any]) -> dict[str, Any] | None:
    keyword = str(raw_rule.get("keyword", "")).strip()
    types = normalize_attachment_types(raw_rule.get("types", []))
    filename_keywords = parse_filename_keywords_input(",".join(raw_rule.get("filename_keywords", [])))
    enabled = bool(raw_rule.get("enabled", True))
    webhook_alias = str(raw_rule.get("webhook_alias", "")).strip()
    webhook_url = str(raw_rule.get("webhook_url", "")).strip()
    mailbox_alias = str(raw_rule.get("mailbox_alias", "")).strip()
    script_path = str(raw_rule.get("script_path", "")).strip()
    script_output_dir = str(raw_rule.get("script_output_dir", "")).strip()
    trigger_mode = normalize_trigger_mode(raw_rule.get("trigger_mode", "periodic"))
    schedule_time = normalize_schedule_time(raw_rule.get("schedule_time", ""))
    poll_interval_seconds = raw_rule.get("poll_interval_seconds")
    max_attachment_size_mb = raw_rule.get("max_attachment_size_mb")
    try:
        poll_interval_seconds = int(poll_interval_seconds)
    except (TypeError, ValueError):
        poll_interval_seconds = None
    try:
        max_attachment_size_mb = int(max_attachment_size_mb)
    except (TypeError, ValueError):
        max_attachment_size_mb = None
    if poll_interval_seconds is not None and poll_interval_seconds <= 0:
        poll_interval_seconds = None
    if max_attachment_size_mb is not None and max_attachment_size_mb <= 0:
        max_attachment_size_mb = None
    if not keyword or not types:
        return None
    return {
        "enabled": enabled,
        "keyword": keyword,
        "types": types,
        "filename_keywords": filename_keywords,
        "webhook_alias": webhook_alias,
        "webhook_url": webhook_url,
        "mailbox_alias": mailbox_alias,
        "script_path": script_path,
        "script_output_dir": script_output_dir,
        "trigger_mode": trigger_mode,
        "schedule_time": schedule_time,
        "poll_interval_seconds": poll_interval_seconds,
        "max_attachment_size_mb": max_attachment_size_mb,
    }


def load_subject_attachment_rules() -> dict[str, Any]:
    if not RULES_FILE.exists():
        return {"rules": []}

    try:
        raw = json.loads(RULES_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {"rules": []}

    rules = []
    for item in raw.get("rules", []):
        if not isinstance(item, dict):
            rules.append({})
            continue
        normalized = _normalize_rule(item)
        rules.append(normalized or {})
    return {"rules": rules}


def save_subject_attachment_rules(rules: list[dict[str, Any]]) -> None:
    normalized_rules: list[dict[str, Any]] = []
    for item in rules:
        if not isinstance(item, dict):
            normalized_rules.append({})
            continue
        normalized = _normalize_rule(item)
        normalized_rules.append(normalized or {})

    RULES_FILE.parent.mkdir(parents=True, exist_ok=True)
    tmp_file = RULES_FILE.with_suffix(RULES_FILE.suffix + ".tmp")
    tmp_file.write_text(
        json.dumps({"rules": normalized_rules}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    tmp_file.replace(RULES_FILE)


def resolve_enabled_rule_for_subject(subject: str) -> tuple[dict[str, Any] | None, str]:
    payload = load_subject_attachment_rules()
    for item in payload.get("rules", []):
        keyword = item.get("keyword", "")
        if not item.get("enabled", True):
            continue
        if keyword and keyword in subject:
            return item, keyword
    return None, ""


def list_enabled_rule_keywords() -> list[str]:
    payload = load_subject_attachment_rules()
    values = []
    for item in payload.get("rules", []):
        if not item.get("enabled", True):
            continue
        keyword = str(item.get("keyword", "")).strip()
        if keyword:
            values.append(keyword)
    # Keep order while deduplicating.
    seen = set()
    deduped = []
    for item in values:
        if item in seen:
            continue
        seen.add(item)
        deduped.append(item)
    return deduped


def list_enabled_rules() -> list[dict[str, Any]]:
    payload = load_subject_attachment_rules()
    rules: list[dict[str, Any]] = []
    for item in payload.get("rules", []):
        if not item.get("enabled", True):
            continue
        if not item.get("keyword"):
            continue
        rules.append(item)
    return rules


def list_enabled_rules_with_slots() -> list[tuple[int, dict[str, Any]]]:
    payload = load_subject_attachment_rules()
    rules: list[tuple[int, dict[str, Any]]] = []
    for index, item in enumerate(payload.get("rules", []), start=1):
        if not item.get("enabled", True):
            continue
        if not item.get("keyword"):
            continue
        rules.append((index, item))
    return rules


def resolve_attachment_filters_for_subject(subject: str) -> tuple[list[str], list[str], str]:
    rule, matched_keyword = resolve_enabled_rule_for_subject(subject)
    if not rule:
        return [], [], ""
    return rule.get("types", []), rule.get("filename_keywords", []), matched_keyword
