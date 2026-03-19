import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urlparse

SETTINGS_DIR = Path("settings")
APP_CONFIG_FILE = SETTINGS_DIR / "app_config.json"


def parse_upload_url(send_url: str) -> str:
    parsed = urlparse(send_url)
    key = parse_qs(parsed.query).get("key", [""])[0]
    if not key:
        raise ValueError("WEBHOOK_SEND_URL must include query param: key")
    return f"{parsed.scheme}://{parsed.netloc}/im-external/v1/webhook/upload-attachment?key={key}&type=2"


def normalize_path_value(raw_value: str) -> Path:
    value = str(raw_value).strip()
    if os.name != "nt":
        value = value.replace("\\", "/")
    return Path(value).expanduser()


def _default_payload() -> dict[str, Any]:
    return {
        "IMAP_HOST": "imap.chinatelecom.cn",
        "IMAP_PORT": "993",
        "EMAIL_USERNAME": "",
        "EMAIL_PASSWORD": "",
        "IMAP_MAILBOX": "INBOX",
        "SUBJECT_KEYWORDS": "",
        "POLL_INTERVAL_SECONDS": "600",
        "MAX_ATTACHMENT_SIZE_MB": "30",
        "WEBHOOK_SEND_URL": "",
        "WEBHOOK_SEND_ALIAS": "",
        "DOWNLOAD_DIR": "downloads",
        "STATE_FILE": "state/mail_state.json",
        "WINDOW_WIDTH": "960",
        "WINDOW_HEIGHT": "820",
        "SIDEBAR_WIDTH": "220",
        "UI_APPEARANCE": "light",
        "UI_COLOR_THEME": "blue",
        "START_PAGE": "execute",
        "UI_LOG_POLL_MS": "100",
        "AUTO_SCROLL_LOG": "true",
        "UI_SCALE": "1.0",
        "SCRIPT_TIMEOUT_SECONDS": "300",
        "APP_TITLE": "量子推送机器人 v5.1",
        "APP_FOOTER_TEXT": "v5.1\nby 不丢西瓜der",
    }


def _load_app_payload() -> dict[str, Any]:
    if not APP_CONFIG_FILE.exists():
        return _default_payload()
    try:
        raw = json.loads(APP_CONFIG_FILE.read_text(encoding="utf-8"))
        if not isinstance(raw, dict):
            return _default_payload()
    except Exception:
        return _default_payload()

    payload = _default_payload()
    for key in payload.keys():
        value = raw.get(key)
        if value is None:
            continue
        payload[key] = str(value)
    return payload


def _save_app_payload(payload: dict[str, Any]) -> None:
    SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
    tmp_file = APP_CONFIG_FILE.with_suffix(APP_CONFIG_FILE.suffix + ".tmp")
    tmp_file.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp_file.replace(APP_CONFIG_FILE)


def upsert_env_file(_dotenv_path: Path, updates: dict[str, str]) -> None:
    """Backwards-compatible API: persist app config into settings/app_config.json."""
    payload = _load_app_payload()
    for key, value in updates.items():
        payload[key] = str(value)
    _save_app_payload(payload)


def validate_config_values(config_values: dict[str, str]) -> list[str]:
    errors: list[str] = []

    required_fields = {
        "IMAP_HOST": "IMAP 服务器地址",
        "IMAP_PORT": "IMAP 端口",
        "EMAIL_USERNAME": "邮箱账号",
        "EMAIL_PASSWORD": "邮箱密码/授权码",
        "IMAP_MAILBOX": "邮箱文件夹",
    }
    for key, label in required_fields.items():
        if not config_values.get(key, "").strip():
            errors.append(f"{label}不能为空")

    int_fields = {
        "IMAP_PORT": "IMAP 端口",
        "POLL_INTERVAL_SECONDS": "轮询间隔",
        "MAX_ATTACHMENT_SIZE_MB": "最大附件大小",
    }
    for key, label in int_fields.items():
        value = config_values.get(key, "").strip()
        if not value:
            continue
        try:
            parsed = int(value)
            if parsed <= 0:
                errors.append(f"{label}必须大于 0")
        except ValueError:
            errors.append(f"{label}必须是整数")

    webhook_send_url = config_values.get("WEBHOOK_SEND_URL", "").strip()
    if webhook_send_url and not webhook_send_url.startswith(("http://", "https://")):
        errors.append("Webhook URL 必须以 http:// 或 https:// 开头")
    return errors


@dataclass(frozen=True)
class AppConfig:
    imap_host: str
    imap_port: int
    email_username: str
    email_password: str
    imap_mailbox: str
    subject_keywords: list[str]
    poll_interval_seconds: int
    download_dir: Path
    state_file: Path
    webhook_send_url: str
    webhook_upload_url: str
    max_attachment_size_mb: int = 30
    window_width: int = 960
    window_height: int = 820
    sidebar_width: int = 220
    ui_appearance: str = "light"
    ui_color_theme: str = "blue"
    start_page: str = "execute"
    ui_log_poll_ms: int = 100
    auto_scroll_log: bool = True
    ui_scale: float = 1.0
    script_timeout_seconds: int = 300
    app_title: str = "量子推送机器人 v5.1"
    app_footer_text: str = "v5.1\nby 不丢西瓜der"

    @property
    def max_attachment_size_bytes(self) -> int:
        return self.max_attachment_size_mb * 1024 * 1024


def load_config() -> AppConfig:
    def parse_positive_int(value: Any, default: int) -> int:
        try:
            parsed = int(str(value).strip())
            return parsed if parsed > 0 else default
        except (TypeError, ValueError):
            return default

    def parse_positive_float(value: Any, default: float) -> float:
        try:
            parsed = float(str(value).strip())
            return parsed if parsed > 0 else default
        except (TypeError, ValueError):
            return default

    def parse_bool(value: Any, default: bool) -> bool:
        if isinstance(value, bool):
            return value
        text = str(value).strip().lower()
        if text in ("1", "true", "yes", "on"):
            return True
        if text in ("0", "false", "no", "off"):
            return False
        return default

    def parse_choice(value: Any, allowed: set[str], default: str) -> str:
        text = str(value).strip()
        return text if text in allowed else default

    payload = _load_app_payload()
    webhook_send_url = str(payload.get("WEBHOOK_SEND_URL", "")).strip()
    webhook_upload_url = ""
    if webhook_send_url:
        try:
            webhook_upload_url = parse_upload_url(webhook_send_url)
        except Exception:
            webhook_upload_url = ""

    subject_keywords_str = str(payload.get("SUBJECT_KEYWORDS", "")).strip()
    subject_keywords = [kw.strip() for kw in subject_keywords_str.split(",") if kw.strip()]
    ui_appearance = parse_choice(payload.get("UI_APPEARANCE", "light"), {"dark", "light", "system"}, "light")
    start_page = parse_choice(
        payload.get("START_PAGE", "execute"),
        {"execute", "folder", "bot_test", "settings", "about"},
        "execute",
    )

    return AppConfig(
        imap_host=str(payload.get("IMAP_HOST", "imap.chinatelecom.cn")),
        imap_port=parse_positive_int(payload.get("IMAP_PORT", "993"), 993),
        email_username=str(payload.get("EMAIL_USERNAME", "")),
        email_password=str(payload.get("EMAIL_PASSWORD", "")),
        imap_mailbox=str(payload.get("IMAP_MAILBOX", "INBOX")),
        subject_keywords=subject_keywords,
        poll_interval_seconds=parse_positive_int(payload.get("POLL_INTERVAL_SECONDS", "600"), 600),
        download_dir=normalize_path_value(str(payload.get("DOWNLOAD_DIR", "downloads"))),
        state_file=normalize_path_value(str(payload.get("STATE_FILE", "state/mail_state.json"))),
        webhook_send_url=webhook_send_url,
        webhook_upload_url=webhook_upload_url,
        max_attachment_size_mb=parse_positive_int(payload.get("MAX_ATTACHMENT_SIZE_MB", "30"), 30),
        window_width=parse_positive_int(payload.get("WINDOW_WIDTH", "960"), 960),
        window_height=parse_positive_int(payload.get("WINDOW_HEIGHT", "820"), 820),
        sidebar_width=parse_positive_int(payload.get("SIDEBAR_WIDTH", "220"), 220),
        ui_appearance=ui_appearance,
        ui_color_theme=str(payload.get("UI_COLOR_THEME", "blue")).strip() or "blue",
        start_page=start_page,
        ui_log_poll_ms=parse_positive_int(payload.get("UI_LOG_POLL_MS", "100"), 100),
        auto_scroll_log=parse_bool(payload.get("AUTO_SCROLL_LOG", "true"), True),
        ui_scale=parse_positive_float(payload.get("UI_SCALE", "1.0"), 1.0),
        script_timeout_seconds=parse_positive_int(payload.get("SCRIPT_TIMEOUT_SECONDS", "300"), 300),
        app_title=str(payload.get("APP_TITLE", "量子推送机器人 v5.1")).strip() or "量子推送机器人 v5.1",
        app_footer_text=str(payload.get("APP_FOOTER_TEXT", "v5.1\nby 不丢西瓜der")) or "v5.1\nby 不丢西瓜der",
    )
