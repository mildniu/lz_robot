from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, List, Literal, Optional
import struct
import hashlib
import os
import subprocess
import sys

from .attachment_service import extract_attachments_by_types
from .config import AppConfig, parse_upload_url
from .imap_client import ImapMailClient
from .mailbox_store import load_mailbox_alias_map
from .mime_utils import decode_mime_text
from .state_store import JsonStateStore
from .subject_attachment_rules import (
    list_enabled_rules,
)
from .webhook_client import QuantumWebhookClient

EventCallback = Optional[Callable[[str, str], None]]
IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg"}


@dataclass
class MailProcessingResult:
    status: Literal["processed", "skipped", "not_found"]
    uid: str = ""
    subject: str = ""
    sender: str = ""
    date: str = ""
    files: List[Path] = field(default_factory=list)
    reason: str = ""
    next_poll_interval_seconds: int = 0


@dataclass
class RuleProcessingResult:
    status: Literal["processed", "skipped", "not_found", "error"]
    rule_keyword: str = ""
    mailbox_alias: str = ""
    mailbox_folder: str = ""
    uid: str = ""
    subject: str = ""
    sender: str = ""
    date: str = ""
    files: List[Path] = field(default_factory=list)
    reason: str = ""


@dataclass
class BatchProcessingResult:
    status: Literal["processed", "skipped", "not_found"]
    results: List[RuleProcessingResult] = field(default_factory=list)
    next_poll_interval_seconds: int = 0


def emit_event(callback: EventCallback, level: str, message: str) -> None:
    if callback:
        callback(level, message)


def build_webhook_client(send_url: str) -> QuantumWebhookClient:
    return QuantumWebhookClient(send_url=send_url, upload_url=parse_upload_url(send_url))


def is_image_file(file_path: Path) -> bool:
    return file_path.suffix.lower() in IMAGE_SUFFIXES


def get_image_dimensions(file_path: Path) -> tuple[int, int]:
    suffix = file_path.suffix.lower()
    if suffix == ".png":
        with file_path.open("rb") as file_handle:
            header = file_handle.read(24)
        if len(header) < 24 or header[:8] != b"\x89PNG\r\n\x1a\n":
            raise RuntimeError(f"Invalid PNG file: {file_path.name}")
        width, height = struct.unpack(">II", header[16:24])
        return int(width), int(height)

    if suffix in {".jpg", ".jpeg"}:
        with file_path.open("rb") as file_handle:
            data = file_handle.read()
        if len(data) < 4 or data[0:2] != b"\xff\xd8":
            raise RuntimeError(f"Invalid JPEG file: {file_path.name}")

        index = 2
        data_len = len(data)
        while index + 9 < data_len:
            if data[index] != 0xFF:
                index += 1
                continue

            marker = data[index + 1]
            index += 2

            # Standalone markers without length field
            if marker in {0xD8, 0xD9, 0x01} or 0xD0 <= marker <= 0xD7:
                continue

            if index + 2 > data_len:
                break
            segment_length = struct.unpack(">H", data[index : index + 2])[0]
            if segment_length < 2 or index + segment_length > data_len:
                break

            if marker in {0xC0, 0xC1, 0xC2, 0xC3, 0xC9, 0xCA, 0xCB}:
                if index + 7 > data_len:
                    break
                height = struct.unpack(">H", data[index + 3 : index + 5])[0]
                width = struct.unpack(">H", data[index + 5 : index + 7])[0]
                return int(width), int(height)

            index += segment_length

        raise RuntimeError(f"Cannot parse JPEG dimensions: {file_path.name}")

    raise RuntimeError(f"Unsupported image type: {file_path.suffix}")


def send_file_via_webhook(
    file_path: Path,
    send_url: str,
    event_callback: EventCallback = None,
) -> str:
    emit_event(event_callback, "INFO", f"上传文件: {file_path.name}")
    webhook = build_webhook_client(send_url)
    is_image = is_image_file(file_path)
    upload_type = 1 if is_image else 2
    file_id, _ = webhook.upload_file_with_meta(file_path, upload_type=upload_type)
    emit_event(event_callback, "SUCCESS", f"文件上传成功: {file_id}")

    if is_image:
        try:
            width, height = get_image_dimensions(file_path)
            webhook.send_image_message(file_id, height=height, width=width)
            emit_event(event_callback, "SUCCESS", f"图片消息推送成功: {width}x{height}")
        except Exception as exc:
            emit_event(event_callback, "WARNING", f"图片消息发送失败，降级为文件消息: {exc}")
            webhook.send_file_message(file_id)
            emit_event(event_callback, "SUCCESS", "文件消息推送成功")
    else:
        webhook.send_file_message(file_id)
        emit_event(event_callback, "SUCCESS", "文件消息推送成功")
    return file_id


def run_rule_script(
    script_path: Path,
    attachment_path: Path,
    output_dir: Path,
    *,
    subject: str,
    sender: str,
    mail_date: str,
    rule_keyword: str,
    mailbox_alias: str,
    webhook_alias: str,
    webhook_url: str,
    event_callback: EventCallback = None,
) -> None:
    script_path = script_path.expanduser().resolve()
    attachment_path = attachment_path.expanduser().resolve()
    output_dir = output_dir.expanduser().resolve()

    output_dir.mkdir(parents=True, exist_ok=True)
    emit_event(event_callback, "INFO", f"执行脚本: {script_path.name}")
    emit_event(event_callback, "INFO", f"脚本输入附件: {attachment_path.name}")
    emit_event(event_callback, "INFO", f"脚本输入附件路径: {attachment_path}")
    emit_event(event_callback, "INFO", f"脚本输出目录: {output_dir}")

    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    env["LZ_RULE_KEYWORD"] = rule_keyword
    env["LZ_MAIL_SUBJECT"] = subject
    env["LZ_MAIL_SENDER"] = sender
    env["LZ_MAIL_DATE"] = mail_date
    env["LZ_MAILBOX_ALIAS"] = mailbox_alias
    env["LZ_WEBHOOK_ALIAS"] = webhook_alias
    env["LZ_WEBHOOK_URL"] = webhook_url
    env["LZ_WEBHOOK_UPLOAD_URL"] = parse_upload_url(webhook_url) if webhook_url else ""
    env["LZ_ATTACHMENT_PATH"] = str(attachment_path)
    env["LZ_OUTPUT_DIR"] = str(output_dir)

    command = [
        sys.executable,
        str(script_path),
        str(attachment_path),
        str(output_dir),
    ]
    completed = subprocess.run(
        command,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=300,
        cwd=str(script_path.parent),
        env=env,
    )
    stdout_text = (completed.stdout or "").strip()
    stderr_text = (completed.stderr or "").strip()
    if stdout_text:
        emit_event(event_callback, "INFO", f"脚本输出: {stdout_text}")
    if completed.returncode != 0:
        if stderr_text:
            emit_event(event_callback, "ERROR", f"脚本错误: {stderr_text}")
        raise RuntimeError(f"脚本执行失败，退出码={completed.returncode}")
    if stderr_text:
        emit_event(event_callback, "WARNING", f"脚本标准错误输出: {stderr_text}")
    emit_event(event_callback, "SUCCESS", f"脚本处理完成: {attachment_path.name}")


class MailProcessingService:
    def __init__(self, config: AppConfig, state_store: Optional[JsonStateStore] = None) -> None:
        self.config = config
        self.state_store = state_store or JsonStateStore(config.state_file)

    def _load_mailboxes(self) -> dict[str, dict]:
        alias_map = load_mailbox_alias_map()
        if alias_map:
            return alias_map

        default_alias = "默认邮箱"
        if (
            self.config.imap_host
            and self.config.imap_port > 0
            and self.config.email_username
            and self.config.email_password
        ):
            return {
                default_alias: {
                    "alias": default_alias,
                    "host": self.config.imap_host,
                    "port": self.config.imap_port,
                    "username": self.config.email_username,
                    "password": self.config.email_password,
                    "mailbox": self.config.imap_mailbox,
                }
            }
        return {}

    def _build_rule_state_key(self, rule: dict) -> str:
        identity_parts = [
            str(rule.get("mailbox_alias", "")).strip(),
            str(rule.get("keyword", "")).strip(),
            ",".join(sorted(str(item).strip() for item in rule.get("types", []) if str(item).strip())),
            ",".join(
                sorted(str(item).strip() for item in rule.get("filename_keywords", []) if str(item).strip())
            ),
            str(rule.get("webhook_alias", "")).strip(),
            str(rule.get("webhook_url", "")).strip(),
            str(rule.get("script_path", "")).strip(),
            str(rule.get("script_output_dir", "")).strip(),
        ]
        digest = hashlib.sha1("|".join(identity_parts).encode("utf-8")).hexdigest()[:16]
        return f"rule_uid::{digest}"

    def _process_rule(
        self,
        rule: dict,
        mailbox_config: dict,
        *,
        force: bool = False,
        update_state: bool = True,
        event_callback: EventCallback = None,
    ) -> RuleProcessingResult:
        keyword = str(rule.get("keyword", "")).strip()
        mailbox_alias = str(rule.get("mailbox_alias", "")).strip() or str(mailbox_config.get("alias", "")).strip()
        mailbox_folder = str(mailbox_config.get("mailbox", "INBOX")).strip() or "INBOX"
        target_send_url = str(rule.get("webhook_url", "")).strip()
        webhook_alias = str(rule.get("webhook_alias", "")).strip()
        script_path_text = str(rule.get("script_path", "")).strip()
        script_output_dir_text = str(rule.get("script_output_dir", "")).strip()
        allowed_types = [str(item).strip() for item in rule.get("types", []) if str(item).strip()]
        filename_keywords = [
            str(item).strip() for item in rule.get("filename_keywords", []) if str(item).strip()
        ]
        rule_max_size_mb = rule.get("max_attachment_size_mb")
        rule_state_key = self._build_rule_state_key(rule)
        last_uid = self.state_store.get_last_sent_uid(rule_state_key)

        emit_event(
            event_callback,
            "INFO",
            f"检查规则“{keyword}”所属邮箱: {mailbox_alias or mailbox_config.get('alias', '')} / 文件夹: {mailbox_folder}",
        )
        emit_event(event_callback, "INFO", f"IMAP 登录成功，开始在文件夹“{mailbox_folder}”中搜索主题关键字")

        with ImapMailClient(
            host=str(mailbox_config.get("host", "")).strip(),
            port=int(mailbox_config.get("port", 0)),
            username=str(mailbox_config.get("username", "")).strip(),
            password=str(mailbox_config.get("password", "")),
            mailbox=mailbox_folder,
        ) as imap_client:
            latest_uid = imap_client.find_latest_uid([keyword])
            if not latest_uid:
                return RuleProcessingResult(
                    status="not_found",
                    rule_keyword=keyword,
                    mailbox_alias=mailbox_alias,
                    mailbox_folder=mailbox_folder,
                    reason=(
                        f"登录成功，但邮箱“{mailbox_alias}”的文件夹“{mailbox_folder}”中"
                        f"未找到主题包含“{keyword}”的邮件"
                    ),
                )

            latest_uid_str = latest_uid.decode()
            if not force and latest_uid_str == last_uid:
                return RuleProcessingResult(
                    status="skipped",
                    rule_keyword=keyword,
                    mailbox_alias=mailbox_alias,
                    mailbox_folder=mailbox_folder,
                    uid=latest_uid_str,
                    reason=(
                        f"找到匹配邮件，但 UID={latest_uid_str} 与上次处理记录相同，"
                        f"规则“{keyword}”暂无新邮件"
                    ),
                )

            message = imap_client.fetch_message(latest_uid)
            subject = decode_mime_text(message.get("Subject"))
            sender = decode_mime_text(message.get("From"))
            date = message.get("Date", "")

            emit_event(event_callback, "INFO", f"规则“{keyword}”命中主题: {subject}")
            emit_event(event_callback, "INFO", f"发件人: {sender}")
            emit_event(event_callback, "INFO", f"日期: {date}")

            if not target_send_url:
                return RuleProcessingResult(
                    status="skipped",
                    rule_keyword=keyword,
                    mailbox_alias=mailbox_alias,
                    mailbox_folder=mailbox_folder,
                    uid=latest_uid_str,
                    subject=subject,
                    sender=sender,
                    date=date,
                    reason=f"规则“{keyword}”未配置推送机器人",
                )
            if not isinstance(rule_max_size_mb, int) or rule_max_size_mb <= 0:
                return RuleProcessingResult(
                    status="skipped",
                    rule_keyword=keyword,
                    mailbox_alias=mailbox_alias,
                    mailbox_folder=mailbox_folder,
                    uid=latest_uid_str,
                    subject=subject,
                    sender=sender,
                    date=date,
                    reason=f"规则“{keyword}”最大附件无效",
                )
            rule_max_size_bytes = rule_max_size_mb * 1024 * 1024

            types_text = ", ".join(allowed_types)
            names_text = ", ".join(filename_keywords) if filename_keywords else "(未启用)"
            if script_path_text:
                emit_event(event_callback, "INFO", f"处理模式: Python 脚本 -> {script_path_text}")
                if script_output_dir_text:
                    emit_event(event_callback, "INFO", f"脚本输出目录: {script_output_dir_text}")
            if webhook_alias:
                emit_event(event_callback, "INFO", f"推送机器人别名: {webhook_alias}")
            emit_event(event_callback, "INFO", f"附件格式 [{types_text}]")
            emit_event(event_callback, "INFO", f"附件文件名关键字过滤: {names_text}")
            emit_event(event_callback, "INFO", f"规则最大附件: {rule_max_size_mb} MB")
            emit_event(event_callback, "INFO", "提取目标附件...")

            files = extract_attachments_by_types(
                message,
                self.config.download_dir,
                latest_uid,
                allowed_types,
                filename_keywords=filename_keywords,
            )
            if not files:
                raise RuntimeError(f"规则“{keyword}”邮件没有匹配附件: [{types_text}]")

            for file_path in files:
                file_size = file_path.stat().st_size
                if file_size > rule_max_size_bytes:
                    raise RuntimeError(
                        f"Attachment too large (>{rule_max_size_mb}MB): {file_path.name}"
                    )
                if script_path_text:
                    script_path = Path(script_path_text).expanduser()
                    if not script_path.exists() or not script_path.is_file():
                        raise RuntimeError(f"脚本文件不存在: {script_path}")
                    if not script_output_dir_text.strip():
                        raise RuntimeError("脚本输出目录不能为空")
                    output_dir = Path(script_output_dir_text).expanduser()
                    run_rule_script(
                        script_path,
                        file_path,
                        output_dir,
                        subject=subject,
                        sender=sender,
                        mail_date=date,
                        rule_keyword=keyword,
                        mailbox_alias=mailbox_alias,
                        webhook_alias=webhook_alias,
                        webhook_url=target_send_url,
                        event_callback=event_callback,
                    )
                else:
                    send_file_via_webhook(
                        file_path,
                        target_send_url,
                        event_callback=event_callback,
                    )

            if update_state:
                self.state_store.set_last_sent_uid(latest_uid_str, rule_state_key)

            return RuleProcessingResult(
                status="processed",
                rule_keyword=keyword,
                mailbox_alias=mailbox_alias,
                mailbox_folder=mailbox_folder,
                uid=latest_uid_str,
                subject=subject,
                sender=sender,
                date=date,
                files=files,
            )

    def process_rule_batch(
        self,
        *,
        force: bool = False,
        update_state: bool = True,
        event_callback: EventCallback = None,
    ) -> BatchProcessingResult:
        enabled_rules = list_enabled_rules()
        if not enabled_rules:
            return BatchProcessingResult(
                status="skipped",
                results=[
                    RuleProcessingResult(
                        status="skipped",
                        reason="未配置启用的邮箱检测规则，请在“邮箱检测规则”中至少启用一条",
                    )
                ],
                next_poll_interval_seconds=0,
            )

        enabled_intervals = [
            int(item.get("poll_interval_seconds"))
            for item in enabled_rules
            if isinstance(item.get("poll_interval_seconds"), int) and int(item.get("poll_interval_seconds")) > 0
        ]
        next_interval = min(enabled_intervals) if enabled_intervals else self.config.poll_interval_seconds
        mailbox_alias_map = self._load_mailboxes()
        if not mailbox_alias_map:
            return BatchProcessingResult(
                status="skipped",
                results=[RuleProcessingResult(status="skipped", reason="未配置可用的 IMAP 邮箱别名")],
                next_poll_interval_seconds=next_interval,
            )

        results: List[RuleProcessingResult] = []
        for rule in enabled_rules:
            mailbox_alias = str(rule.get("mailbox_alias", "")).strip()
            if not mailbox_alias and len(mailbox_alias_map) == 1:
                mailbox_alias = next(iter(mailbox_alias_map.keys()), "")
            mailbox_config = mailbox_alias_map.get(mailbox_alias)
            keyword = str(rule.get("keyword", "")).strip()

            if not mailbox_alias:
                results.append(
                    RuleProcessingResult(
                        status="skipped",
                        rule_keyword=keyword,
                        mailbox_folder="",
                        reason=f"规则“{keyword}”未选择邮箱别名",
                    )
                )
                continue
            if not mailbox_config:
                results.append(
                    RuleProcessingResult(
                        status="skipped",
                        rule_keyword=keyword,
                        mailbox_alias=mailbox_alias,
                        mailbox_folder="",
                        reason=f"规则“{keyword}”邮箱别名无效: {mailbox_alias}",
                    )
                )
                continue

            try:
                rule_result = self._process_rule(
                    rule,
                    mailbox_config,
                    force=force,
                    update_state=update_state,
                    event_callback=event_callback,
                )
            except Exception as exc:
                rule_result = RuleProcessingResult(
                    status="error",
                    rule_keyword=keyword,
                    mailbox_alias=mailbox_alias,
                    mailbox_folder=str(mailbox_config.get("mailbox", "INBOX")).strip() or "INBOX",
                    reason=str(exc),
                )
            results.append(rule_result)

        processed_count = sum(1 for item in results if item.status == "processed")
        not_found_count = sum(1 for item in results if item.status == "not_found")
        if processed_count > 0:
            status: Literal["processed", "skipped", "not_found"] = "processed"
        elif not_found_count == len(results):
            status = "not_found"
        else:
            status = "skipped"

        return BatchProcessingResult(
            status=status,
            results=results,
            next_poll_interval_seconds=next_interval,
        )

    def process_latest_mail(
        self,
        *,
        force: bool = False,
        update_state: bool = True,
        event_callback: EventCallback = None,
    ) -> MailProcessingResult:
        batch_result = self.process_rule_batch(
            force=force,
            update_state=update_state,
            event_callback=event_callback,
        )
        first_processed = next((item for item in batch_result.results if item.status == "processed"), None)
        first_non_ok = batch_result.results[0] if batch_result.results else None
        selected = first_processed or first_non_ok

        return MailProcessingResult(
            status=batch_result.status,
            uid=(selected.uid if selected else ""),
            subject=(selected.subject if selected else ""),
            sender=(selected.sender if selected else ""),
            date=(selected.date if selected else ""),
            files=(selected.files if selected else []),
            reason=(selected.reason if selected else ""),
            next_poll_interval_seconds=batch_result.next_poll_interval_seconds,
        )
