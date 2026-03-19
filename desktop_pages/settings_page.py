import json
import math
import threading
import tkinter as tk
import tkinter.messagebox as messagebox
from pathlib import Path
from tkinter import filedialog
from typing import Dict, Optional

import customtkinter as ctk

from mail_forwarder import AppConfig, load_config
from mail_forwarder.config import upsert_env_file, validate_config_values
from mail_forwarder.imap_client import ImapMailClient
from mail_forwarder.mailbox_store import load_mailbox_configs, save_mailbox_configs
from mail_forwarder.subject_attachment_rules import (
    load_subject_attachment_rules,
    parse_filename_keywords_input,
    parse_types_input,
    save_subject_attachment_rules,
)

from .common import ModernButton
from .webhook_alias_store import load_webhook_aliases, save_webhook_aliases


class SettingsPage(ctk.CTkScrollableFrame):
    NO_ALIAS_LABEL = "(未配置)"
    NO_MAILBOX_LABEL = "(未配置邮箱)"
    THEME_LABEL_TO_VALUE = {
        "标准蓝色": "blue",
        "清新绿色": "green",
        "深邃蓝色": "dark-blue",
    }
    START_PAGE_LABEL_TO_VALUE = {
        "邮件检测": "execute",
        "文件夹检测": "folder",
        "机器人测试": "bot_test",
        "设置": "settings",
        "关于": "about",
    }
    TRIGGER_LABEL_TO_VALUE = {
        "周期检测": "periodic",
        "定时检测": "timed",
    }

    def __init__(self, master, config: Optional[AppConfig] = None, on_config_changed=None, **kwargs):
        super().__init__(master, **kwargs)
        self.on_config_changed = on_config_changed
        self.config = config or self.load_current_config()
        self.alias_slot_count = 6
        self.mailbox_slot_count = 6
        self.subject_rule_slot_count = 6
        self.folder_slot_count = 3
        self._alias_refresh_job = None
        self.entries = {}
        self.rule_checkboxes = {}
        self.rule_alias_vars = {}
        self.rule_alias_menus = {}
        self.rule_mailbox_vars = {}
        self.rule_mailbox_menus = {}
        self.rule_detail_frames = {}
        self.rule_toggle_buttons = {}
        self.rule_summary_labels = {}
        self.mailbox_status_labels = {}
        self.folder_checkboxes = {}
        self.folder_alias_vars = {}
        self.folder_alias_menus = {}
        self.ui_vars = {}
        self.alias_config = self.load_alias_config()
        self.mailbox_config = self.load_mailbox_config()
        self.subject_rules_payload = load_subject_attachment_rules()
        self.setup_ui()

    def _notify_config_changed(self):
        self.config = self.load_current_config()
        self.alias_config = self.load_alias_config()
        self.mailbox_config = self.load_mailbox_config()
        self.subject_rules_payload = load_subject_attachment_rules()
        if callable(self.on_config_changed):
            self.on_config_changed()

    def load_current_config(self) -> Optional[AppConfig]:
        try:
            return load_config()
        except Exception:
            return None

    def load_alias_config(self) -> dict:
        alias_config = load_webhook_aliases()
        aliases = alias_config.get("aliases", {})

        if not aliases and self.config and self.config.webhook_send_url:
            aliases = {"默认机器人": self.config.webhook_send_url}

        email_alias = alias_config.get("email_alias", "")
        if email_alias not in aliases:
            email_alias = next(iter(aliases.keys()), "")

        return {"aliases": aliases, "email_alias": email_alias}

    @staticmethod
    def interval_seconds_to_minutes(seconds: int) -> int:
        try:
            value = int(seconds)
        except (TypeError, ValueError):
            return 1
        if value <= 0:
            return 1
        return max(1, math.ceil(value / 60))

    @staticmethod
    def interval_minutes_to_seconds(minutes: int) -> int:
        return max(1, int(minutes)) * 60

    def load_mailbox_config(self) -> dict:
        mailbox_config = load_mailbox_configs()
        mailboxes = mailbox_config.get("mailboxes", [])
        if not mailboxes and self.config:
            if (
                self.config.imap_host
                and self.config.imap_port > 0
                and self.config.email_username
                and self.config.email_password
            ):
                mailboxes = [
                    {
                        "alias": "默认邮箱",
                        "host": self.config.imap_host,
                        "port": self.config.imap_port,
                        "username": self.config.email_username,
                        "password": self.config.email_password,
                        "mailbox": self.config.imap_mailbox,
                    }
                ]
        return {"mailboxes": mailboxes}

    @staticmethod
    def _label_for_value(mapping: Dict[str, str], value: str, default_label: str) -> str:
        for label, mapped_value in mapping.items():
            if mapped_value == value:
                return label
        return default_label

    @staticmethod
    def _value_for_label(mapping: Dict[str, str], label: str, default_value: str) -> str:
        return mapping.get(label, default_value)

    def get_alias_option_values(self, aliases: Optional[Dict[str, str]] = None) -> list[str]:
        alias_map = aliases if aliases is not None else self.alias_config.get("aliases", {})
        names = sorted([name for name in alias_map.keys() if name.strip()])
        return [self.NO_ALIAS_LABEL] + names

    def get_mailbox_option_values(self, mailboxes: Optional[list[dict]] = None) -> list[str]:
        mailbox_items = mailboxes if mailboxes is not None else self.mailbox_config.get("mailboxes", [])
        names = sorted(
            [
                str(item.get("alias", "")).strip()
                for item in mailbox_items
                if str(item.get("alias", "")).strip()
            ]
        )
        return [self.NO_MAILBOX_LABEL] + names

    def find_alias_by_url(self, target_url: str) -> str:
        if not target_url:
            return ""
        for name, url in self.alias_config.get("aliases", {}).items():
            if url.strip() == target_url.strip():
                return name
        return ""

    def collect_aliases_from_inputs(self, strict: bool = True) -> tuple[Dict[str, str], list[str]]:
        aliases = {}
        errors = []

        for index in range(1, self.alias_slot_count + 1):
            name_entry = self.entries.get(f"alias_{index}_name")
            url_entry = self.entries.get(f"alias_{index}_url")
            name = name_entry.get().strip() if name_entry else ""
            url = url_entry.get().strip() if url_entry else ""

            if not name and not url:
                continue

            if strict and (not name or not url):
                errors.append(f"别名{index}: 名称和 URL 必须同时填写")
                continue
            if not name or not url:
                continue

            if name in aliases:
                errors.append(f"别名重复: {name}")
                continue

            if strict and not url.startswith(("http://", "https://")):
                errors.append(f"别名 {name}: URL 必须以 http:// 或 https:// 开头")
                continue

            aliases[name] = url

        return aliases, errors

    def collect_mailboxes_from_inputs(self, strict: bool = True) -> tuple[list[dict], list[str]]:
        mailboxes = []
        errors = []
        seen_aliases = set()

        for index in range(1, self.mailbox_slot_count + 1):
            alias = self.entries.get(f"mailbox_{index}_alias").get().strip() if self.entries.get(f"mailbox_{index}_alias") else ""
            host = self.entries.get(f"mailbox_{index}_host").get().strip() if self.entries.get(f"mailbox_{index}_host") else ""
            port = self.entries.get(f"mailbox_{index}_port").get().strip() if self.entries.get(f"mailbox_{index}_port") else ""
            username = self.entries.get(f"mailbox_{index}_username").get().strip() if self.entries.get(f"mailbox_{index}_username") else ""
            password = self.entries.get(f"mailbox_{index}_password").get() if self.entries.get(f"mailbox_{index}_password") else ""
            mailbox = self.entries.get(f"mailbox_{index}_folder").get().strip() if self.entries.get(f"mailbox_{index}_folder") else ""

            if not alias and not host and not port and not username and not password and not mailbox:
                continue

            if strict and (not alias or not host or not port or not username or not password):
                errors.append(f"邮箱{index}: 别名、服务器、端口、邮箱账号、密码必须同时填写")
                continue
            if not alias or not host or not port or not username or not password:
                continue
            if alias in seen_aliases:
                errors.append(f"邮箱别名重复: {alias}")
                continue
            try:
                port_value = int(port)
                if port_value <= 0:
                    raise ValueError
            except ValueError:
                errors.append(f"邮箱{index}: IMAP 端口必须是大于 0 的整数")
                continue

            seen_aliases.add(alias)
            mailboxes.append(
                {
                    "alias": alias,
                    "host": host,
                    "port": port_value,
                    "username": username,
                    "password": password,
                    "mailbox": mailbox or "INBOX",
                }
            )

        return mailboxes, errors

    def schedule_alias_refresh(self, _event=None):
        if self._alias_refresh_job is not None:
            self.after_cancel(self._alias_refresh_job)
        self._alias_refresh_job = self.after(120, self.refresh_alias_options)

    def refresh_alias_options(self):
        self._alias_refresh_job = None
        aliases, _ = self.collect_aliases_from_inputs(strict=False)
        option_values = self.get_alias_option_values(aliases)

        for key, alias_menu in self.rule_alias_menus.items():
            alias_var = self.rule_alias_vars.get(key)
            alias_menu.configure(values=option_values)
            if alias_var and alias_var.get() not in option_values:
                alias_var.set(option_values[0])

        for key, alias_menu in self.folder_alias_menus.items():
            alias_var = self.folder_alias_vars.get(key)
            alias_menu.configure(values=option_values)
            if alias_var and alias_var.get() not in option_values:
                alias_var.set(option_values[0])
        self.refresh_rule_summaries_from_saved_payload()

    def refresh_mailbox_options(self):
        mailboxes, _ = self.collect_mailboxes_from_inputs(strict=False)
        option_values = self.get_mailbox_option_values(mailboxes)

        for key, mailbox_menu in self.rule_mailbox_menus.items():
            mailbox_var = self.rule_mailbox_vars.get(key)
            mailbox_menu.configure(values=option_values)
            if mailbox_var and mailbox_var.get() not in option_values:
                mailbox_var.set(option_values[0])
        self.refresh_rule_summaries_from_saved_payload()

    def _build_rule_summary_text(
        self,
        *,
        enabled: bool,
        keyword: str,
        mailbox_alias: str,
        robot_alias: str,
        script_path: str,
        trigger_mode: str = "periodic",
        schedule_time: str = "",
        poll_interval_seconds: int = 0,
    ) -> str:
        status_text = "已启用" if enabled else "未启用"
        keyword_text = keyword or "未填写主题"
        mailbox_text = mailbox_alias if mailbox_alias and mailbox_alias != self.NO_MAILBOX_LABEL else "未选邮箱"
        robot_text = robot_alias if robot_alias and robot_alias != self.NO_ALIAS_LABEL else "未选机器人"
        mode_text = "脚本处理" if script_path.strip() else "直接推送"
        trigger_text = (
            f"定时 {schedule_time or '--:--'}"
            if trigger_mode == "timed"
            else f"周期 {self.format_rule_interval_text(poll_interval_seconds)}"
        )
        return f"{status_text} | {keyword_text} | {mailbox_text} -> {robot_text} | {trigger_text} | {mode_text}"

    @staticmethod
    def format_rule_interval_text(seconds: int) -> str:
        try:
            value = int(seconds)
        except (TypeError, ValueError):
            value = 60
        if value % 60 == 0:
            return f"{max(1, value // 60)} 分钟"
        return f"{value} 秒"

    def toggle_rule_card(self, index: int):
        detail_frame = self.rule_detail_frames.get(index)
        toggle_btn = self.rule_toggle_buttons.get(index)
        if not detail_frame or not toggle_btn:
            return
        if detail_frame.winfo_manager():
            detail_frame.pack_forget()
            toggle_btn.configure(text="展开")
        else:
            detail_frame.pack(fill="x", padx=14, pady=(0, 12))
            toggle_btn.configure(text="收起")

    def refresh_rule_summaries_from_saved_payload(self):
        saved_rules = self.subject_rules_payload.get("rules", [])
        for index in range(1, self.subject_rule_slot_count + 1):
            summary_label = self.rule_summary_labels.get(index)
            if not summary_label:
                continue
            rule_data = saved_rules[index - 1] if index <= len(saved_rules) else {}
            trigger_mode = str(rule_data.get("trigger_mode", "periodic")).strip() or "periodic"
            summary_label.configure(
                text=self._build_rule_summary_text(
                    enabled=bool(rule_data.get("enabled", False)) if rule_data else False,
                    keyword=str(rule_data.get("keyword", "")).strip(),
                    mailbox_alias=str(rule_data.get("mailbox_alias", "")).strip(),
                    robot_alias=str(rule_data.get("webhook_alias", "")).strip(),
                    script_path=str(rule_data.get("script_path", "")).strip(),
                    trigger_mode=trigger_mode,
                    schedule_time=str(rule_data.get("schedule_time", "")).strip(),
                    poll_interval_seconds=rule_data.get("poll_interval_seconds") or self.config.poll_interval_seconds,
                )
            )

    def setup_ui(self):
        title = ctk.CTkLabel(self, text="⚙️ 系统设置", font=ctk.CTkFont(size=28, weight="bold"))
        title.pack(pady=(20, 12))

        if not self.config:
            warning = ctk.CTkLabel(
                self,
                text="⚠️ 配置加载失败，请检查 settings/app_config.json",
                font=ctk.CTkFont(size=14),
                text_color="yellow",
            )
            warning.pack(pady=20)
            return

        tabview = ctk.CTkTabview(self, height=680)
        tabview.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        rules_tab = tabview.add("邮箱检测规则")
        email_tab = tabview.add("邮箱配置")
        alias_tab = tabview.add("机器人别名")
        folder_tab = tabview.add("文件夹检测")
        path_tab = tabview.add("路径设置")
        ui_tab = tabview.add("界面设置")

        self.build_rules_tab(rules_tab)
        self.build_email_tab(email_tab)
        self.build_alias_tab(alias_tab)
        self.build_folder_tab(folder_tab)
        self.build_path_tab(path_tab)
        self.build_ui_tab(ui_tab)
        tabview.set("邮箱检测规则")

    def create_config_card(self, parent, title: str, fields: list):
        card = ctk.CTkFrame(parent, fg_color=("gray95", "gray20"))
        card.pack(fill="x", padx=12, pady=10)

        card_title = ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=18, weight="bold"))
        card_title.pack(pady=(15, 10), anchor="w", padx=15)

        fields_frame = ctk.CTkFrame(card, fg_color="transparent")
        fields_frame.pack(fill="x", padx=15, pady=(0, 15))

        for label_text, field_name, default_value, *is_password in fields:
            row = ctk.CTkFrame(fields_frame, fg_color="transparent")
            row.pack(fill="x", pady=8)

            label = ctk.CTkLabel(
                row,
                text=label_text,
                width=120,
                anchor="w",
                font=ctk.CTkFont(size=14),
            )
            label.pack(side="left", padx=(0, 10))

            show = "•" if is_password and is_password[0] else None
            entry = ctk.CTkEntry(
                row,
                placeholder_text=default_value,
                show=show,
                height=38,
                font=ctk.CTkFont(size=13),
            )
            entry.insert(0, default_value)
            entry.pack(side="left", fill="x", expand=True)
            self.entries[field_name] = entry

    def _get_mailbox_row_values(self, index: int) -> dict[str, str]:
        return {
            "alias": self.entries.get(f"mailbox_{index}_alias").get().strip() if self.entries.get(f"mailbox_{index}_alias") else "",
            "host": self.entries.get(f"mailbox_{index}_host").get().strip() if self.entries.get(f"mailbox_{index}_host") else "",
            "port": self.entries.get(f"mailbox_{index}_port").get().strip() if self.entries.get(f"mailbox_{index}_port") else "",
            "username": self.entries.get(f"mailbox_{index}_username").get().strip() if self.entries.get(f"mailbox_{index}_username") else "",
            "password": self.entries.get(f"mailbox_{index}_password").get() if self.entries.get(f"mailbox_{index}_password") else "",
            "mailbox": self.entries.get(f"mailbox_{index}_folder").get().strip() if self.entries.get(f"mailbox_{index}_folder") else "",
        }

    def update_mailbox_status(
        self,
        index: int,
        text: str,
        text_color: str | tuple[str, str] = "gray",
    ):
        label = self.mailbox_status_labels.get(index)
        if not label:
            return
        label.configure(text=text, text_color=text_color)

    def reset_mailbox_status(self, index: int):
        self.update_mailbox_status(index, "未测试", ("gray35", "gray70"))

    def validate_single_mailbox(self, index: int) -> tuple[Optional[dict], list[str]]:
        values = self._get_mailbox_row_values(index)
        errors = []
        if not any(values.values()):
            errors.append(f"邮箱{index}: 当前行还是空的，请先填写连接参数")
            return None, errors

        required_fields = {
            "alias": "邮箱别名",
            "host": "服务器地址",
            "port": "端口",
            "username": "邮箱账号",
            "password": "密码/授权码",
        }
        for key, label in required_fields.items():
            if not values.get(key, "").strip():
                errors.append(f"邮箱{index}: {label}不能为空")

        try:
            port_value = int(values.get("port", "0"))
            if port_value <= 0:
                raise ValueError
        except ValueError:
            errors.append(f"邮箱{index}: 端口必须是大于 0 的整数")
            port_value = 0

        if errors:
            return None, errors

        return {
            "alias": values["alias"],
            "host": values["host"],
            "port": port_value,
            "username": values["username"],
            "password": values["password"],
            "mailbox": values["mailbox"] or "INBOX",
        }, []

    def test_mailbox_connection(self, index: int):
        mailbox, errors = self.validate_single_mailbox(index)
        if errors:
            self.update_mailbox_status(index, "未通过", "#EF5350")
            messagebox.showerror("验证失败", "\n".join(f"• {error}" for error in errors))
            return

        alias = str(mailbox.get("alias", "")).strip()
        host = str(mailbox.get("host", "")).strip()
        port = int(mailbox.get("port", 0))
        username = str(mailbox.get("username", "")).strip()
        folder = str(mailbox.get("mailbox", "INBOX")).strip() or "INBOX"
        self.update_mailbox_status(index, "检测中...", "#FFB300")

        progress_dialog = tk.Toplevel(self)
        progress_dialog.title("测试邮箱连接")
        progress_dialog.geometry("420x150")
        progress_dialog.resizable(False, False)
        progress_dialog.transient(self.winfo_toplevel())
        progress_dialog.grab_set()

        message_var = tk.StringVar(value=f"正在测试邮箱“{alias}”连接...\n{host}:{port}\n文件夹: {folder}")
        ctk.CTkLabel(progress_dialog, textvariable=message_var, justify="left", font=ctk.CTkFont(size=13)).pack(
            padx=18, pady=(18, 12), anchor="w"
        )
        close_btn = ctk.CTkButton(progress_dialog, text="关闭", state="disabled", command=progress_dialog.destroy)
        close_btn.pack(pady=(0, 14))

        def finalize(title: str, body: str, is_error: bool):
            def _update():
                try:
                    message_var.set(body)
                    close_btn.configure(state="normal")
                    if is_error:
                        self.update_mailbox_status(index, "未通过", "#EF5350")
                    else:
                        self.update_mailbox_status(index, "已通过", "#2CC985")
                    if is_error:
                        messagebox.showerror(title, body, parent=progress_dialog)
                    else:
                        messagebox.showinfo(title, body, parent=progress_dialog)
                except Exception:
                    pass

            self.after(0, _update)

        def worker():
            try:
                with ImapMailClient(
                    host=host,
                    port=port,
                    username=username,
                    password=str(mailbox.get("password", "")),
                    mailbox=folder,
                    timeout_seconds=15,
                ):
                    pass
                finalize(
                    "连接成功",
                    f"邮箱“{alias}”连接成功\n\n服务器: {host}:{port}\n账号: {username}\n文件夹: {folder}",
                    False,
                )
            except Exception as exc:
                finalize(
                    "连接失败",
                    f"邮箱“{alias}”连接失败\n\n服务器: {host}:{port}\n账号: {username}\n文件夹: {folder}\n错误: {exc}",
                    True,
                )

        threading.Thread(target=worker, daemon=True).start()

    def build_email_tab(self, parent):
        card = ctk.CTkFrame(parent, fg_color=("gray95", "gray20"))
        card.pack(fill="x", padx=12, pady=10)

        card_title = ctk.CTkLabel(card, text="📧 IMAP 邮箱配置", font=ctk.CTkFont(size=18, weight="bold"))
        card_title.pack(pady=(15, 10), anchor="w", padx=15)

        tip = ctk.CTkLabel(
            card,
            text=(
                "这里维护所有可用的 IMAP 邮箱。每个邮箱都要先设置“邮箱别名”，"
                "后续规则就是靠这个别名去绑定具体邮箱。"
            ),
            font=ctk.CTkFont(size=12),
            text_color=("gray35", "gray70"),
        )
        tip.pack(anchor="w", padx=15, pady=(0, 8))

        guide = ctk.CTkLabel(
            card,
            text=(
                "填写建议：服务器一般是 imap 域名，端口常见为 993，文件夹通常填 INBOX。"
                "可先点右侧“测试连接”，确认当前这一行能连通后再保存。状态列会显示未测试/检测中/已通过/未通过。"
            ),
            font=ctk.CTkFont(size=12),
            text_color=("gray40", "gray68"),
            justify="left",
        )
        guide.pack(anchor="w", padx=15, pady=(0, 10))

        rows_frame = ctk.CTkFrame(card, fg_color="transparent")
        rows_frame.pack(fill="x", padx=15, pady=(0, 12))

        header_row = ctk.CTkFrame(rows_frame, fg_color="transparent")
        header_row.pack(fill="x", pady=(0, 4))
        for text, width in [
            ("邮箱别名", 120),
            ("服务器地址", 170),
            ("端口", 70),
            ("邮箱账号", 190),
            ("密码/授权码", 150),
            ("邮箱文件夹", 120),
            ("连接测试", 88),
            ("状态", 70),
        ]:
            ctk.CTkLabel(
                header_row,
                text=text,
                width=width,
                anchor="w",
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color=("gray25", "gray75"),
            ).pack(side="left", padx=(0, 8))

        mailboxes = self.mailbox_config.get("mailboxes", [])
        for index in range(1, self.mailbox_slot_count + 1):
            row = ctk.CTkFrame(rows_frame, fg_color="transparent")
            row.pack(fill="x", pady=4)

            mailbox_data = mailboxes[index - 1] if index <= len(mailboxes) else {}
            field_specs = [
                (f"mailbox_{index}_alias", f"邮箱别名{index}", mailbox_data.get("alias", ""), 120, None),
                (f"mailbox_{index}_host", "imap.example.com", mailbox_data.get("host", ""), 170, None),
                (f"mailbox_{index}_port", "993", str(mailbox_data.get("port", "")) if mailbox_data else "", 70, None),
                (f"mailbox_{index}_username", "邮箱账号", mailbox_data.get("username", ""), 190, None),
                (f"mailbox_{index}_password", "密码/授权码", mailbox_data.get("password", ""), 150, "•"),
                (f"mailbox_{index}_folder", "INBOX", mailbox_data.get("mailbox", ""), 120, None),
            ]

            for key, placeholder, value, width, show in field_specs:
                entry = ctk.CTkEntry(
                    row,
                    placeholder_text=placeholder,
                    height=32,
                    width=width,
                    font=ctk.CTkFont(size=12),
                    show=show,
                )
                if value:
                    entry.insert(0, value)
                entry.pack(side="left", padx=(0, 8))
                self.entries[key] = entry
                if key.endswith("_alias"):
                    entry.bind("<KeyRelease>", lambda _event, row_index=index: (self.refresh_mailbox_options(), self.reset_mailbox_status(row_index)))
                    entry.bind("<FocusOut>", lambda _event, row_index=index: (self.refresh_mailbox_options(), self.reset_mailbox_status(row_index)))
                else:
                    entry.bind("<KeyRelease>", lambda _event, row_index=index: self.reset_mailbox_status(row_index))
                    entry.bind("<FocusOut>", lambda _event, row_index=index: self.reset_mailbox_status(row_index))

            test_btn = ctk.CTkButton(
                row,
                text="测试连接",
                width=88,
                height=32,
                font=ctk.CTkFont(size=12),
                command=lambda row_index=index: self.test_mailbox_connection(row_index),
            )
            test_btn.pack(side="left", padx=(0, 8))

            status_label = ctk.CTkLabel(
                row,
                text="未测试",
                width=70,
                anchor="w",
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color=("gray35", "gray70"),
            )
            status_label.pack(side="left")
            self.mailbox_status_labels[index] = status_label

        save_btn = ModernButton(
            parent,
            text="保存邮箱配置",
            icon="💾",
            command=self.save_email_settings,
            height=42,
            width=220,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#2CC985",
            hover_color="#239B6D",
        )
        save_btn.pack(anchor="w", padx=12, pady=(6, 12))

    def build_rules_tab(self, parent):
        card = ctk.CTkFrame(parent, fg_color=("gray95", "gray20"))
        card.pack(fill="x", padx=12, pady=10)

        card_title = ctk.CTkLabel(card, text="📎 邮箱检测规则", font=ctk.CTkFont(size=18, weight="bold"))
        card_title.pack(pady=(15, 10), anchor="w", padx=15)

        rows_frame = ctk.CTkFrame(card, fg_color="transparent")
        rows_frame.pack(fill="x", padx=15, pady=(0, 12))

        rules = self.subject_rules_payload.get("rules", [])
        option_values = self.get_alias_option_values()
        mailbox_option_values = self.get_mailbox_option_values()
        initially_expanded_index = 0
        for index in range(1, self.subject_rule_slot_count + 1):
            rule_data = rules[index - 1] if index <= len(rules) else {}
            enabled = bool(rule_data.get("enabled", False)) if rule_data else False
            keyword = rule_data.get("keyword", "")
            types = (rule_data.get("types", [""]) or [""])[0]
            name_keywords = (rule_data.get("filename_keywords", [""]) or [""])[0]
            script_path = rule_data.get("script_path", "")
            script_output_dir = rule_data.get("script_output_dir", "")
            trigger_mode = str(rule_data.get("trigger_mode", "periodic")).strip() or "periodic"
            schedule_time = str(rule_data.get("schedule_time", "")).strip()
            poll_interval_seconds = rule_data.get("poll_interval_seconds") or self.config.poll_interval_seconds
            row_interval = str(
                self.interval_seconds_to_minutes(
                    poll_interval_seconds
                )
            )
            row_max_size = str(rule_data.get("max_attachment_size_mb") or self.config.max_attachment_size_mb)
            selected_mailbox_alias = rule_data.get("mailbox_alias", "")
            if not selected_mailbox_alias and len(mailbox_option_values) > 1:
                selected_mailbox_alias = mailbox_option_values[1]
            if selected_mailbox_alias not in mailbox_option_values:
                selected_mailbox_alias = self.NO_MAILBOX_LABEL
            selected_alias = rule_data.get("webhook_alias", "")
            if not selected_alias:
                selected_alias = self.find_alias_by_url(rule_data.get("webhook_url", ""))
            if not selected_alias and len(option_values) > 1:
                selected_alias = option_values[1]
            if selected_alias not in option_values:
                selected_alias = self.NO_ALIAS_LABEL

            rule_card = ctk.CTkFrame(rows_frame, fg_color=("gray92", "gray18"))
            rule_card.pack(fill="x", pady=8)

            card_header = ctk.CTkFrame(rule_card, fg_color="transparent")
            card_header.pack(fill="x", padx=14, pady=(12, 8))

            enabled_var = tk.BooleanVar(value=enabled)
            enable_box = ctk.CTkCheckBox(
                card_header,
                text="",
                width=20,
                onvalue=True,
                offvalue=False,
                variable=enabled_var,
                command=lambda: None,
            )
            if enabled:
                enable_box.select()
            enable_box.pack(side="left", padx=(0, 8))
            self.rule_checkboxes[f"rule_{index}_enabled"] = enable_box

            ctk.CTkLabel(
                card_header,
                text=f"规则 {index}",
                font=ctk.CTkFont(size=16, weight="bold"),
            ).pack(side="left")

            summary_label = ctk.CTkLabel(
                card_header,
                text=self._build_rule_summary_text(
                    enabled=enabled,
                    keyword=keyword,
                    mailbox_alias=selected_mailbox_alias,
                    robot_alias=selected_alias,
                    script_path=script_path,
                    trigger_mode=trigger_mode,
                    schedule_time=schedule_time,
                    poll_interval_seconds=poll_interval_seconds,
                ),
                font=ctk.CTkFont(size=11),
                text_color=("gray35", "gray70"),
            )
            summary_label.pack(side="left", padx=(10, 0))
            self.rule_summary_labels[index] = summary_label

            save_btn = ctk.CTkButton(
                card_header,
                text="保存规则",
                width=84,
                height=28,
                font=ctk.CTkFont(size=11, weight="bold"),
                fg_color="#D97706",
                hover_color="#B45309",
                text_color="white",
                command=lambda target=index: self.save_single_subject_rule(target),
            )
            save_btn.pack(side="right", padx=(0, 8))

            toggle_btn = ctk.CTkButton(
                card_header,
                text="收起" if index == initially_expanded_index else "展开",
                width=68,
                height=28,
                font=ctk.CTkFont(size=11),
                fg_color="#FB8C00",
                hover_color="#EF6C00",
                text_color="white",
                command=lambda target=index: self.toggle_rule_card(target),
            )
            toggle_btn.pack(side="right")
            self.rule_toggle_buttons[index] = toggle_btn

            detail_frame = ctk.CTkFrame(rule_card, fg_color="transparent")
            self.rule_detail_frames[index] = detail_frame
            if index == initially_expanded_index:
                detail_frame.pack(fill="x", padx=14, pady=(0, 12))

            basic_block = ctk.CTkFrame(detail_frame, fg_color="transparent")
            basic_block.pack(fill="x", pady=(0, 8))

            ctk.CTkLabel(
                basic_block,
                text="基础匹配",
                font=ctk.CTkFont(size=13, weight="bold"),
                text_color=("gray20", "gray75"),
            ).pack(anchor="w", pady=(0, 6))

            basic_row = ctk.CTkFrame(basic_block, fg_color="transparent")
            basic_row.pack(fill="x")

            keyword_col = ctk.CTkFrame(basic_row, fg_color="transparent")
            keyword_col.pack(side="left", fill="x", expand=True, padx=(0, 8))
            ctk.CTkLabel(keyword_col, text="主题关键字", anchor="w", font=ctk.CTkFont(size=12)).pack(anchor="w")
            keyword_entry = ctk.CTkEntry(
                keyword_col,
                placeholder_text=f"主题关键字{index} (例: 装维日报)",
                height=32,
                font=ctk.CTkFont(size=12),
            )
            keyword_entry.insert(0, keyword)
            keyword_entry.pack(fill="x", pady=(4, 0))
            self.entries[f"rule_{index}_keyword"] = keyword_entry

            type_col = ctk.CTkFrame(basic_row, fg_color="transparent")
            type_col.pack(side="left", fill="x", expand=True, padx=(0, 8))
            ctk.CTkLabel(type_col, text="附件格式", anchor="w", font=ctk.CTkFont(size=12)).pack(anchor="w")
            types_entry = ctk.CTkEntry(
                type_col,
                placeholder_text="附件类型 (例: png)",
                height=32,
                font=ctk.CTkFont(size=12),
            )
            types_entry.insert(0, types)
            types_entry.pack(fill="x", pady=(4, 0))
            self.entries[f"rule_{index}_types"] = types_entry

            name_col = ctk.CTkFrame(basic_row, fg_color="transparent")
            name_col.pack(side="left", fill="x", expand=True)
            ctk.CTkLabel(name_col, text="附件文件名关键字", anchor="w", font=ctk.CTkFont(size=12)).pack(anchor="w")
            name_entry = ctk.CTkEntry(
                name_col,
                placeholder_text="附件文件名关键字(可空, 只能填写一个)",
                height=32,
                font=ctk.CTkFont(size=12),
            )
            name_entry.insert(0, name_keywords)
            name_entry.pack(fill="x", pady=(4, 0))
            self.entries[f"rule_{index}_name_keywords"] = name_entry

            route_block = ctk.CTkFrame(detail_frame, fg_color="transparent")
            route_block.pack(fill="x", pady=(0, 8))
            ctk.CTkLabel(
                route_block,
                text="来源与推送",
                font=ctk.CTkFont(size=13, weight="bold"),
                text_color=("gray20", "gray75"),
            ).pack(anchor="w", pady=(0, 6))

            route_row = ctk.CTkFrame(route_block, fg_color="transparent")
            route_row.pack(fill="x")

            mailbox_col = ctk.CTkFrame(route_row, fg_color="transparent")
            mailbox_col.pack(side="left", fill="x", expand=True, padx=(0, 8))
            ctk.CTkLabel(mailbox_col, text="所属邮箱", anchor="w", font=ctk.CTkFont(size=12)).pack(anchor="w")
            mailbox_alias_var = tk.StringVar(value=selected_mailbox_alias)
            mailbox_alias_menu = ctk.CTkOptionMenu(
                mailbox_col,
                values=mailbox_option_values,
                variable=mailbox_alias_var,
                height=32,
                font=ctk.CTkFont(size=12),
            )
            mailbox_alias_menu.pack(fill="x", pady=(4, 0))
            rule_mailbox_key = f"rule_{index}_mailbox_alias"
            self.rule_mailbox_vars[rule_mailbox_key] = mailbox_alias_var
            self.rule_mailbox_menus[rule_mailbox_key] = mailbox_alias_menu

            alias_col = ctk.CTkFrame(route_row, fg_color="transparent")
            alias_col.pack(side="left", fill="x", expand=True)
            ctk.CTkLabel(alias_col, text="推送机器人", anchor="w", font=ctk.CTkFont(size=12)).pack(anchor="w")
            alias_var = tk.StringVar(value=selected_alias)
            alias_menu = ctk.CTkOptionMenu(
                alias_col,
                values=option_values,
                variable=alias_var,
                height=32,
                font=ctk.CTkFont(size=12),
            )
            alias_menu.pack(fill="x", pady=(4, 0))
            rule_alias_key = f"rule_{index}_alias"
            self.rule_alias_vars[rule_alias_key] = alias_var
            self.rule_alias_menus[rule_alias_key] = alias_menu

            script_block = ctk.CTkFrame(detail_frame, fg_color="transparent")
            script_block.pack(fill="x", pady=(0, 8))
            ctk.CTkLabel(
                script_block,
                text="脚本处理",
                font=ctk.CTkFont(size=13, weight="bold"),
                text_color=("gray20", "gray75"),
            ).pack(anchor="w", pady=(0, 6))

            script_hint = ctk.CTkLabel(
                script_block,
                text="不选脚本时由程序直接推送原附件；选了脚本时，脚本会接管处理并使用本规则选中的机器人推送。",
                font=ctk.CTkFont(size=11),
                text_color=("gray35", "gray68"),
                justify="left",
            )
            script_hint.pack(anchor="w", pady=(0, 6))

            script_row = ctk.CTkFrame(script_block, fg_color="transparent")
            script_row.pack(fill="x")

            script_path_col = ctk.CTkFrame(script_row, fg_color="transparent")
            script_path_col.pack(side="left", fill="x", expand=True, padx=(0, 8))
            ctk.CTkLabel(script_path_col, text="处理程序", anchor="w", font=ctk.CTkFont(size=12)).pack(anchor="w")
            script_entry = ctk.CTkEntry(
                script_path_col,
                placeholder_text="可选：选择 .py 或 .exe，处理程序会接管附件处理与推送",
                height=30,
                font=ctk.CTkFont(size=12),
            )
            script_entry.insert(0, script_path)
            script_entry.pack(fill="x", pady=(4, 0))
            self.entries[f"rule_{index}_script_path"] = script_entry

            script_path_actions = ctk.CTkFrame(script_row, fg_color="transparent")
            script_path_actions.pack(side="left", padx=(0, 8), pady=(22, 0))
            script_browse_btn = ctk.CTkButton(
                script_path_actions,
                text="选择程序",
                width=78,
                height=30,
                font=ctk.CTkFont(size=11),
                command=lambda entry=script_entry: self.browse_python_script(entry),
            )
            script_browse_btn.pack()

            output_col = ctk.CTkFrame(script_row, fg_color="transparent")
            output_col.pack(side="left", fill="x", expand=True, padx=(0, 8))
            ctk.CTkLabel(output_col, text="输出目录", anchor="w", font=ctk.CTkFont(size=12)).pack(anchor="w")
            output_entry = ctk.CTkEntry(
                output_col,
                placeholder_text="脚本生成的新文件留存在这里",
                height=30,
                font=ctk.CTkFont(size=12),
            )
            output_entry.insert(0, script_output_dir)
            output_entry.pack(fill="x", pady=(4, 0))
            self.entries[f"rule_{index}_script_output_dir"] = output_entry

            output_actions = ctk.CTkFrame(script_row, fg_color="transparent")
            output_actions.pack(side="left", pady=(22, 0))
            output_browse_btn = ctk.CTkButton(
                output_actions,
                text="输出目录",
                width=78,
                height=30,
                font=ctk.CTkFont(size=11),
                command=lambda entry=output_entry: self.browse_folder(entry),
            )
            output_browse_btn.pack()

            runtime_block = ctk.CTkFrame(detail_frame, fg_color="transparent")
            runtime_block.pack(fill="x")
            ctk.CTkLabel(
                runtime_block,
                text="运行参数",
                font=ctk.CTkFont(size=13, weight="bold"),
                text_color=("gray20", "gray75"),
            ).pack(anchor="w", pady=(0, 6))

            runtime_row = ctk.CTkFrame(runtime_block, fg_color="transparent")
            runtime_row.pack(fill="x")

            trigger_col = ctk.CTkFrame(runtime_row, fg_color="transparent")
            trigger_col.pack(side="left", padx=(0, 12))
            ctk.CTkLabel(trigger_col, text="检测方式", anchor="w", font=ctk.CTkFont(size=12)).pack(anchor="w")
            trigger_label = self._label_for_value(self.TRIGGER_LABEL_TO_VALUE, trigger_mode, "周期检测")
            trigger_var = tk.StringVar(value=trigger_label)
            trigger_menu = ctk.CTkOptionMenu(
                trigger_col,
                values=list(self.TRIGGER_LABEL_TO_VALUE.keys()),
                variable=trigger_var,
                height=32,
                width=120,
                font=ctk.CTkFont(size=12),
            )
            trigger_menu.pack(pady=(4, 0))
            self.ui_vars[f"rule_{index}_trigger_mode"] = trigger_var

            interval_col = ctk.CTkFrame(runtime_row, fg_color="transparent")
            interval_col.pack(side="left", padx=(0, 12))
            ctk.CTkLabel(interval_col, text="轮询间隔(min)", anchor="w", font=ctk.CTkFont(size=12)).pack(anchor="w")
            interval_entry = ctk.CTkEntry(
                interval_col,
                placeholder_text="轮询间隔(min)",
                height=32,
                width=120,
                font=ctk.CTkFont(size=12),
            )
            interval_entry.insert(0, row_interval)
            interval_entry.pack(pady=(4, 0))
            self.entries[f"rule_{index}_interval"] = interval_entry

            schedule_col = ctk.CTkFrame(runtime_row, fg_color="transparent")
            schedule_col.pack(side="left", padx=(0, 12))
            ctk.CTkLabel(schedule_col, text="定时时刻(HH:MM)", anchor="w", font=ctk.CTkFont(size=12)).pack(anchor="w")
            schedule_entry = ctk.CTkEntry(
                schedule_col,
                placeholder_text="例: 08:30",
                height=32,
                width=120,
                font=ctk.CTkFont(size=12),
            )
            schedule_entry.insert(0, schedule_time)
            schedule_entry.pack(pady=(4, 0))
            self.entries[f"rule_{index}_schedule_time"] = schedule_entry

            max_size_col = ctk.CTkFrame(runtime_row, fg_color="transparent")
            max_size_col.pack(side="left")
            ctk.CTkLabel(max_size_col, text="最大附件(MB)", anchor="w", font=ctk.CTkFont(size=12)).pack(anchor="w")
            max_size_entry = ctk.CTkEntry(
                max_size_col,
                placeholder_text="最大附件(MB)",
                height=32,
                width=120,
                font=ctk.CTkFont(size=12),
            )
            max_size_entry.insert(0, row_max_size)
            max_size_entry.pack(pady=(4, 0))
            self.entries[f"rule_{index}_max_size"] = max_size_entry

        self.refresh_rule_summaries_from_saved_payload()

    def build_alias_tab(self, parent):
        card = ctk.CTkFrame(parent, fg_color=("gray95", "gray20"))
        card.pack(fill="x", padx=12, pady=10)

        card_title = ctk.CTkLabel(card, text="📤 机器人别名配置", font=ctk.CTkFont(size=18, weight="bold"))
        card_title.pack(pady=(15, 10), anchor="w", padx=15)

        tip = ctk.CTkLabel(
            card,
            text="先维护别名和地址，再在邮件检测/文件夹检测中选择别名",
            font=ctk.CTkFont(size=12),
            text_color=("gray35", "gray70"),
        )
        tip.pack(anchor="w", padx=15, pady=(0, 8))

        aliases = list(self.alias_config.get("aliases", {}).items())
        aliases_frame = ctk.CTkFrame(card, fg_color="transparent")
        aliases_frame.pack(fill="x", padx=15, pady=(0, 10))

        for index in range(1, self.alias_slot_count + 1):
            row = ctk.CTkFrame(aliases_frame, fg_color="transparent")
            row.pack(fill="x", pady=4)

            name_default = aliases[index - 1][0] if index <= len(aliases) else ""
            url_default = aliases[index - 1][1] if index <= len(aliases) else ""

            name_entry = ctk.CTkEntry(
                row,
                placeholder_text=f"别名{index} (例: 售后群机器人)",
                height=32,
                font=ctk.CTkFont(size=12),
            )
            name_entry.configure(width=220)
            name_entry.insert(0, name_default)
            name_entry.pack(side="left", padx=(0, 8))
            self.entries[f"alias_{index}_name"] = name_entry
            name_entry.bind("<KeyRelease>", self.schedule_alias_refresh)
            name_entry.bind("<FocusOut>", self.schedule_alias_refresh)

            url_entry = ctk.CTkEntry(row, placeholder_text="Webhook URL", height=32, font=ctk.CTkFont(size=12))
            url_entry.insert(0, url_default)
            url_entry.pack(side="left", fill="x", expand=True)
            self.entries[f"alias_{index}_url"] = url_entry
            url_entry.bind("<KeyRelease>", self.schedule_alias_refresh)
            url_entry.bind("<FocusOut>", self.schedule_alias_refresh)

        refresh_btn = ctk.CTkButton(
            card,
            text="重新读取下拉",
            command=self.refresh_alias_options,
            width=120,
            height=28,
            font=ctk.CTkFont(size=11),
            fg_color=("gray88", "gray24"),
            hover_color=("gray80", "gray30"),
            text_color=("gray25", "gray85"),
            border_width=1,
            border_color=("gray78", "gray38"),
        )
        refresh_btn.pack(anchor="w", padx=15, pady=(0, 12))

        save_btn = ModernButton(
            parent,
            text="保存机器人别名",
            icon="💾",
            command=self.save_webhook_alias_settings,
            height=42,
            width=220,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#2CC985",
            hover_color="#239B6D",
        )
        save_btn.pack(anchor="w", padx=12, pady=(6, 12))

    def build_folder_tab(self, parent):
        card = ctk.CTkFrame(parent, fg_color=("gray95", "gray20"))
        card.pack(fill="x", padx=12, pady=10)

        card_title = ctk.CTkLabel(card, text="📁 文件夹检测配置 (最多3个)", font=ctk.CTkFont(size=18, weight="bold"))
        card_title.pack(pady=(15, 10), anchor="w", padx=15)

        tip = ctk.CTkLabel(
            card,
            text="每个监测项独立配置路径和推送机器人，适合把脚本产物或人工放入的文件自动推送。",
            font=ctk.CTkFont(size=12),
            text_color=("gray35", "gray70"),
        )
        tip.pack(anchor="w", padx=15, pady=(0, 10))

        monitor_config = self.load_folder_monitor_config()
        for index in range(1, self.folder_slot_count + 1):
            self.create_single_folder_monitor(card, index, monitor_config.get(f"folder_{index}", {}))

        save_btn = ModernButton(
            parent,
            text="保存文件夹检测",
            icon="💾",
            command=self.save_folder_settings,
            height=42,
            width=220,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#2CC985",
            hover_color="#239B6D",
        )
        save_btn.pack(anchor="w", padx=12, pady=(6, 12))

    def build_path_tab(self, parent):
        self.create_config_card(
            parent,
            "💾 文件路径配置",
            [
                ("下载目录", "downloads", str(self.config.download_dir)),
                ("状态文件", "state", str(self.config.state_file)),
            ],
        )

        save_btn = ModernButton(
            parent,
            text="保存路径配置",
            icon="💾",
            command=self.save_path_settings,
            height=42,
            width=220,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#2CC985",
            hover_color="#239B6D",
        )
        save_btn.pack(anchor="w", padx=12, pady=(6, 12))

    def build_ui_tab(self, parent):
        card = ctk.CTkFrame(parent, fg_color=("gray95", "gray20"))
        card.pack(fill="x", padx=12, pady=10)

        card_title = ctk.CTkLabel(card, text="🖥️ 界面设置", font=ctk.CTkFont(size=18, weight="bold"))
        card_title.pack(pady=(15, 10), anchor="w", padx=15)

        tip = ctk.CTkLabel(
            card,
            text="窗口大小会在关闭程序时自动更新；其余界面项保存后下次启动生效。",
            font=ctk.CTkFont(size=12),
            text_color=("gray35", "gray70"),
        )
        tip.pack(anchor="w", padx=15, pady=(0, 8))

        fields = ctk.CTkFrame(card, fg_color="transparent")
        fields.pack(fill="x", padx=15, pady=(0, 15))

        row1 = ctk.CTkFrame(fields, fg_color="transparent")
        row1.pack(fill="x", pady=6)
        ctk.CTkLabel(row1, text="界面模式", width=120, anchor="w").pack(side="left", padx=(0, 8))
        appearance_var = tk.StringVar(value=self.config.ui_appearance)
        ctk.CTkOptionMenu(row1, values=["dark", "light", "system"], variable=appearance_var, width=140).pack(
            side="left"
        )
        self.ui_vars["appearance"] = appearance_var

        ctk.CTkLabel(row1, text="颜色主题", width=120, anchor="w").pack(side="left", padx=(16, 8))
        theme_label = self._label_for_value(self.THEME_LABEL_TO_VALUE, self.config.ui_color_theme, "标准蓝色")
        theme_var = tk.StringVar(value=theme_label)
        ctk.CTkOptionMenu(
            row1,
            values=list(self.THEME_LABEL_TO_VALUE.keys()),
            variable=theme_var,
            width=180,
        ).pack(side="left")
        self.ui_vars["color_theme"] = theme_var

        ctk.CTkLabel(row1, text="启动页", width=90, anchor="w").pack(side="left", padx=(16, 8))
        start_page_label = self._label_for_value(self.START_PAGE_LABEL_TO_VALUE, self.config.start_page, "邮件检测")
        start_page_var = tk.StringVar(value=start_page_label)
        ctk.CTkOptionMenu(
            row1,
            values=list(self.START_PAGE_LABEL_TO_VALUE.keys()),
            variable=start_page_var,
            width=150,
        ).pack(side="left")
        self.ui_vars["start_page"] = start_page_var

        row2 = ctk.CTkFrame(fields, fg_color="transparent")
        row2.pack(fill="x", pady=6)
        ctk.CTkLabel(row2, text="窗口宽度(px)", width=120, anchor="w").pack(side="left", padx=(0, 8))
        width_entry = ctk.CTkEntry(row2, width=140)
        width_entry.insert(0, str(self.config.window_width))
        width_entry.pack(side="left")
        self.entries["ui_window_width"] = width_entry

        ctk.CTkLabel(row2, text="窗口高度(px)", width=120, anchor="w").pack(side="left", padx=(16, 8))
        height_entry = ctk.CTkEntry(row2, width=140)
        height_entry.insert(0, str(self.config.window_height))
        height_entry.pack(side="left")
        self.entries["ui_window_height"] = height_entry

        ctk.CTkLabel(row2, text="侧栏宽度(px)", width=100, anchor="w").pack(side="left", padx=(16, 8))
        sidebar_entry = ctk.CTkEntry(row2, width=120)
        sidebar_entry.insert(0, str(self.config.sidebar_width))
        sidebar_entry.pack(side="left")
        self.entries["ui_sidebar_width"] = sidebar_entry

        row3 = ctk.CTkFrame(fields, fg_color="transparent")
        row3.pack(fill="x", pady=6)
        ctk.CTkLabel(row3, text="日志刷新(ms)", width=120, anchor="w").pack(side="left", padx=(0, 8))
        poll_entry = ctk.CTkEntry(row3, width=140)
        poll_entry.insert(0, str(self.config.ui_log_poll_ms))
        poll_entry.pack(side="left")
        self.entries["ui_log_poll_ms"] = poll_entry

        ctk.CTkLabel(row3, text="界面缩放", width=120, anchor="w").pack(side="left", padx=(16, 8))
        scale_entry = ctk.CTkEntry(row3, width=140, placeholder_text="例: 1.0")
        scale_entry.insert(0, str(self.config.ui_scale))
        scale_entry.pack(side="left")
        self.entries["ui_scale"] = scale_entry

        auto_scroll_var = tk.BooleanVar(value=bool(self.config.auto_scroll_log))
        auto_scroll_box = ctk.CTkCheckBox(row3, text="日志自动滚动", variable=auto_scroll_var)
        auto_scroll_box.pack(side="left", padx=(24, 0))
        self.ui_vars["auto_scroll_log"] = auto_scroll_var

        row4 = ctk.CTkFrame(fields, fg_color="transparent")
        row4.pack(fill="x", pady=6)
        ctk.CTkLabel(row4, text="处理程序超时(s)", width=120, anchor="w").pack(side="left", padx=(0, 8))
        timeout_entry = ctk.CTkEntry(row4, width=140, placeholder_text="建议 300")
        timeout_entry.insert(0, str(self.config.script_timeout_seconds))
        timeout_entry.pack(side="left")
        self.entries["ui_script_timeout_seconds"] = timeout_entry

        ctk.CTkLabel(
            row4,
            text="脚本/.exe 处理单个附件的最长等待时间",
            font=ctk.CTkFont(size=11),
            text_color=("gray35", "gray70"),
        ).pack(side="left", padx=(16, 0))

        save_btn = ModernButton(
            parent,
            text="保存界面设置",
            icon="💾",
            command=self.save_ui_settings,
            height=42,
            width=220,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#2CC985",
            hover_color="#239B6D",
        )
        save_btn.pack(anchor="w", padx=12, pady=(6, 12))

    def create_single_folder_monitor(self, parent, index: int, config: Dict):
        monitor_card = ctk.CTkFrame(parent, fg_color=("gray92", "gray18"))
        monitor_card.pack(fill="x", padx=15, pady=8)

        header = ctk.CTkFrame(monitor_card, fg_color="transparent")
        header.pack(fill="x", padx=14, pady=(12, 8))

        checkbox = ctk.CTkCheckBox(header, text="", width=20, onvalue=True, offvalue=False)
        if config.get("enabled", False):
            checkbox.select()
        checkbox.pack(side="left", padx=(0, 8))
        self.folder_checkboxes[f"folder_{index}_enabled"] = checkbox

        ctk.CTkLabel(header, text=f"检测 {index}", font=ctk.CTkFont(size=15, weight="bold")).pack(side="left")
        ctk.CTkLabel(
            header,
            text="勾选即启用该监测项",
            font=ctk.CTkFont(size=11),
            text_color=("gray35", "gray70"),
        ).pack(side="left", padx=(10, 0))

        body = ctk.CTkFrame(monitor_card, fg_color="transparent")
        body.pack(fill="x", padx=14, pady=(0, 12))

        config_row = ctk.CTkFrame(body, fg_color="transparent")
        config_row.pack(fill="x")

        path_label = ctk.CTkLabel(config_row, text="监测路径", width=64, anchor="w", font=ctk.CTkFont(size=12))
        path_label.pack(side="left")

        path_entry = ctk.CTkEntry(
            config_row,
            placeholder_text=config.get("path", ""),
            height=32,
            font=ctk.CTkFont(size=12),
        )
        path_entry.insert(0, config.get("path", ""))
        path_entry.pack(side="left", fill="x", expand=True, padx=(4, 6))
        self.entries[f"folder_{index}_path"] = path_entry

        browse_btn = ctk.CTkButton(
            config_row,
            text="浏览",
            width=60,
            height=32,
            font=ctk.CTkFont(size=11),
            command=lambda entry=path_entry: self.browse_folder(entry),
            fg_color=("gray70", "gray30"),
            hover_color=("gray50", "gray40"),
        )
        browse_btn.pack(side="left", padx=(0, 8))

        alias_label = ctk.CTkLabel(config_row, text="推送机器人", width=70, anchor="w", font=ctk.CTkFont(size=12))
        alias_label.pack(side="left")

        option_values = self.get_alias_option_values()
        selected_alias = config.get("webhook_alias", "")
        if not selected_alias:
            selected_alias = self.find_alias_by_url(config.get("webhook_url", ""))
        if selected_alias not in option_values:
            selected_alias = self.NO_ALIAS_LABEL

        alias_var = tk.StringVar(value=selected_alias)
        alias_menu = ctk.CTkOptionMenu(
            config_row,
            values=option_values,
            variable=alias_var,
            width=170,
            height=32,
            font=ctk.CTkFont(size=12),
        )
        alias_menu.pack(side="left", padx=(4, 0))

        folder_alias_key = f"folder_{index}_alias"
        self.folder_alias_vars[folder_alias_key] = alias_var
        self.folder_alias_menus[folder_alias_key] = alias_menu

    def load_folder_monitor_config(self) -> Dict:
        config_file = Path("settings/folder_monitor_config.json")
        if config_file.exists():
            try:
                return json.loads(config_file.read_text(encoding="utf-8"))
            except Exception:
                return {}
        return {}

    def browse_folder(self, entry_widget):
        root = tk.Tk()
        root.withdraw()
        root.wm_attributes("-topmost", 1)
        folder_path = filedialog.askdirectory(title="选择要检测的文件夹", mustexist=False)
        root.destroy()
        if folder_path:
            entry_widget.delete(0, "end")
            entry_widget.insert(0, folder_path)

    def browse_python_script(self, entry_widget):
        root = tk.Tk()
        root.withdraw()
        root.wm_attributes("-topmost", 1)
        file_path = filedialog.askopenfilename(
            title="选择处理程序",
            filetypes=[("处理程序", "*.py;*.exe"), ("Python 脚本", "*.py"), ("EXE 程序", "*.exe"), ("所有文件", "*.*")],
        )
        root.destroy()
        if file_path:
            entry_widget.delete(0, "end")
            entry_widget.insert(0, file_path)

    def collect_email_config_values(self) -> dict:
        cfg = load_config()
        return {
            "IMAP_HOST": cfg.imap_host,
            "IMAP_PORT": str(cfg.imap_port),
            "EMAIL_USERNAME": cfg.email_username,
            "EMAIL_PASSWORD": cfg.email_password,
            "IMAP_MAILBOX": cfg.imap_mailbox,
            "SUBJECT_KEYWORDS": ",".join(cfg.subject_keywords),
            "POLL_INTERVAL_SECONDS": str(cfg.poll_interval_seconds),
            "MAX_ATTACHMENT_SIZE_MB": str(cfg.max_attachment_size_mb),
            "WEBHOOK_SEND_URL": cfg.webhook_send_url,
            "WEBHOOK_SEND_ALIAS": "",
            "DOWNLOAD_DIR": self.entries["downloads"].get().strip(),
            "STATE_FILE": self.entries["state"].get().strip(),
        }

    def _sync_legacy_app_config_with_mailbox(self, mailbox: Optional[dict], *, webhook_alias: str = "", webhook_url: str = ""):
        cfg = load_config()
        mailbox = mailbox or {}
        upsert_env_file(
            Path("settings/app_config.json"),
            {
                "IMAP_HOST": str(mailbox.get("host", cfg.imap_host)),
                "IMAP_PORT": str(mailbox.get("port", cfg.imap_port)),
                "EMAIL_USERNAME": str(mailbox.get("username", cfg.email_username)),
                "EMAIL_PASSWORD": str(mailbox.get("password", cfg.email_password)),
                "IMAP_MAILBOX": str(mailbox.get("mailbox", cfg.imap_mailbox)),
                "SUBJECT_KEYWORDS": ",".join(cfg.subject_keywords),
                "POLL_INTERVAL_SECONDS": str(cfg.poll_interval_seconds),
                "MAX_ATTACHMENT_SIZE_MB": str(cfg.max_attachment_size_mb),
                "WEBHOOK_SEND_URL": webhook_url or cfg.webhook_send_url,
                "WEBHOOK_SEND_ALIAS": webhook_alias or self.alias_config.get("email_alias", ""),
                "DOWNLOAD_DIR": str(cfg.download_dir),
                "STATE_FILE": str(cfg.state_file),
            },
        )

    def save_email_settings(self):
        try:
            mailboxes, validation_errors = self.collect_mailboxes_from_inputs(strict=True)
            if validation_errors:
                messagebox.showerror("验证失败", "\n".join(f"• {error}" for error in validation_errors))
                return
            if not mailboxes:
                messagebox.showerror("验证失败", "请至少配置一个 IMAP 邮箱")
                return

            save_mailbox_configs(mailboxes)
            self.mailbox_config = {"mailboxes": mailboxes}
            self.refresh_mailbox_options()
            self._sync_legacy_app_config_with_mailbox(mailboxes[0])
            self._notify_config_changed()
            messagebox.showinfo("保存成功", "邮箱配置已保存")
        except Exception as exc:
            messagebox.showerror("错误", f"保存邮箱配置失败:\n{exc}")

    def save_webhook_alias_settings(self):
        try:
            aliases, alias_errors = self.collect_aliases_from_inputs(strict=True)
            if alias_errors:
                messagebox.showerror("验证失败", "\n".join(f"• {error}" for error in alias_errors))
                return
            if not aliases:
                messagebox.showerror("验证失败", "请至少配置一个机器人别名")
                return

            previous_alias = self.alias_config.get("email_alias", "")
            email_alias = previous_alias if previous_alias in aliases else next(iter(aliases.keys()), "")

            save_webhook_aliases(aliases, email_alias)
            self.alias_config = {"aliases": aliases, "email_alias": email_alias}
            self.refresh_alias_options()
            self._notify_config_changed()
            messagebox.showinfo("保存成功", "机器人别名配置已保存")
        except Exception as exc:
            messagebox.showerror("错误", f"保存别名配置失败:\n{exc}")

    def _collect_single_rule_payload(self, index: int) -> tuple[Optional[dict], list[str], bool]:
        aliases = self.alias_config.get("aliases", {})
        mailboxes = self.mailbox_config.get("mailboxes", [])
        mailbox_alias_map = {
            str(item.get("alias", "")).strip(): item
            for item in mailboxes
            if str(item.get("alias", "")).strip()
        }
        default_webhook_url = (self.config.webhook_send_url if self.config else "").strip()

        keyword_entry = self.entries.get(f"rule_{index}_keyword")
        types_entry = self.entries.get(f"rule_{index}_types")
        names_entry = self.entries.get(f"rule_{index}_name_keywords")
        interval_entry = self.entries.get(f"rule_{index}_interval")
        schedule_entry = self.entries.get(f"rule_{index}_schedule_time")
        max_size_entry = self.entries.get(f"rule_{index}_max_size")
        script_entry = self.entries.get(f"rule_{index}_script_path")
        output_dir_entry = self.entries.get(f"rule_{index}_script_output_dir")
        enabled_checkbox = self.rule_checkboxes.get(f"rule_{index}_enabled")
        alias_var = self.rule_alias_vars.get(f"rule_{index}_alias")
        mailbox_alias_var = self.rule_mailbox_vars.get(f"rule_{index}_mailbox_alias")
        trigger_mode_var = self.ui_vars.get(f"rule_{index}_trigger_mode")

        enabled = enabled_checkbox.get() if enabled_checkbox else False
        keyword = keyword_entry.get().strip() if keyword_entry else ""
        raw_types = types_entry.get().strip() if types_entry else ""
        raw_name_keywords = names_entry.get().strip() if names_entry else ""
        raw_interval = interval_entry.get().strip() if interval_entry else ""
        raw_schedule_time = schedule_entry.get().strip() if schedule_entry else ""
        raw_max_size = max_size_entry.get().strip() if max_size_entry else ""
        script_path = script_entry.get().strip() if script_entry else ""
        script_output_dir = output_dir_entry.get().strip() if output_dir_entry else ""
        trigger_mode_label = trigger_mode_var.get().strip() if trigger_mode_var else "周期检测"
        trigger_mode = self._value_for_label(self.TRIGGER_LABEL_TO_VALUE, trigger_mode_label, "periodic")
        parsed_name_keywords = parse_filename_keywords_input(raw_name_keywords)
        selected_mailbox_alias = mailbox_alias_var.get().strip() if mailbox_alias_var else ""
        if selected_mailbox_alias == self.NO_MAILBOX_LABEL:
            selected_mailbox_alias = ""
        selected_alias = alias_var.get().strip() if alias_var else ""
        if selected_alias == self.NO_ALIAS_LABEL:
            selected_alias = ""
        if enabled and not selected_mailbox_alias and mailbox_alias_map:
            selected_mailbox_alias = next(iter(mailbox_alias_map.keys()), "")
        if enabled and not selected_alias and aliases:
            selected_alias = next(iter(aliases.keys()), "")
        webhook_url = aliases.get(selected_alias, "").strip()
        if enabled and not webhook_url:
            webhook_url = default_webhook_url

        is_empty = (
            not keyword
            and not raw_types
            and not raw_name_keywords
            and not raw_interval
            and not raw_schedule_time
            and not raw_max_size
            and not script_path
            and not script_output_dir
            and not selected_mailbox_alias
            and not selected_alias
            and not enabled
        )
        if is_empty:
            return None, [], True

        errors = []
        if not keyword or not raw_types:
            state_text = "已启用" if enabled else "未启用"
            errors.append(f"规则{index}({state_text}): 主题关键字和附件格式必须同时填写")
            return None, errors, False

        parsed_types = parse_types_input(raw_types)
        if any(separator in raw_types for separator in [",", ";", "|"]):
            errors.append(f"规则{index}: 附件格式只允许填写一个值，例如 png")
        elif not parsed_types:
            errors.append(f"规则{index}: 附件格式无效，示例 png")

        if any(separator in raw_name_keywords for separator in [",", ";", "|"]):
            errors.append(f"规则{index}: 附件文件名关键字只允许填写一个值")

        try:
            parsed_max_size = int(raw_max_size)
            if parsed_max_size <= 0:
                raise ValueError
        except ValueError:
            errors.append(f"规则{index}: 最大附件(MB)必须是大于0的整数")
            parsed_max_size = None

        if trigger_mode == "timed":
            parsed_schedule_time = raw_schedule_time
            try:
                hour_text, minute_text = parsed_schedule_time.split(":", 1)
                hour = int(hour_text)
                minute = int(minute_text)
                if not (0 <= hour <= 23 and 0 <= minute <= 59):
                    raise ValueError
            except Exception:
                errors.append(f"规则{index}: 定时检测请输入有效时间，格式如 08:30")
            parsed_interval_seconds = self.interval_minutes_to_seconds(
                int(raw_interval) if raw_interval.isdigit() and int(raw_interval) > 0 else 1
            )
        else:
            try:
                parsed_interval_minutes = int(raw_interval)
                if parsed_interval_minutes <= 0:
                    raise ValueError
            except ValueError:
                errors.append(f"规则{index}: 周期检测时，轮询间隔(min)必须是大于0的整数")
                parsed_interval_minutes = 1
            parsed_interval_seconds = self.interval_minutes_to_seconds(parsed_interval_minutes)
            parsed_schedule_time = ""

        if enabled and not selected_mailbox_alias:
            errors.append(f"规则{index}: 已启用但未选择所属邮箱")
        elif enabled and selected_mailbox_alias not in mailbox_alias_map:
            errors.append(f"规则{index}: 所属邮箱无效 - {selected_mailbox_alias}")

        use_script = bool(script_path)
        if use_script:
            if not Path(script_path).exists():
                errors.append(f"规则{index}: 脚本文件不存在 - {script_path}")
            elif Path(script_path).suffix.lower() not in {'.py', '.exe'}:
                errors.append(f"规则{index}: 仅支持选择 .py 或 .exe 处理程序")
            if not script_output_dir:
                errors.append(f"规则{index}: 选择脚本后必须填写输出目录")

        if enabled and not webhook_url:
            errors.append(f"规则{index}: 已启用但未配置可用的推送地址")

        if errors:
            return None, errors, False

        return {
            "enabled": bool(enabled),
            "keyword": keyword,
            "types": parsed_types,
            "filename_keywords": parsed_name_keywords,
            "mailbox_alias": selected_mailbox_alias,
            "webhook_alias": selected_alias,
            "webhook_url": webhook_url,
            "script_path": script_path,
            "script_output_dir": script_output_dir,
            "trigger_mode": trigger_mode,
            "schedule_time": parsed_schedule_time,
            "poll_interval_seconds": parsed_interval_seconds,
            "max_attachment_size_mb": parsed_max_size,
        }, [], False

    def save_single_subject_rule(self, index: int):
        try:
            rule_payload, rule_errors, is_empty = self._collect_single_rule_payload(index)
            if rule_errors:
                messagebox.showerror("验证失败", "\n".join(f"• {error}" for error in rule_errors))
                return

            current_rules = list(load_subject_attachment_rules().get("rules", []))
            while len(current_rules) < index:
                current_rules.append({})

            if is_empty:
                current_rules[index - 1] = {}
            else:
                current_rules[index - 1] = rule_payload

            save_subject_attachment_rules(current_rules)
            self.subject_rules_payload = load_subject_attachment_rules()
            self.refresh_rule_summaries_from_saved_payload()

            saved_rules = self.subject_rules_payload.get("rules", [])
            default_webhook_alias = self.alias_config.get("email_alias", "")
            mailboxes = self.mailbox_config.get("mailboxes", [])
            mailbox_alias_map = {
                str(item.get("alias", "")).strip(): item
                for item in mailboxes
                if str(item.get("alias", "")).strip()
            }
            cfg = load_config()
            default_rule = next((item for item in saved_rules if item.get("enabled")), None) or (
                saved_rules[0] if saved_rules else None
            )
            saved_poll = str(default_rule.get("poll_interval_seconds")) if default_rule else str(cfg.poll_interval_seconds)
            saved_max = str(default_rule.get("max_attachment_size_mb")) if default_rule else str(cfg.max_attachment_size_mb)
            default_send_url = (default_rule.get("webhook_url", "").strip() if default_rule else "") or cfg.webhook_send_url
            default_send_alias = (
                (default_rule.get("webhook_alias", "").strip() if default_rule else "")
                or default_webhook_alias
            )
            default_mailbox_alias = (default_rule.get("mailbox_alias", "").strip() if default_rule else "")
            default_mailbox = mailbox_alias_map.get(default_mailbox_alias) or (mailboxes[0] if mailboxes else None)
            self._sync_legacy_app_config_with_mailbox(
                default_mailbox,
                webhook_alias=default_send_alias,
                webhook_url=default_send_url,
            )
            upsert_env_file(
                Path("settings/app_config.json"),
                {
                    "POLL_INTERVAL_SECONDS": saved_poll,
                    "MAX_ATTACHMENT_SIZE_MB": saved_max,
                },
            )
            self._notify_config_changed()

            if is_empty:
                messagebox.showinfo("保存成功", f"规则{index}已清空并保存")
            else:
                messagebox.showinfo("保存成功", f"规则{index}已保存，顶部摘要已更新为最新已保存内容")
        except Exception as exc:
            messagebox.showerror("错误", f"保存规则{index}失败:\n{exc}")

    def save_folder_settings(self):
        try:
            aliases, _ = self.collect_aliases_from_inputs(strict=False)
            if not aliases:
                aliases = self.alias_config.get("aliases", {})

            folder_monitor_config = {}
            folder_config_warnings = []
            folder_config_errors = []

            for index in range(1, self.folder_slot_count + 1):
                folder_key = f"folder_{index}"
                path_entry = self.entries.get(f"{folder_key}_path")
                enabled_checkbox = self.folder_checkboxes.get(f"{folder_key}_enabled")
                alias_var = self.folder_alias_vars.get(f"{folder_key}_alias")

                path = path_entry.get() if path_entry else ""
                selected_alias = alias_var.get().strip() if alias_var else ""
                if selected_alias == self.NO_ALIAS_LABEL:
                    selected_alias = ""
                webhook_url = aliases.get(selected_alias, "")
                enabled = enabled_checkbox.get() if enabled_checkbox else False

                if enabled or path or selected_alias:
                    if enabled:
                        if not path:
                            folder_config_warnings.append(f"检测{index}: 已启用但未设置文件夹路径")
                        elif not Path(path).exists():
                            folder_config_warnings.append(f"检测{index}: 文件夹不存在 - {path}")

                        if not selected_alias:
                            folder_config_errors.append(f"检测{index}: 已启用但未选择推送别名")
                        elif not webhook_url:
                            folder_config_errors.append(f"检测{index}: 推送别名无效 - {selected_alias}")

                    folder_monitor_config[folder_key] = {
                        "path": path,
                        "webhook_alias": selected_alias,
                        "webhook_url": webhook_url,
                        "enabled": enabled,
                    }

            if folder_config_errors:
                messagebox.showerror("验证失败", "\n".join(f"• {error}" for error in folder_config_errors))
                return

            monitor_config_file = Path("settings/folder_monitor_config.json")
            monitor_config_file.parent.mkdir(parents=True, exist_ok=True)
            temp_file = monitor_config_file.with_suffix(monitor_config_file.suffix + ".tmp")
            temp_file.write_text(
                json.dumps(folder_monitor_config, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )
            temp_file.replace(monitor_config_file)

            enabled_count = sum(1 for cfg in folder_monitor_config.values() if cfg.get("enabled", False))
            self._notify_config_changed()
            success_message = f"✅ 文件夹检测配置已保存\n\n已启用 {enabled_count} 个文件夹检测"
            if folder_config_warnings:
                success_message += "\n\n⚠️ 警告：\n" + "\n".join(f"• {warning}" for warning in folder_config_warnings)
            messagebox.showinfo("保存成功", success_message)
        except Exception as exc:
            messagebox.showerror("错误", f"保存文件夹检测配置失败:\n{exc}")

    def save_path_settings(self):
        try:
            cfg = load_config()
            config_values = {
                "IMAP_HOST": cfg.imap_host,
                "IMAP_PORT": str(cfg.imap_port),
                "EMAIL_USERNAME": cfg.email_username,
                "EMAIL_PASSWORD": cfg.email_password,
                "IMAP_MAILBOX": cfg.imap_mailbox,
                "SUBJECT_KEYWORDS": ",".join(cfg.subject_keywords),
                "POLL_INTERVAL_SECONDS": str(cfg.poll_interval_seconds),
                "MAX_ATTACHMENT_SIZE_MB": str(cfg.max_attachment_size_mb),
                "WEBHOOK_SEND_URL": cfg.webhook_send_url,
                "WEBHOOK_SEND_ALIAS": self.alias_config.get("email_alias", ""),
                "DOWNLOAD_DIR": self.entries["downloads"].get().strip(),
                "STATE_FILE": self.entries["state"].get().strip(),
            }
            validation_errors = validate_config_values(config_values)
            if validation_errors:
                messagebox.showerror("验证失败", "\n".join(f"• {error}" for error in validation_errors))
                return

            upsert_env_file(Path("settings/app_config.json"), config_values)
            self._notify_config_changed()
            messagebox.showinfo("保存成功", "路径配置已保存")
        except Exception as exc:
            messagebox.showerror("错误", f"保存路径配置失败:\n{exc}")

    def save_ui_settings(self):
        try:
            appearance = self.ui_vars.get("appearance").get().strip()
            start_page_label = self.ui_vars.get("start_page").get().strip()
            auto_scroll = bool(self.ui_vars.get("auto_scroll_log").get())
            theme_label = self.ui_vars.get("color_theme").get().strip()
            start_page = self._value_for_label(self.START_PAGE_LABEL_TO_VALUE, start_page_label, "execute")
            color_theme = self._value_for_label(self.THEME_LABEL_TO_VALUE, theme_label, "blue")

            try:
                window_width = int(self.entries["ui_window_width"].get().strip())
                window_height = int(self.entries["ui_window_height"].get().strip())
                sidebar_width = int(self.entries["ui_sidebar_width"].get().strip())
                log_poll_ms = int(self.entries["ui_log_poll_ms"].get().strip())
                script_timeout_seconds = int(self.entries["ui_script_timeout_seconds"].get().strip())
                ui_scale = float(self.entries["ui_scale"].get().strip())
                if (
                    window_width <= 0
                    or window_height <= 0
                    or sidebar_width <= 0
                    or log_poll_ms <= 0
                    or script_timeout_seconds <= 0
                    or ui_scale <= 0
                ):
                    raise ValueError
            except ValueError:
                messagebox.showerror(
                    "验证失败",
                    "窗口宽高/侧栏宽度/日志刷新/处理程序超时必须是大于0的整数，界面缩放必须是大于0的数字",
                )
                return

            if appearance not in {"dark", "light", "system"}:
                messagebox.showerror("验证失败", "界面模式仅支持 dark/light/system")
                return
            if start_page not in {"execute", "folder", "bot_test", "settings", "about"}:
                messagebox.showerror("验证失败", "启动页无效")
                return

            upsert_env_file(
                Path("settings/app_config.json"),
                {
                    "UI_APPEARANCE": appearance,
                    "UI_COLOR_THEME": color_theme,
                    "START_PAGE": start_page,
                    "WINDOW_WIDTH": str(window_width),
                    "WINDOW_HEIGHT": str(window_height),
                    "SIDEBAR_WIDTH": str(sidebar_width),
                    "UI_LOG_POLL_MS": str(log_poll_ms),
                    "AUTO_SCROLL_LOG": "true" if auto_scroll else "false",
                    "UI_SCALE": str(ui_scale),
                    "SCRIPT_TIMEOUT_SECONDS": str(script_timeout_seconds),
                },
            )
            self._notify_config_changed()
            messagebox.showinfo("保存成功", "界面设置已保存（重启程序后全部生效）")
        except Exception as exc:
            messagebox.showerror("错误", f"保存界面设置失败:\n{exc}")
