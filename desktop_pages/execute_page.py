import threading
import time
from typing import Dict
from datetime import datetime
import tkinter.messagebox as messagebox
from pathlib import Path
from tkinter import filedialog

import customtkinter as ctk

from mail_forwarder import MailProcessingService, load_config
from mail_forwarder.subject_attachment_rules import list_enabled_rules_with_slots

from .common import LogHandler, ModernButton


class ExecutePage(ctk.CTkFrame):
    def __init__(self, master, log_handler: LogHandler, auto_scroll_log: bool = True, **kwargs):
        super().__init__(master, **kwargs)
        self.log_handler = log_handler
        self.auto_scroll_log = auto_scroll_log
        self.rule_configs: dict[str, dict] = {}
        self.rule_cards: dict[str, dict] = {}
        self.rule_runtimes: dict[str, dict] = {}
        self.rule_log_texts: dict[str, ctk.CTkTextbox] = {}
        self._rule_test_running: set[str] = set()

        self.monitor_list = None
        self.log_container = None
        self.log_tabview = None
        self.overview_log_text = None
        self.rule_count_label = None
        self.running_count_label = None
        self._tab_names: dict[str, str] = {}

        self.setup_ui()
        self._create_log_tabview()
        self.sync_rule_views(initial=True)
        self.log_handler.add_callback(self.append_log)

    def setup_ui(self):
        control_bar = ctk.CTkFrame(self, height=100)
        control_bar.pack(fill="x", padx=25, pady=(20, 10))

        left_panel = ctk.CTkFrame(control_bar, fg_color="transparent")
        left_panel.pack(side="left", fill="both", expand=True)

        title = ctk.CTkLabel(left_panel, text="📧 邮件检测", font=ctk.CTkFont(size=26, weight="bold"))
        title.pack(pady=(5, 5), anchor="w")

        self.status_label = ctk.CTkLabel(
            left_panel,
            text="⏸️ 待机中",
            font=ctk.CTkFont(size=16),
            text_color="gray",
        )
        self.status_label.pack(anchor="w")

        meta_row = ctk.CTkFrame(left_panel, fg_color="transparent")
        meta_row.pack(anchor="w", pady=(4, 0))

        self.rule_count_label = ctk.CTkLabel(
            meta_row,
            text="启用规则: 0",
            font=ctk.CTkFont(size=12),
            text_color=("gray35", "gray70"),
        )
        self.rule_count_label.pack(side="left")

        self.running_count_label = ctk.CTkLabel(
            meta_row,
            text="",
            font=ctk.CTkFont(size=12),
            text_color=("#2E7D32", "#81C784"),
        )
        self.running_count_label.pack(side="left", padx=(12, 0))

        right_panel = ctk.CTkFrame(control_bar, fg_color="transparent")
        right_panel.pack(side="right", padx=10)

        self.start_btn = ModernButton(
            right_panel,
            text="全部启动",
            icon="▶️",
            command=self.start_worker,
            width=120,
            height=45,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color="#2CC985",
            hover_color="#239B6D",
        )
        self.start_btn.pack(side="left", padx=5)

        self.stop_btn = ModernButton(
            right_panel,
            text="全部停止",
            icon="⏹️",
            command=self.stop_worker,
            width=120,
            height=45,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color="#EF5350",
            hover_color="#C62828",
        )
        self.stop_btn.pack(side="left", padx=5)

        self.test_btn = ModernButton(
            right_panel,
            text="全部测试",
            icon="🧪",
            command=self.test_once,
            width=120,
            height=45,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color="#42A5F5",
            hover_color="#1976D2",
        )
        self.test_btn.pack(side="left", padx=5)

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=25, pady=(0, 20))

        self.monitor_list = ctk.CTkScrollableFrame(body, fg_color=("gray96", "gray17"), width=430)
        self.monitor_list.pack(side="left", fill="y", padx=(0, 16))

        self.log_container = ctk.CTkFrame(body)
        self.log_container.pack(side="right", fill="both", expand=True)

    @staticmethod
    def format_wait_text(seconds: int) -> str:
        if seconds % 60 == 0 and seconds >= 60:
            minutes = seconds // 60
            return f"{minutes} 分钟"
        return f"{seconds} 秒"

    @staticmethod
    def _build_rule_id(slot_index: int) -> str:
        return f"rule_slot_{slot_index}"

    @staticmethod
    def _slot_from_rule_id(rule_id: str) -> int:
        try:
            return int(str(rule_id).rsplit("_", 1)[-1])
        except (TypeError, ValueError):
            return 0

    @staticmethod
    def _display_value(value: str, fallback: str) -> str:
        text = str(value).strip()
        return text or fallback

    def _rule_summary_text(self, rule: dict) -> str:
        mailbox_alias = self._display_value(rule.get("mailbox_alias", ""), "未选邮箱")
        webhook_alias = self._display_value(rule.get("webhook_alias", ""), "未选机器人")
        interval_seconds = int(rule.get("poll_interval_seconds", 60) or 60)
        trigger_mode = str(rule.get("trigger_mode", "periodic")).strip() or "periodic"
        schedule_time = str(rule.get("schedule_time", "")).strip()
        mode = "脚本处理" if str(rule.get("script_path", "")).strip() else "直接推送"
        trigger_text = (
            f"定时: {schedule_time or '--:--'}"
            if trigger_mode == "timed"
            else f"周期: {self.format_wait_text(interval_seconds)}"
        )
        return (
            f"邮箱: {mailbox_alias}\n"
            f"机器人: {webhook_alias}\n"
            f"{trigger_text}\n"
            f"模式: {mode}"
        )

    def _create_log_tabview(self):
        for child in self.log_container.winfo_children():
            child.destroy()

        log_header = ctk.CTkFrame(self.log_container, height=50)
        log_header.pack(fill="x")

        log_title = ctk.CTkLabel(log_header, text="📋 邮件检测日志", font=ctk.CTkFont(size=18, weight="bold"))
        log_title.pack(side="left", padx=10, pady=10)

        clear_btn = ctk.CTkButton(
            log_header,
            text="🗑️ 清空当前",
            command=self.clear_logs,
            width=100,
            height=35,
            font=ctk.CTkFont(size=12),
        )
        clear_btn.pack(side="right", padx=10, pady=7)

        export_btn = ctk.CTkButton(
            log_header,
            text="导出当前",
            command=self.export_current_logs,
            width=100,
            height=35,
            font=ctk.CTkFont(size=12),
        )
        export_btn.pack(side="right", padx=(0, 8), pady=7)

        self.log_tabview = ctk.CTkTabview(self.log_container)
        self.log_tabview.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        overview_tab = self.log_tabview.add("总览")
        self.overview_log_text = ctk.CTkTextbox(
            overview_tab,
            font=ctk.CTkFont(family="Consolas", size=12),
            wrap="word",
            fg_color=("gray98", "gray15"),
        )
        self.overview_log_text.pack(fill="both", expand=True, padx=6, pady=6)

    def _get_enabled_rule_map(self) -> dict[str, dict]:
        result: dict[str, dict] = {}
        for slot_index, rule in list_enabled_rules_with_slots():
            result[self._build_rule_id(slot_index)] = rule
        return result

    def _ensure_rule_card(self, rule_id: str, rule: dict):
        if rule_id in self.rule_cards:
            self._update_rule_card_content(rule_id, rule)
            return

        slot_index = self._slot_from_rule_id(rule_id)
        card = ctk.CTkFrame(self.monitor_list, fg_color=("gray95", "gray20"))
        card.pack(fill="x", padx=6, pady=8)

        header = ctk.CTkFrame(card, fg_color="transparent")
        header.pack(fill="x", padx=12, pady=(12, 8))

        title_label = ctk.CTkLabel(
            header,
            text=self._display_value(rule.get("keyword", ""), f"规则 {slot_index}"),
            font=ctk.CTkFont(size=16, weight="bold"),
        )
        title_label.pack(side="left")

        slot_label = ctk.CTkLabel(
            header,
            text=f"槽位 {slot_index}",
            font=ctk.CTkFont(size=11),
            text_color=("gray35", "gray70"),
        )
        slot_label.pack(side="left", padx=(10, 0))

        status_label = ctk.CTkLabel(
            header,
            text="⏸️ 待机中",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="gray",
        )
        status_label.pack(side="right")

        summary_label = ctk.CTkLabel(
            card,
            text=self._rule_summary_text(rule),
            justify="left",
            anchor="w",
            font=ctk.CTkFont(size=12),
        )
        summary_label.pack(fill="x", padx=12, pady=(0, 10))

        result_label = ctk.CTkLabel(
            card,
            text="最近结果: 暂无记录",
            justify="left",
            anchor="w",
            font=ctk.CTkFont(size=12),
            text_color=("gray35", "gray70"),
        )
        result_label.pack(fill="x", padx=12, pady=(0, 10))

        actions = ctk.CTkFrame(card, fg_color="transparent")
        actions.pack(fill="x", padx=12, pady=(0, 12))

        start_btn = ModernButton(
            actions,
            text="启动",
            icon="▶️",
            command=lambda target=rule_id: self.start_single_rule(target),
            width=86,
            height=34,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#2CC985",
            hover_color="#239B6D",
        )
        start_btn.pack(side="left", padx=(0, 8))

        stop_btn = ModernButton(
            actions,
            text="停止",
            icon="⏹️",
            command=lambda target=rule_id: self.stop_single_rule(target),
            width=86,
            height=34,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#EF5350",
            hover_color="#C62828",
            state="disabled",
        )
        stop_btn.pack(side="left", padx=(0, 8))

        test_btn = ctk.CTkButton(
            actions,
            text="测试",
            width=80,
            height=34,
            font=ctk.CTkFont(size=12),
            command=lambda target=rule_id: self.test_single_rule(target),
        )
        test_btn.pack(side="left", padx=(0, 8))

        tab_name = self._tab_names.get(rule_id) or f"规则{slot_index}"
        self._tab_names[rule_id] = tab_name
        text_widget = self.rule_log_texts.get(rule_id)
        if text_widget is None:
            tab = self.log_tabview.add(tab_name)
            text_widget = ctk.CTkTextbox(
                tab,
                font=ctk.CTkFont(family="Consolas", size=12),
                wrap="word",
                fg_color=("gray98", "gray15"),
            )
            text_widget.pack(fill="both", expand=True, padx=6, pady=6)

        log_btn = ctk.CTkButton(
            actions,
            text="查看日志",
            width=90,
            height=34,
            font=ctk.CTkFont(size=12),
            command=lambda name=tab_name: self.log_tabview.set(name),
        )
        log_btn.pack(side="left")

        self.rule_cards[rule_id] = {
            "frame": card,
            "title_label": title_label,
            "summary_label": summary_label,
            "result_label": result_label,
            "status_label": status_label,
            "start_btn": start_btn,
            "stop_btn": stop_btn,
            "test_btn": test_btn,
            "log_btn": log_btn,
            "slot_index": slot_index,
        }
        self.rule_log_texts[rule_id] = text_widget
        self.rule_runtimes[rule_id] = {
            "thread": None,
            "stop_event": threading.Event(),
            "is_running": False,
            "pending_remove": False,
            "last_refresh_marker": self._rule_refresh_marker(rule),
            "last_result": None,
        }

    def _update_rule_card_content(self, rule_id: str, rule: dict):
        card = self.rule_cards.get(rule_id, {})
        if card.get("title_label"):
            card["title_label"].configure(
                text=self._display_value(rule.get("keyword", ""), f"规则 {card.get('slot_index', 0)}")
            )
        if card.get("summary_label"):
            card["summary_label"].configure(text=self._rule_summary_text(rule))
        self._refresh_result_label(rule_id)

    def _remove_rule_card(self, rule_id: str):
        card = self.rule_cards.pop(rule_id, None)
        if card and card.get("frame"):
            card["frame"].destroy()
        self.rule_runtimes.pop(rule_id, None)
        self._rule_test_running.discard(rule_id)

    def _set_rule_pending_remove(self, rule_id: str, reason: str):
        runtime = self.rule_runtimes.get(rule_id)
        if not runtime:
            return
        runtime["pending_remove"] = True
        self.log_handler.warning(reason, source=rule_id)
        self.request_stop_rule(rule_id)

    def _rule_refresh_marker(self, rule: dict) -> tuple:
        return (
            str(rule.get("keyword", "")).strip(),
            str(rule.get("mailbox_alias", "")).strip(),
            tuple(str(item).strip() for item in rule.get("types", []) if str(item).strip()),
            tuple(str(item).strip() for item in rule.get("filename_keywords", []) if str(item).strip()),
            str(rule.get("webhook_alias", "")).strip(),
            str(rule.get("webhook_url", "")).strip(),
            str(rule.get("script_path", "")).strip(),
            str(rule.get("script_output_dir", "")).strip(),
            str(rule.get("trigger_mode", "")).strip(),
            str(rule.get("schedule_time", "")).strip(),
            int(rule.get("poll_interval_seconds", 0) or 0),
            int(rule.get("max_attachment_size_mb", 0) or 0),
        )

    @staticmethod
    def _seconds_until_daily_time(schedule_time: str) -> tuple[int, str]:
        try:
            hour_text, minute_text = schedule_time.split(":", 1)
            hour = int(hour_text)
            minute = int(minute_text)
        except Exception:
            raise RuntimeError(f"定时时刻无效: {schedule_time or '(空)'}")
        now = datetime.now()
        target = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
        if now > target:
            from datetime import timedelta

            target = target + timedelta(days=1)
        seconds = max(0, int((target - now).total_seconds()))
        return seconds, target.strftime("%Y-%m-%d %H:%M")

    def sync_rule_views(self, *, initial: bool = False):
        enabled_rule_map = self._get_enabled_rule_map()
        self.rule_count_label.configure(text=f"启用规则: {len(enabled_rule_map)}")

        for rule_id, rule in enabled_rule_map.items():
            self._ensure_rule_card(rule_id, rule)
            previous_marker = self.rule_runtimes[rule_id].get("last_refresh_marker")
            current_marker = self._rule_refresh_marker(rule)
            self.rule_configs[rule_id] = rule
            self._update_rule_card_content(rule_id, rule)
            self.rule_runtimes[rule_id]["last_refresh_marker"] = current_marker
            if not initial and previous_marker and previous_marker != current_marker:
                self.log_handler.info("规则配置已热刷新，新配置将在下一轮自动生效", source=rule_id)

        existing_ids = list(self.rule_cards.keys())
        for rule_id in existing_ids:
            if rule_id in enabled_rule_map:
                continue
            runtime = self.rule_runtimes.get(rule_id, {})
            if runtime.get("is_running") or (runtime.get("thread") and runtime["thread"].is_alive()):
                self.rule_configs.pop(rule_id, None)
                self._set_rule_pending_remove(rule_id, "规则已被禁用或移除，正在自动停止")
                card = self.rule_cards.get(rule_id, {})
                if card.get("summary_label"):
                    card["summary_label"].configure(text="该规则已在设置中禁用或移除，线程退出后自动清理。")
            else:
                self.rule_configs.pop(rule_id, None)
                self._remove_rule_card(rule_id)

        self.update_running_status()

    def append_log(self, log_entry: Dict[str, str]):
        icons = {"INFO": "ℹ️", "SUCCESS": "✅", "WARNING": "⚠️", "ERROR": "❌"}
        log_line = f"[{log_entry['time']}] {icons.get(log_entry['level'], '•')} {log_entry['message']}\n"

        if self.overview_log_text is not None:
            self.overview_log_text.insert("end", log_line)
            if self.auto_scroll_log:
                self.overview_log_text.see("end")

        source = str(log_entry.get("source", "global"))
        rule_log = self.rule_log_texts.get(source)
        if rule_log:
            rule_log.insert("end", log_line)
            if self.auto_scroll_log:
                rule_log.see("end")

    def clear_logs(self):
        if not self.log_tabview:
            return
        current_tab = self.log_tabview.get()
        if current_tab == "总览" and self.overview_log_text is not None:
            self.overview_log_text.delete("1.0", "end")
            return
        for rule_id, text_widget in self.rule_log_texts.items():
            if current_tab == self._tab_names.get(rule_id):
                text_widget.delete("1.0", "end")
                return

    def export_current_logs(self):
        if not self.log_tabview:
            return
        current_tab = self.log_tabview.get()
        content = ""
        if current_tab == "总览" and self.overview_log_text is not None:
            content = self.overview_log_text.get("1.0", "end").strip()
        else:
            for rule_id, text_widget in self.rule_log_texts.items():
                if current_tab == self._tab_names.get(rule_id):
                    content = text_widget.get("1.0", "end").strip()
                    break
        if not content:
            messagebox.showwarning("提示", "当前日志为空，暂无可导出的内容")
            return

        export_dir = Path("exports")
        export_dir.mkdir(parents=True, exist_ok=True)
        default_name = f"mail_{current_tab}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log".replace(" ", "_")
        file_path = filedialog.asksaveasfilename(
            title="导出当前日志",
            initialdir=str(export_dir.resolve()),
            initialfile=default_name,
            defaultextension=".log",
            filetypes=[("Log Files", "*.log"), ("Text Files", "*.txt"), ("All Files", "*.*")],
        )
        if not file_path:
            return
        Path(file_path).write_text(content + "\n", encoding="utf-8")
        messagebox.showinfo("导出成功", f"日志已导出到:\n{file_path}")

    def _set_rule_card_state(self, rule_id: str, *, running: bool, text: str, color: str):
        card = self.rule_cards.get(rule_id, {})
        if card.get("status_label"):
            card["status_label"].configure(text=text, text_color=color)
        if card.get("start_btn"):
            card["start_btn"].configure(state="disabled" if running else "normal")
        if card.get("stop_btn"):
            card["stop_btn"].configure(state="normal" if running else "disabled")
        if card.get("test_btn"):
            card["test_btn"].configure(state="disabled" if running else "normal")
        self.update_running_status()

    def _format_result_text(self, payload: dict | None) -> str:
        if not payload:
            return "最近结果: 暂无记录"
        timestamp = str(payload.get("time", "")).strip() or "--"
        status = str(payload.get("status", "")).strip() or "unknown"
        status_text_map = {
            "processed": "成功",
            "skipped": "跳过",
            "not_found": "未找到",
            "error": "失败",
        }
        status_text = status_text_map.get(status, status)
        uid_text = str(payload.get("uid", "")).strip() or "-"
        file_count = int(payload.get("file_count", 0) or 0)
        reason = str(payload.get("reason", "")).strip()
        base = f"最近结果: {status_text} | 时间: {timestamp}\nUID: {uid_text} | 附件数: {file_count}"
        if reason:
            return f"{base}\n说明: {reason}"
        return base

    def _refresh_result_label(self, rule_id: str):
        card = self.rule_cards.get(rule_id, {})
        result_label = card.get("result_label")
        runtime = self.rule_runtimes.get(rule_id, {})
        if result_label:
            payload = runtime.get("last_result")
            status = str((payload or {}).get("status", "")).strip()
            text_color = {
                "processed": ("#1B5E20", "#9BE7A7"),
                "skipped": ("#8D6E00", "#FFD54F"),
                "not_found": ("#546E7A", "#B0BEC5"),
                "error": ("#B71C1C", "#FF8A80"),
            }.get(status, ("gray35", "gray70"))
            result_label.configure(text=self._format_result_text(payload), text_color=text_color)

    def _remember_rule_result(self, rule_id: str, result):
        runtime = self.rule_runtimes.get(rule_id)
        if not runtime:
            return
        runtime["last_result"] = {
            "time": datetime.now().strftime("%H:%M:%S"),
            "status": getattr(result, "status", ""),
            "uid": getattr(result, "uid", ""),
            "file_count": len(getattr(result, "files", []) or []),
            "reason": getattr(result, "reason", ""),
        }
        self.after(0, lambda target=rule_id: self._refresh_result_label(target))

    def update_running_status(self):
        running_count = sum(1 for runtime in self.rule_runtimes.values() if runtime.get("is_running"))
        total_count = len(self.rule_cards)
        if running_count > 0:
            self.status_label.configure(text="🟢 规则检测中", text_color="#2CC985")
            self.running_count_label.configure(text=f"运行中: {running_count}/{total_count}")
        else:
            self.status_label.configure(text="⏸️ 待机中", text_color="gray")
            self.running_count_label.configure(text="")

    def _log_rule_result(self, rule_id: str, result):
        self._remember_rule_result(rule_id, result)
        mailbox_text = f"[{result.mailbox_alias}] " if getattr(result, "mailbox_alias", "") else ""
        folder_text = f"(文件夹: {result.mailbox_folder}) " if getattr(result, "mailbox_folder", "") else ""
        keyword_text = f"规则“{result.rule_keyword}”" if getattr(result, "rule_keyword", "") else "规则"
        if result.status == "processed":
            self.log_handler.success(
                f"{mailbox_text}{folder_text}{keyword_text}完成: uid={result.uid}, 附件数量={len(result.files)}",
                source=rule_id,
            )
        elif result.status == "not_found":
            self.log_handler.info(
                f"{mailbox_text}{folder_text}{keyword_text}: {result.reason}",
                source=rule_id,
            )
        elif result.status == "skipped":
            self.log_handler.warning(
                f"{mailbox_text}{folder_text}{keyword_text}: {result.reason}",
                source=rule_id,
            )
        else:
            self.log_handler.error(
                f"{mailbox_text}{folder_text}{keyword_text}: {result.reason}",
                source=rule_id,
            )

    def _run_rule_once(self, rule_id: str, *, force: bool, update_state: bool):
        rule = self.rule_configs.get(rule_id)
        if not rule:
            raise RuntimeError("规则不存在或已失效，请刷新页面")
        config = load_config()
        service = MailProcessingService(config)
        return service.process_single_rule(
            rule,
            force=force,
            update_state=update_state,
            event_callback=lambda level, message, src=rule_id: self.log_handler.log(level, message, source=src),
        )

    def start_single_rule(self, rule_id: str):
        runtime = self.rule_runtimes.get(rule_id)
        rule = self.rule_configs.get(rule_id)
        if not runtime or not rule:
            self.log_handler.error("未找到可启动的规则", source=rule_id)
            return
        if runtime["is_running"]:
            return
        if runtime["thread"] and runtime["thread"].is_alive():
            self.log_handler.warning("该规则仍在退出中，请稍后", source=rule_id)
            return

        runtime["pending_remove"] = False
        runtime["stop_event"] = threading.Event()
        runtime["is_running"] = True
        self._set_rule_card_state(rule_id, running=True, text="🟢 运行中", color="#2CC985")

        keyword = self._display_value(rule.get("keyword", ""), rule_id)
        trigger_mode = str(rule.get("trigger_mode", "periodic")).strip() or "periodic"
        interval_seconds = int(rule.get("poll_interval_seconds", 60) or 60)
        schedule_time = str(rule.get("schedule_time", "")).strip()
        self.log_handler.info("=" * 40, source=rule_id)
        self.log_handler.info(f"启动规则检测: {keyword}", source=rule_id)
        if trigger_mode == "timed":
            self.log_handler.info(f"检测方式: 每天定时 {schedule_time or '--:--'}", source=rule_id)
        else:
            self.log_handler.info(f"检测方式: 周期检测 / 间隔 {self.format_wait_text(interval_seconds)}", source=rule_id)
        self.log_handler.info("=" * 40, source=rule_id)

        def runner():
            cycle_count = 0
            try:
                while not runtime["stop_event"].is_set():
                    current_rule = self.rule_configs.get(rule_id)
                    if not current_rule:
                        runtime["stop_event"].set()
                        break
                    trigger_mode = str(current_rule.get("trigger_mode", "periodic")).strip() or "periodic"

                    if trigger_mode == "timed":
                        schedule_time = str(current_rule.get("schedule_time", "")).strip()
                        wait_seconds, target_text = self._seconds_until_daily_time(schedule_time)
                        self.log_handler.info(f"下次定时执行: {target_text}", source=rule_id)
                        for remaining in range(wait_seconds, 0, -1):
                            if runtime["stop_event"].is_set():
                                break
                            latest_rule = self.rule_configs.get(rule_id)
                            if not latest_rule:
                                runtime["stop_event"].set()
                                break
                            latest_mode = str(latest_rule.get("trigger_mode", "periodic")).strip() or "periodic"
                            latest_time = str(latest_rule.get("schedule_time", "")).strip()
                            if latest_mode != trigger_mode or latest_time != schedule_time:
                                self.log_handler.info("检测到新定时配置，正在按新配置重新计算执行时间", source=rule_id)
                                break
                            if remaining in (3600, 1800, 600, 300, 60, 30, 10, 5, 3, 2, 1):
                                self.log_handler.info(
                                    f"距离下一次定时执行还有 {self.format_wait_text(remaining)}",
                                    source=rule_id,
                                )
                            time.sleep(1)
                        else:
                            pass
                        if runtime["stop_event"].is_set():
                            break
                        latest_rule = self.rule_configs.get(rule_id)
                        if not latest_rule:
                            runtime["stop_event"].set()
                            break
                        if (
                            str(latest_rule.get("trigger_mode", "periodic")).strip() or "periodic"
                        ) != trigger_mode or str(latest_rule.get("schedule_time", "")).strip() != schedule_time:
                            continue

                    cycle_count += 1
                    cycle_started_at = time.time()
                    self.log_handler.info(f"第 {cycle_count} 轮开始检查", source=rule_id)
                    try:
                        result = self._run_rule_once(rule_id, force=False, update_state=True)
                        self._log_rule_result(rule_id, result)
                    except Exception as exc:
                        self.log_handler.error(f"执行出错: {exc}", source=rule_id)
                    finally:
                        elapsed = time.time() - cycle_started_at
                        self.log_handler.info(f"第 {cycle_count} 轮结束, 耗时 {elapsed:.1f} 秒", source=rule_id)

                    if runtime["stop_event"].is_set():
                        break

                    current_rule = self.rule_configs.get(rule_id)
                    if not current_rule:
                        runtime["stop_event"].set()
                        break
                    trigger_mode = str(current_rule.get("trigger_mode", "periodic")).strip() or "periodic"
                    if trigger_mode == "timed":
                        continue
                    wait_seconds = int(current_rule.get("poll_interval_seconds", 60) or 60)
                    self.log_handler.info(
                        f"进入等待: {self.format_wait_text(wait_seconds)} 后开始下一轮",
                        source=rule_id,
                    )
                    for remaining in range(wait_seconds, 0, -1):
                        if runtime["stop_event"].is_set():
                            break
                        latest_rule = self.rule_configs.get(rule_id)
                        if not latest_rule:
                            runtime["stop_event"].set()
                            break
                        latest_mode = str(latest_rule.get("trigger_mode", "periodic")).strip() or "periodic"
                        latest_wait_seconds = int(latest_rule.get("poll_interval_seconds", 60) or 60)
                        if latest_mode != trigger_mode:
                            self.log_handler.info("检测方式已切换，正在按新配置重新计算", source=rule_id)
                            break
                        if latest_wait_seconds != wait_seconds:
                            wait_seconds = latest_wait_seconds
                            self.log_handler.info(
                                f"检测到新轮询间隔，已切换为 {self.format_wait_text(wait_seconds)}",
                                source=rule_id,
                            )
                            break
                        if remaining in (60, 30, 10, 5, 3, 2, 1):
                            self.log_handler.info(
                                f"距离下一轮还有 {self.format_wait_text(remaining)}",
                                source=rule_id,
                            )
                        time.sleep(1)
            finally:
                runtime["is_running"] = False
                runtime["thread"] = None
                pending_remove = runtime.get("pending_remove", False)
                if pending_remove:
                    self.after(0, lambda target=rule_id: self._remove_rule_card(target))
                else:
                    self.after(
                        0,
                        lambda target=rule_id: self._set_rule_card_state(
                            target,
                            running=False,
                            text="⏸️ 已停止",
                            color="gray",
                        ),
                    )
                self.log_handler.info("规则线程已停止", source=rule_id)
                self.after(0, self.update_running_status)

        runtime["thread"] = threading.Thread(target=runner, daemon=True)
        runtime["thread"].start()

    def request_stop_rule(self, rule_id: str):
        runtime = self.rule_runtimes.get(rule_id)
        if not runtime:
            return
        if not runtime["is_running"] and not (runtime["thread"] and runtime["thread"].is_alive()):
            return
        runtime["is_running"] = False
        runtime["stop_event"].set()
        self._set_rule_card_state(rule_id, running=False, text="🟡 停止中", color="#FFB300")

    def stop_single_rule(self, rule_id: str):
        runtime = self.rule_runtimes.get(rule_id)
        if not runtime:
            return
        if not runtime["is_running"] and not (runtime["thread"] and runtime["thread"].is_alive()):
            return

        self.log_handler.info("正在停止规则...", source=rule_id)
        self.request_stop_rule(rule_id)

        worker = runtime.get("thread")
        if worker and worker.is_alive():
            worker.join(timeout=5)
        if worker and worker.is_alive():
            self.log_handler.warning("线程仍在退出中，可能正等待网络请求超时", source=rule_id)
            return

        if not runtime.get("pending_remove"):
            self._set_rule_card_state(rule_id, running=False, text="⏸️ 已停止", color="gray")
            self.log_handler.success("规则检测已停止", source=rule_id)

    def start_worker(self):
        if not self.rule_configs:
            self.log_handler.warning("当前没有启用规则，请先到设置页启用邮箱检测规则")
            return
        started_count = 0
        for rule_id in list(self.rule_configs.keys()):
            runtime = self.rule_runtimes.get(rule_id, {})
            before = bool(runtime.get("is_running"))
            self.start_single_rule(rule_id)
            after = bool(self.rule_runtimes.get(rule_id, {}).get("is_running"))
            if not before and after:
                started_count += 1
        if started_count == 0:
            self.log_handler.info("没有新的规则被启动")

    def stop_worker(self):
        any_running = any(runtime.get("is_running") for runtime in self.rule_runtimes.values())
        for rule_id in list(self.rule_cards.keys()):
            self.stop_single_rule(rule_id)
        if any_running:
            self.log_handler.success("全部规则已发送停止指令")

    @property
    def is_running(self) -> bool:
        return any(runtime.get("is_running") for runtime in self.rule_runtimes.values())

    def _run_rule_test(self, rule_id: str):
        try:
            rule = self.rule_configs.get(rule_id, {})
            keyword = self._display_value(rule.get("keyword", ""), rule_id)
            self.log_handler.info("=" * 40, source=rule_id)
            self.log_handler.info(f"开始测试规则: {keyword}", source=rule_id)
            self.log_handler.info("=" * 40, source=rule_id)
            result = self._run_rule_once(rule_id, force=True, update_state=False)
            self._log_rule_result(rule_id, result)

            if result.status != "processed":
                self.log_handler.warning("测试完成，但本次没有成功处理邮件", source=rule_id)
                self.log_handler.info("排查建议：确认所属邮箱、文件夹、主题关键字、附件格式是否配置正确", source=rule_id)
                return

            self.log_handler.success(f"找到邮件 UID: {result.uid}", source=rule_id)
            self.log_handler.success(f"提取到 {len(result.files)} 个附件:", source=rule_id)
            for index, file_path in enumerate(result.files, 1):
                size_mb = file_path.stat().st_size / (1024 * 1024)
                self.log_handler.info(f"  {index}. {file_path.name} ({size_mb:.2f} MB)", source=rule_id)
            self.log_handler.success("测试完成！规则已成功处理邮件", source=rule_id)
        except Exception as exc:
            self.log_handler.error(f"测试过程出错: {exc}", source=rule_id)
        finally:
            self._rule_test_running.discard(rule_id)

    def test_single_rule(self, rule_id: str):
        runtime = self.rule_runtimes.get(rule_id)
        if runtime and runtime.get("is_running"):
            self.log_handler.warning("规则正在运行中，请先停止后再测试", source=rule_id)
            return
        if rule_id in self._rule_test_running:
            self.log_handler.warning("该规则测试任务仍在执行中，请稍后", source=rule_id)
            return
        self._rule_test_running.add(rule_id)
        threading.Thread(target=lambda: self._run_rule_test(rule_id), daemon=True).start()

    def test_once(self):
        if not self.rule_configs:
            self.log_handler.warning("当前没有启用规则，请先到设置页启用邮箱检测规则")
            return
        for rule_id in self.rule_configs.keys():
            runtime = self.rule_runtimes.get(rule_id)
            if runtime and runtime.get("is_running"):
                self.log_handler.warning("存在运行中的规则，请先停止后再执行全部测试")
                return
        for rule_id in self.rule_configs.keys():
            if rule_id in self._rule_test_running:
                self.log_handler.warning("存在测试中的规则，请稍后再试")
                return
        for rule_id in self.rule_configs.keys():
            self.test_single_rule(rule_id)

    def on_page_activated(self):
        self.sync_rule_views()

    def on_external_config_updated(self):
        self.sync_rule_views()
