import json
import threading
import time
import tkinter.messagebox as messagebox
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from tkinter import filedialog
from typing import Dict

import customtkinter as ctk
from watchdog.observers import Observer

from mail_forwarder.processing_service import send_file_via_webhook

from .common import FileSentTracker, FolderMonitorHandler, LogHandler, ModernButton
from .webhook_alias_store import load_webhook_aliases, resolve_webhook_url


class FolderMonitorPage(ctk.CTkFrame):
    def __init__(self, master, log_handler: LogHandler, auto_scroll_log: bool = True, **kwargs):
        super().__init__(master, **kwargs)
        self.log_handler = log_handler
        self.auto_scroll_log = auto_scroll_log
        self.tracker = FileSentTracker()
        self.monitor_cards: dict[str, dict] = {}
        self.monitor_runtimes: dict[str, dict] = {}
        self.monitor_log_texts: dict[str, ctk.CTkTextbox] = {}
        self.monitor_config_markers: dict[str, tuple] = {}
        self.setup_ui()
        self._build_monitor_cards()
        self.log_handler.add_callback(self.append_log)

    def setup_ui(self):
        control_bar = ctk.CTkFrame(self, height=100)
        control_bar.pack(fill="x", padx=25, pady=(20, 10))

        left_panel = ctk.CTkFrame(control_bar, fg_color="transparent")
        left_panel.pack(side="left", fill="both", expand=True)

        title = ctk.CTkLabel(left_panel, text="📁 文件夹检测", font=ctk.CTkFont(size=26, weight="bold"))
        title.pack(pady=(5, 5), anchor="w")

        self.status_label = ctk.CTkLabel(
            left_panel, text="⏸️ 待机中", font=ctk.CTkFont(size=16), text_color="gray"
        )
        self.status_label.pack(anchor="w")

        right_panel = ctk.CTkFrame(control_bar, fg_color="transparent")
        right_panel.pack(side="right", padx=10)

        self.start_all_btn = ModernButton(
            right_panel,
            text="全部启动",
            icon="▶️",
            command=self.start_all_monitors,
            width=120,
            height=45,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color="#2CC985",
            hover_color="#239B6D",
        )
        self.start_all_btn.pack(side="left", padx=5)

        self.stop_all_btn = ModernButton(
            right_panel,
            text="全部停止",
            icon="⏹️",
            command=self.stop_all_monitors,
            width=120,
            height=45,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color="#EF5350",
            hover_color="#C62828",
        )
        self.stop_all_btn.pack(side="left", padx=5)

        self.monitor_status_label = ctk.CTkLabel(right_panel, text="", font=ctk.CTkFont(size=14))
        self.monitor_status_label.pack(side="left", padx=10)

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=25, pady=(0, 20))

        self.monitor_list = ctk.CTkScrollableFrame(body, fg_color=("gray96", "gray17"), width=420)
        self.monitor_list.pack(side="left", fill="y", padx=(0, 16))

        self.log_container = ctk.CTkFrame(body)
        self.log_container.pack(side="right", fill="both", expand=True)

        log_header = ctk.CTkFrame(self.log_container, height=50)
        log_header.pack(fill="x")

        log_title = ctk.CTkLabel(log_header, text="📋 文件夹检测日志", font=ctk.CTkFont(size=18, weight="bold"))
        log_title.pack(side="left", padx=10, pady=10)

        clear_btn = ctk.CTkButton(
            log_header,
            text="🗑️ 清空当前",
            command=self.clear_current_logs,
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

    def _build_monitor_cards(self):
        monitor_config = self.load_monitor_config()
        for index in range(1, 4):
            key = f"folder_{index}"
            config = monitor_config.get(key, {})
            self._create_monitor_card(key, index, config)
            tab = self.log_tabview.add(f"检测{index}")
            text_widget = ctk.CTkTextbox(
                tab,
                font=ctk.CTkFont(family="Consolas", size=12),
                wrap="word",
                fg_color=("gray98", "gray15"),
            )
            text_widget.pack(fill="both", expand=True, padx=6, pady=6)
            self.monitor_log_texts[key] = text_widget

    def _create_monitor_card(self, key: str, index: int, config: Dict):
        card = ctk.CTkFrame(self.monitor_list, fg_color=("gray95", "gray20"))
        card.pack(fill="x", padx=6, pady=8)

        header = ctk.CTkFrame(card, fg_color="transparent")
        header.pack(fill="x", padx=12, pady=(12, 8))

        ctk.CTkLabel(header, text=f"检测 {index}", font=ctk.CTkFont(size=16, weight="bold")).pack(side="left")
        enabled = bool(config.get("enabled", False))
        enabled_text = "配置已启用" if enabled else "配置未启用"
        enabled_label = ctk.CTkLabel(
            header,
            text=enabled_text,
            font=ctk.CTkFont(size=11),
            text_color=("#2E7D32", "#81C784") if enabled else ("gray35", "gray70"),
        )
        enabled_label.pack(side="left", padx=(10, 0))

        status_label = ctk.CTkLabel(
            header,
            text="⏸️ 待机中",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="gray",
        )
        status_label.pack(side="right")

        summary = ctk.CTkFrame(card, fg_color="transparent")
        summary.pack(fill="x", padx=12, pady=(0, 8))

        path_value = str(config.get("path", "")).strip() or "(未配置路径)"
        alias_value = str(config.get("webhook_alias", "")).strip() or "(未配置机器人)"
        path_label = ctk.CTkLabel(summary, text=f"路径: {path_value}", anchor="w", justify="left")
        path_label.pack(fill="x")
        alias_label = ctk.CTkLabel(summary, text=f"机器人: {alias_value}", anchor="w", justify="left")
        alias_label.pack(fill="x", pady=(4, 0))

        actions = ctk.CTkFrame(card, fg_color="transparent")
        actions.pack(fill="x", padx=12, pady=(0, 12))

        start_btn = ModernButton(
            actions,
            text="启动",
            icon="▶️",
            command=lambda target=key: self.start_single_monitor(target),
            width=90,
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
            command=lambda target=key: self.stop_single_monitor(target),
            width=90,
            height=34,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#EF5350",
            hover_color="#C62828",
            state="disabled",
        )
        stop_btn.pack(side="left", padx=(0, 8))

        open_log_btn = ctk.CTkButton(
            actions,
            text="查看日志",
            width=90,
            height=34,
            font=ctk.CTkFont(size=12),
            command=lambda idx=index: self.log_tabview.set(f"检测{idx}"),
        )
        open_log_btn.pack(side="left")

        self.monitor_cards[key] = {
            "enabled_label": enabled_label,
            "status_label": status_label,
            "start_btn": start_btn,
            "stop_btn": stop_btn,
            "path_label": path_label,
            "alias_label": alias_label,
        }

    def refresh_monitor_cards(self):
        monitor_config = self.load_monitor_config()
        for index in range(1, 4):
            key = f"folder_{index}"
            config = monitor_config.get(key, {})
            card = self.monitor_cards.get(key, {})
            enabled = bool(config.get("enabled", False))
            enabled_label = card.get("enabled_label")
            path_label = card.get("path_label")
            alias_label = card.get("alias_label")
            if enabled_label:
                enabled_label.configure(
                    text="配置已启用" if enabled else "配置未启用",
                    text_color=("#2E7D32", "#81C784") if enabled else ("gray35", "gray70"),
                )
            if path_label:
                path_label.configure(text=f"路径: {str(config.get('path', '')).strip() or '(未配置路径)'}")
            if alias_label:
                alias_label.configure(text=f"机器人: {str(config.get('webhook_alias', '')).strip() or '(未配置机器人)'}")
        self.update_monitor_status()

    @staticmethod
    def _config_marker(config: dict | None) -> tuple:
        if not config:
            return ("", "", "", False)
        return (
            str(config.get("path", "")),
            str(config.get("webhook_alias", "")),
            str(config.get("webhook_url", "")),
            bool(config.get("enabled", False)),
        )

    def append_log(self, log_entry: Dict[str, str]):
        icons = {"INFO": "ℹ️", "SUCCESS": "✅", "WARNING": "⚠️", "ERROR": "❌"}
        log_line = f"[{log_entry['time']}] {icons.get(log_entry['level'], '•')} {log_entry['message']}\n"
        self.overview_log_text.insert("end", log_line)
        if self.auto_scroll_log:
            self.overview_log_text.see("end")

        source = str(log_entry.get("source", "global"))
        monitor_text = self.monitor_log_texts.get(source)
        if monitor_text:
            monitor_text.insert("end", log_line)
            if self.auto_scroll_log:
                monitor_text.see("end")

    def clear_current_logs(self):
        current_tab = self.log_tabview.get()
        if current_tab == "总览":
            self.overview_log_text.delete("1.0", "end")
            return
        for key, text_widget in self.monitor_log_texts.items():
            tab_name = "检测" + key.split("_")[-1]
            if tab_name == current_tab:
                text_widget.delete("1.0", "end")
                return

    def export_current_logs(self):
        current_tab = self.log_tabview.get()
        content = ""
        if current_tab == "总览":
            content = self.overview_log_text.get("1.0", "end").strip()
        else:
            for key, text_widget in self.monitor_log_texts.items():
                tab_name = "检测" + key.split("_")[-1]
                if tab_name == current_tab:
                    content = text_widget.get("1.0", "end").strip()
                    break
        if not content:
            messagebox.showwarning("提示", "当前日志为空，暂无可导出的内容")
            return
        export_dir = Path("exports")
        export_dir.mkdir(parents=True, exist_ok=True)
        file_path = filedialog.asksaveasfilename(
            title="导出当前日志",
            initialdir=str(export_dir.resolve()),
            initialfile=f"folder_{current_tab}_{time.strftime('%Y%m%d_%H%M%S')}.log".replace(" ", "_"),
            defaultextension=".log",
            filetypes=[("Log Files", "*.log"), ("Text Files", "*.txt"), ("All Files", "*.*")],
        )
        if not file_path:
            return
        Path(file_path).write_text(content + "\n", encoding="utf-8")
        messagebox.showinfo("导出成功", f"日志已导出到:\n{file_path}")

    def load_monitor_config(self) -> Dict:
        config_file = Path("settings/folder_monitor_config.json")
        if config_file.exists():
            try:
                return json.loads(config_file.read_text(encoding="utf-8"))
            except Exception:
                return {}
        return {}

    def _resolve_monitor_config(self, key: str) -> dict | None:
        monitor_config = self.load_monitor_config()
        raw = monitor_config.get(key, {})
        if not raw:
            return None

        aliases = load_webhook_aliases().get("aliases", {})
        folder_path = Path(str(raw.get("path", "")).strip())
        webhook_alias = str(raw.get("webhook_alias", "")).strip()
        webhook_url = str(raw.get("webhook_url", "")).strip()
        if webhook_alias:
            resolved = resolve_webhook_url(webhook_alias, aliases)
            if resolved:
                webhook_url = resolved

        return {
            "enabled": bool(raw.get("enabled", False)),
            "path": folder_path,
            "webhook_alias": webhook_alias,
            "webhook_url": webhook_url,
        }

    def _get_runtime(self, key: str) -> dict:
        runtime = self.monitor_runtimes.get(key)
        if runtime is None:
            runtime = {
                "observer": None,
                "executor": None,
                "is_running": False,
                "config": None,
                "inflight_events": set(),
                "lock": threading.Lock(),
                "pending_refresh": False,
            }
            self.monitor_runtimes[key] = runtime
        return runtime

    def _set_card_state(self, key: str, *, running: bool, text: str, color: str):
        card = self.monitor_cards.get(key, {})
        status_label = card.get("status_label")
        start_btn = card.get("start_btn")
        stop_btn = card.get("stop_btn")
        if status_label:
            status_label.configure(text=text, text_color=color)
        if start_btn:
            start_btn.configure(state="disabled" if running else "normal")
        if stop_btn:
            stop_btn.configure(state="normal" if running else "disabled")
        self.update_monitor_status()

    def start_single_monitor(self, key: str):
        runtime = self._get_runtime(key)
        if runtime["is_running"]:
            return

        config = self._resolve_monitor_config(key)
        if not config:
            self.log_handler.error(f"{key} 未找到有效配置", source=key)
            return
        if not config.get("enabled", False):
            messagebox.showwarning("提示", f"{key} 未启用，请先在设置页启用该监测项")
            return

        folder_path = config["path"]
        webhook_url = str(config.get("webhook_url", "")).strip()
        webhook_alias = str(config.get("webhook_alias", "")).strip()
        if not folder_path.exists():
            self.log_handler.error(f"文件夹不存在: {folder_path}", source=key)
            return
        if not webhook_url:
            self.log_handler.error("未配置推送机器人", source=key)
            return
        if not webhook_url.startswith("http"):
            self.log_handler.error(f"Webhook URL格式不正确: {webhook_url}", source=key)
            return

        observer = Observer()
        executor = ThreadPoolExecutor(max_workers=2, thread_name_prefix=f"{key}-monitor")
        with runtime["lock"]:
            runtime["inflight_events"].clear()
        handler = FolderMonitorHandler(
            callback=lambda path, event, monitor_key=key: self.on_file_event(monitor_key, path, event),
            log_handler=self.log_handler,
            source=key,
        )
        observer.schedule(handler, str(folder_path), recursive=False)
        observer.start()

        runtime["observer"] = observer
        runtime["executor"] = executor
        runtime["is_running"] = True
        runtime["config"] = config
        self.monitor_config_markers[key] = self._config_marker(config)

        alias_text = f" (别名: {webhook_alias})" if webhook_alias else ""
        self.log_handler.info(f"启动监测: {folder_path}{alias_text}", source=key)
        self._set_card_state(key, running=True, text="🟢 运行中", color="#2CC985")
        self.scan_existing_files(key)

    def stop_single_monitor(self, key: str):
        runtime = self._get_runtime(key)
        if not runtime["is_running"]:
            return

        self.log_handler.info("正在停止监测...", source=key)
        runtime["is_running"] = False
        observer = runtime.get("observer")
        if observer:
            observer.stop()
            observer.join(timeout=5)
        executor = runtime.get("executor")
        if executor:
            executor.shutdown(wait=False, cancel_futures=True)
        with runtime["lock"]:
            runtime["inflight_events"].clear()
        runtime["observer"] = None
        runtime["executor"] = None
        runtime["config"] = None
        runtime["pending_refresh"] = False
        self._set_card_state(key, running=False, text="⏸️ 已停止", color="gray")
        self.log_handler.success("监测已停止", source=key)

    def start_all_monitors(self):
        started_count = 0
        for index in range(1, 4):
            key = f"folder_{index}"
            before = self._get_runtime(key)["is_running"]
            self.start_single_monitor(key)
            after = self._get_runtime(key)["is_running"]
            if not before and after:
                started_count += 1
        if started_count == 0:
            self.status_label.configure(text="⏸️ 待机中", text_color="gray")
        else:
            self.status_label.configure(text="🟢 检测中", text_color="#2CC985")

    def stop_all_monitors(self):
        any_running = any(self._get_runtime(f"folder_{index}")["is_running"] for index in range(1, 4))
        for index in range(1, 4):
            self.stop_single_monitor(f"folder_{index}")
        if any_running:
            self.status_label.configure(text="⏸️ 已停止", text_color="gray")
        else:
            self.status_label.configure(text="⏸️ 待机中", text_color="gray")

    def stop_monitor(self):
        self.stop_all_monitors()

    def on_file_event(self, key: str, file_path: str, event_type: str):
        runtime = self._get_runtime(key)
        if not runtime["is_running"]:
            return
        config = runtime.get("config") or {}
        webhook_url = str(config.get("webhook_url", "")).strip()
        event_key = f"{webhook_url}|{file_path}"

        with runtime["lock"]:
            if event_key in runtime["inflight_events"]:
                return
            runtime["inflight_events"].add(event_key)

        executor = runtime.get("executor")
        if not executor:
            with runtime["lock"]:
                runtime["inflight_events"].discard(event_key)
            return

        try:
            executor.submit(self._process_file_event, key, file_path, event_type, webhook_url, event_key)
        except RuntimeError:
            with runtime["lock"]:
                runtime["inflight_events"].discard(event_key)

    def _process_file_event(self, key: str, file_path: str, event_type: str, webhook_url: str, event_key: str):
        runtime = self._get_runtime(key)
        try:
            if event_type == "modified":
                time.sleep(1)

            path = Path(file_path)
            if not path.exists() or not path.is_file():
                self.log_handler.warning(f"文件不存在或不可用，跳过: {path}", source=key)
                return
            if not webhook_url:
                self.log_handler.error(f"Webhook URL为空，无法处理文件: {path.name}", source=key)
                return
            if self.tracker.is_sent(path, webhook_url):
                self.log_handler.info(f"文件已处理，跳过: {path.name}", source=key)
                return

            self.log_handler.info(f"处理文件: {path.name} ({path.suffix}) [{event_type}]", source=key)
            file_id = send_file_via_webhook(
                path,
                webhook_url,
                event_callback=lambda level, message, src=key: self.log_handler.log(level, message, source=src),
            )
            self.tracker.mark_sent(path, webhook_url, file_id)
        except Exception as exc:
            self.log_handler.error(f"处理失败: {exc}", source=key)
            self._set_card_state(key, running=runtime["is_running"], text="❌ 最近出错", color="#EF5350")
        finally:
            with runtime["lock"]:
                runtime["inflight_events"].discard(event_key)

    def scan_existing_files(self, key: str):
        runtime = self._get_runtime(key)
        config = runtime.get("config") or {}
        folder_path = config.get("path")
        webhook_url = str(config.get("webhook_url", "")).strip()
        if not folder_path or not webhook_url:
            return
        self.log_handler.info("扫描现有文件...", source=key)
        try:
            for file_path in Path(folder_path).iterdir():
                if file_path.is_file() and not self.tracker.is_sent(file_path, webhook_url):
                    self.log_handler.info(f"发现未处理文件: {file_path.name} ({file_path.suffix})", source=key)
                    self.on_file_event(key, str(file_path), "existing")
        except Exception as exc:
            self.log_handler.error(f"扫描文件夹失败 {folder_path}: {exc}", source=key)

    def update_monitor_status(self):
        running_count = sum(1 for item in self.monitor_runtimes.values() if item.get("is_running"))
        if running_count > 0:
            self.monitor_status_label.configure(text=f"运行中: {running_count} 个监测项", text_color="#2CC985")
            self.status_label.configure(text="🟢 检测中", text_color="#2CC985")
        else:
            self.monitor_status_label.configure(text="", text_color="gray")
            self.status_label.configure(text="⏸️ 待机中", text_color="gray")

    def on_page_activated(self):
        self.refresh_monitor_cards()

    def on_external_config_updated(self):
        self.refresh_monitor_cards()
        self.apply_runtime_config_updates()

    def _restart_monitor_with_config(self, key: str, config: dict):
        runtime = self._get_runtime(key)
        observer = runtime.get("observer")
        if observer:
            observer.stop()
            observer.join(timeout=5)
        executor = runtime.get("executor")
        if executor:
            executor.shutdown(wait=False, cancel_futures=True)
        with runtime["lock"]:
            runtime["inflight_events"].clear()

        folder_path = config["path"]
        webhook_url = str(config.get("webhook_url", "")).strip()
        webhook_alias = str(config.get("webhook_alias", "")).strip()
        if not folder_path.exists():
            runtime["observer"] = None
            runtime["executor"] = None
            runtime["is_running"] = False
            runtime["config"] = None
            runtime["pending_refresh"] = False
            self._set_card_state(key, running=False, text="❌ 路径无效", color="#EF5350")
            self.log_handler.error(f"热刷新失败，文件夹不存在: {folder_path}", source=key)
            return
        if not webhook_url or not webhook_url.startswith("http"):
            runtime["observer"] = None
            runtime["executor"] = None
            runtime["is_running"] = False
            runtime["config"] = None
            runtime["pending_refresh"] = False
            self._set_card_state(key, running=False, text="❌ 机器人无效", color="#EF5350")
            self.log_handler.error("热刷新失败，推送机器人配置无效", source=key)
            return

        observer = Observer()
        executor = ThreadPoolExecutor(max_workers=2, thread_name_prefix=f"{key}-monitor")
        handler = FolderMonitorHandler(
            callback=lambda path, event, monitor_key=key: self.on_file_event(monitor_key, path, event),
            log_handler=self.log_handler,
            source=key,
        )
        observer.schedule(handler, str(folder_path), recursive=False)
        observer.start()

        runtime["observer"] = observer
        runtime["executor"] = executor
        runtime["is_running"] = True
        runtime["config"] = config
        runtime["pending_refresh"] = False
        self.monitor_config_markers[key] = self._config_marker(config)
        self._set_card_state(key, running=True, text="🟢 运行中", color="#2CC985")
        self.log_handler.info(
            f"配置热刷新完成: 路径={folder_path}，机器人={webhook_alias or '(未命名)'}",
            source=key,
        )
        self.scan_existing_files(key)

    def apply_runtime_config_updates(self):
        monitor_config = self.load_monitor_config()
        for index in range(1, 4):
            key = f"folder_{index}"
            runtime = self._get_runtime(key)
            new_config = self._resolve_monitor_config(key)
            new_marker = self._config_marker(new_config)
            old_marker = self.monitor_config_markers.get(key, ("", "", "", False))

            if not runtime.get("is_running"):
                self.monitor_config_markers[key] = new_marker
                continue

            if new_marker == old_marker:
                continue

            if not new_config or not new_config.get("enabled", False):
                self.log_handler.warning("配置已禁用，正在自动停止当前监测", source=key)
                self.stop_single_monitor(key)
                self.monitor_config_markers[key] = new_marker
                continue

            old_config = runtime.get("config") or {}
            old_path = Path(str(old_config.get("path", "")))
            new_path = Path(str(new_config.get("path", "")))
            old_webhook_url = str(old_config.get("webhook_url", "")).strip()
            new_webhook_url = str(new_config.get("webhook_url", "")).strip()

            if old_path != new_path:
                self.log_handler.info(f"检测到路径变更，正在热切换到: {new_path}", source=key)
                self._restart_monitor_with_config(key, new_config)
                continue

            runtime["config"] = new_config
            runtime["pending_refresh"] = False
            self.monitor_config_markers[key] = new_marker
            if old_webhook_url != new_webhook_url:
                self.log_handler.info("检测到推送机器人变更，新配置已立即生效", source=key)
            else:
                self.log_handler.info("检测配置已热刷新", source=key)
