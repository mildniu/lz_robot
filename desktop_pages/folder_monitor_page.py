import json
import threading
import time
import tkinter.messagebox as messagebox
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
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
        self.is_running = False
        self.observer = None
        self.executor = None
        self.tracker = FileSentTracker()
        self.monitors = {}
        self._inflight_events = set()
        self._inflight_lock = threading.Lock()
        self.setup_ui()
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

        self.start_btn = ModernButton(
            right_panel,
            text="启动",
            icon="▶️",
            command=self.start_monitor,
            width=120,
            height=45,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color="#2CC985",
            hover_color="#239B6D",
        )
        self.start_btn.pack(side="left", padx=5)

        self.stop_btn = ModernButton(
            right_panel,
            text="停止",
            icon="⏹️",
            command=self.stop_monitor,
            width=120,
            height=45,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color="#EF5350",
            hover_color="#C62828",
            state="disabled",
        )
        self.stop_btn.pack(side="left", padx=5)

        self.monitor_status_label = ctk.CTkLabel(right_panel, text="", font=ctk.CTkFont(size=14))
        self.monitor_status_label.pack(side="left", padx=10)

        log_container = ctk.CTkFrame(self)
        log_container.pack(fill="both", expand=True, padx=25, pady=(0, 20))

        log_header = ctk.CTkFrame(log_container, height=50)
        log_header.pack(fill="x")

        log_title = ctk.CTkLabel(log_header, text="📋 文件夹检测日志", font=ctk.CTkFont(size=18, weight="bold"))
        log_title.pack(side="left", padx=10, pady=10)

        clear_btn = ctk.CTkButton(
            log_header, text="🗑️ 清空", command=self.clear_logs, width=80, height=35, font=ctk.CTkFont(size=12)
        )
        clear_btn.pack(side="right", padx=10, pady=7)

        self.log_text = ctk.CTkTextbox(
            log_container,
            font=ctk.CTkFont(family="Consolas", size=12),
            wrap="word",
            fg_color=("gray98", "gray15"),
        )
        self.log_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def append_log(self, log_entry: Dict[str, str]):
        icons = {"INFO": "ℹ️", "SUCCESS": "✅", "WARNING": "⚠️", "ERROR": "❌"}
        log_line = f"[{log_entry['time']}] {icons.get(log_entry['level'], '•')} {log_entry['message']}\n"
        self.log_text.insert("end", log_line)
        if self.auto_scroll_log:
            self.log_text.see("end")

    def clear_logs(self):
        self.log_text.delete("1.0", "end")

    def load_monitor_config(self) -> Dict:
        config_file = Path("settings/folder_monitor_config.json")
        if config_file.exists():
            try:
                return json.loads(config_file.read_text(encoding="utf-8"))
            except Exception:
                return {}
        return {}

    def start_monitor(self):
        if self.is_running:
            return

        try:
            monitor_config = self.load_monitor_config()
            alias_config = load_webhook_aliases()
            aliases = alias_config.get("aliases", {})
            enabled_monitors = {
                key: value
                for key, value in monitor_config.items()
                if value.get("enabled", False) and value.get("path")
            }
            if not enabled_monitors:
                messagebox.showwarning("提示", "没有启用的文件夹检测\n\n请在设置页面配置文件夹检测并勾选启用")
                return

            self.is_running = True
            self.start_btn.configure(state="disabled")
            self.stop_btn.configure(state="normal")
            self.status_label.configure(text="🟢 检测中", text_color="#2CC985")

            self.log_handler.info("=" * 50)
            self.log_handler.info("启动文件夹检测...")
            self.log_handler.info(f"启用的检测数量: {len(enabled_monitors)}")

            self.observer = Observer()
            self.executor = ThreadPoolExecutor(max_workers=4, thread_name_prefix="folder-monitor")
            with self._inflight_lock:
                self._inflight_events.clear()
            for folder_key, config in enabled_monitors.items():
                folder_path = Path(config["path"])
                webhook_alias = config.get("webhook_alias", "")
                webhook_url = (config.get("webhook_url") or "").strip()
                if webhook_alias:
                    resolved = resolve_webhook_url(webhook_alias, aliases)
                    if resolved:
                        webhook_url = resolved

                if not folder_path.exists():
                    self.log_handler.warning(f"文件夹不存在: {folder_path}")
                    continue
                if not webhook_url:
                    self.log_handler.warning(f"Webhook 地址为空: {folder_path}")
                    self.log_handler.warning("跳过此检测，请在设置中配置推送别名")
                    continue
                if not webhook_url.startswith("http"):
                    self.log_handler.warning(f"Webhook URL格式不正确: {webhook_url}")
                    self.log_handler.warning("跳过此检测")
                    continue

                handler = FolderMonitorHandler(
                    callback=lambda path, event, webhook=webhook_url: self.on_file_event(path, event, webhook),
                    log_handler=self.log_handler,
                )
                self.observer.schedule(handler, str(folder_path), recursive=False)
                self.monitors[folder_key] = {
                    "path": str(folder_path),
                    "webhook_url": webhook_url,
                    "webhook_alias": webhook_alias,
                }
                alias_text = f" (别名: {webhook_alias})" if webhook_alias else ""
                self.log_handler.success(f"检测文件夹: {folder_path}{alias_text}")

            if not self.monitors:
                self.is_running = False
                self.start_btn.configure(state="normal")
                self.stop_btn.configure(state="disabled")
                self.status_label.configure(text="⏸️ 待机中", text_color="gray")
                messagebox.showwarning("提示", "未找到可用的文件夹检测配置，请检查路径和推送别名")
                return

            self.observer.start()
            self.log_handler.success("=" * 50)
            self.log_handler.success("文件夹检测已启动")
            self.update_monitor_status()
            self.scan_existing_files(enabled_monitors, aliases)
        except Exception as exc:
            self.is_running = False
            self.start_btn.configure(state="normal")
            self.stop_btn.configure(state="disabled")
            self.status_label.configure(text="❌ 错误", text_color="#EF5350")
            self.log_handler.error(f"启动失败: {exc}")

    def stop_monitor(self):
        if not self.is_running:
            return

        self.log_handler.info("正在停止文件夹检测...")
        self.is_running = False

        if self.observer:
            self.observer.stop()
            self.observer.join(timeout=5)
            self.observer = None

        if self.executor:
            self.executor.shutdown(wait=False, cancel_futures=True)
            self.executor = None
        with self._inflight_lock:
            self._inflight_events.clear()

        self.monitors.clear()
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        self.status_label.configure(text="⏸️ 已停止", text_color="gray")
        self.update_monitor_status()
        self.log_handler.success("文件夹检测已停止")

    def on_file_event(self, file_path: str, event_type: str, webhook_url: str):
        if not self.is_running:
            return
        event_key = f"{webhook_url}|{file_path}"
        with self._inflight_lock:
            if event_key in self._inflight_events:
                return
            self._inflight_events.add(event_key)

        if not self.executor:
            with self._inflight_lock:
                self._inflight_events.discard(event_key)
            return
        try:
            self.executor.submit(self._process_file_event, file_path, event_type, webhook_url, event_key)
        except RuntimeError:
            with self._inflight_lock:
                self._inflight_events.discard(event_key)

    def _process_file_event(self, file_path: str, event_type: str, webhook_url: str, event_key: str):
        try:
            if event_type == "modified":
                # Give writers some time to finish flushing the file.
                time.sleep(1)

            path = Path(file_path)
            if not path.exists() or not path.is_file():
                self.log_handler.warning(f"文件不存在或不可用，跳过: {path}")
                return
            if not webhook_url:
                self.log_handler.error(f"Webhook URL为空，无法处理文件: {path.name}")
                self.log_handler.error("请在设置页面配置Webhook URL")
                return
            if not webhook_url.startswith("http"):
                self.log_handler.error(f"Webhook URL格式不正确: {webhook_url}")
                return
            if self.tracker.is_sent(path, webhook_url):
                self.log_handler.info(f"文件已处理，跳过: {path.name}")
                return

            self.log_handler.info(f"处理文件: {path.name} ({path.suffix}) [{event_type}]")
            try:
                file_id = send_file_via_webhook(
                    path,
                    webhook_url,
                    event_callback=lambda level, message: self.log_handler.log(level, message),
                )
                self.tracker.mark_sent(path, webhook_url, file_id)
            except Exception as exc:
                self.log_handler.error(f"处理失败: {exc}")
        except Exception as exc:
            self.log_handler.error(f"事件处理错误: {exc}")
        finally:
            with self._inflight_lock:
                self._inflight_events.discard(event_key)

    def scan_existing_files(self, monitors: Dict, aliases: Dict[str, str]):
        self.log_handler.info("扫描现有文件...")
        for config in monitors.values():
            folder_path = Path(config["path"])
            webhook_alias = config.get("webhook_alias", "")
            webhook_url = (config.get("webhook_url") or "").strip()
            if webhook_alias:
                resolved = resolve_webhook_url(webhook_alias, aliases)
                if resolved:
                    webhook_url = resolved
            if not folder_path.exists():
                continue
            if not webhook_url:
                continue
            try:
                for file_path in folder_path.iterdir():
                    if file_path.is_file() and not self.tracker.is_sent(file_path, webhook_url):
                        self.log_handler.info(f"发现未处理文件: {file_path.name} ({file_path.suffix})")
                        self.on_file_event(str(file_path), "existing", webhook_url)
            except Exception as exc:
                self.log_handler.error(f"扫描文件夹失败 {folder_path}: {exc}")

    def update_monitor_status(self):
        if self.is_running and self.monitors:
            self.monitor_status_label.configure(
                text=f"检测中: {len(self.monitors)}个文件夹",
                text_color="#2CC985",
            )
        else:
            self.monitor_status_label.configure(text="")
