from __future__ import annotations

import json
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QTextCursor
from PySide6.QtWidgets import (
    QCheckBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)
from watchdog.observers import Observer

from app_services import FileSentTracker, FolderMonitorHandler, load_webhook_aliases, resolve_webhook_url
from mail_forwarder import load_config
from mail_forwarder.processing_service import send_file_via_webhook
from qt_components import create_status_pill, set_button_variant

from .base import BasePage


class FolderMonitorPage(BasePage):
    OBSERVER_JOIN_TIMEOUT_SECONDS = 5

    log_signal = Signal(str, str, str)
    state_signal = Signal(str, str, str, bool)
    structure_refresh_signal = Signal()

    def __init__(self, log_bus) -> None:
        super().__init__(log_bus, "文件夹检测")
        self.config = load_config()
        self.tracker = FileSentTracker()
        self.monitor_cards: dict[str, dict[str, object]] = {}
        self.monitor_runtimes: dict[str, dict] = {}
        self.monitor_config_markers: dict[str, tuple] = {}
        self.log_buffers: dict[str, list[str]] = {"global": []}
        self.log_widgets: dict[str, QTextEdit] = {}
        self.monitor_rows_layout: QVBoxLayout | None = None

        self.log_signal.connect(self._append_log_entry)
        self.state_signal.connect(self._apply_monitor_state)
        self.structure_refresh_signal.connect(self.reload_monitors)

        self._build_ui()
        self.reload_monitors()

    def _build_ui(self) -> None:
        action_strip = QFrame(self)
        action_strip.setObjectName("ActionStrip")
        actions = QHBoxLayout(action_strip)
        actions.setContentsMargins(14, 12, 14, 12)
        actions.setSpacing(8)

        start_all_btn = QPushButton("全部启动", self)
        start_all_btn.clicked.connect(self.start_all_monitors)
        stop_all_btn = QPushButton("全部停止", self)
        stop_all_btn.clicked.connect(self.stop_all_monitors)
        refresh_btn = QPushButton("刷新配置", self)
        refresh_btn.clicked.connect(self.reload_monitors)

        for button in [start_all_btn, stop_all_btn, refresh_btn]:
            button.setMinimumHeight(36)
            actions.addWidget(button)
        set_button_variant(start_all_btn, "primary")
        set_button_variant(stop_all_btn, "danger")
        actions.addStretch(1)
        self.layout.addWidget(action_strip)

        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        container = QWidget(scroll)
        self.monitor_rows_layout = QVBoxLayout(container)
        self.monitor_rows_layout.setContentsMargins(0, 0, 0, 0)
        self.monitor_rows_layout.setSpacing(12)
        scroll.setWidget(container)
        self.layout.addWidget(scroll, 1)

    @staticmethod
    def _monitor_id(index: int) -> str:
        return f"folder_{index}"

    def load_monitor_config(self) -> dict:
        config_file = Path("settings/folder_monitor_config.json")
        if config_file.exists():
            try:
                return json.loads(config_file.read_text(encoding="utf-8"))
            except Exception:
                return {}
        return {}

    def _resolve_monitor_config(self, key: str) -> dict | None:
        raw = self.load_monitor_config().get(key, {})
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
            }
            self.monitor_runtimes[key] = runtime
        return runtime

    def reload_monitors(self) -> None:
        self.config = load_config()
        self._rebuild_monitor_rows()

    def _rebuild_monitor_rows(self) -> None:
        assert self.monitor_rows_layout is not None
        while self.monitor_rows_layout.count():
            item = self.monitor_rows_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        self.monitor_cards.clear()
        self.log_widgets.clear()
        config = self.load_monitor_config()

        for index in range(1, 4):
            key = self._monitor_id(index)
            raw = config.get(key, {})
            runtime = self._get_runtime(key)
            enabled = bool(raw.get("enabled", False))
            path_value = str(raw.get("path", "")).strip() or "(未配置路径)"
            alias_value = str(raw.get("webhook_alias", "")).strip() or "(未配置机器人)"

            row = QWidget(self)
            row.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
            row_layout = QHBoxLayout(row)
            row_layout.setContentsMargins(0, 0, 0, 0)
            row_layout.setSpacing(12)

            card = QFrame(row)
            card.setObjectName("PanelCard")
            card.setMinimumWidth(336)
            card.setMaximumWidth(372)
            card.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
            card_layout = QVBoxLayout(card)
            card_layout.setContentsMargins(14, 12, 14, 12)
            card_layout.setSpacing(8)

            header = QHBoxLayout()
            header.setSpacing(6)
            title = QLabel(f"检测 {index}", card)
            title.setStyleSheet("font-size: 14px; font-weight: 700; color: #0F172A;")
            enabled_label = create_status_pill(card, "配置已启用" if enabled else "配置未启用", "success" if enabled else "neutral")
            status_label = create_status_pill(card, "运行中" if runtime.get("is_running") else "待机中", "success" if runtime.get("is_running") else "neutral")
            action_btn = QPushButton("停止" if runtime.get("is_running") else "启动", card)
            action_btn.setMinimumHeight(30)
            action_btn.setMinimumWidth(64)
            action_btn.clicked.connect(lambda _checked=False, target=key: self.toggle_single_monitor(target))
            set_button_variant(action_btn, "danger" if runtime.get("is_running") else "primary")
            enabled_label.setStyleSheet("font-size: 11px; padding: 1px 10px; min-height: 22px;")
            status_label.setStyleSheet("font-size: 11px; padding: 1px 10px; min-height: 22px;")
            header.addWidget(title)
            header.addWidget(enabled_label)
            header.addStretch(1)
            header.addWidget(status_label)
            header.addWidget(action_btn)
            card_layout.addLayout(header)

            summary_wrap = QFrame(card)
            summary_wrap.setStyleSheet("background: transparent; border: none;")
            summary_wrap.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
            summary_wrap_layout = QGridLayout(summary_wrap)
            summary_wrap_layout.setContentsMargins(0, 2, 0, 2)
            summary_wrap_layout.setHorizontalSpacing(6)
            summary_wrap_layout.setVerticalSpacing(6)
            summary_wrap_layout.setColumnStretch(0, 1)
            summary_wrap_layout.setColumnStretch(1, 1)
            summary_wrap_layout.setRowStretch(0, 0)

            for item_index, value_text in enumerate([path_value, alias_value]):
                info_card = QFrame(summary_wrap)
                info_card.setStyleSheet("background: #F7FAFE; border: none; border-radius: 10px;")
                info_card.setMinimumHeight(52)
                info_card.setMaximumHeight(64)
                info_layout = QVBoxLayout(info_card)
                info_layout.setContentsMargins(10, 8, 10, 8)
                info_layout.setSpacing(0)
                value_label = QLabel(value_text, info_card)
                value_label.setWordWrap(True)
                value_label.setMaximumHeight(42)
                value_label.setStyleSheet("color: #17304B; font-size: 11px; font-weight: 600;")
                info_layout.addWidget(value_label)
                summary_wrap_layout.addWidget(info_card, 0, item_index)
                summary_wrap_layout.setAlignment(info_card, Qt.AlignTop)
            card_layout.addWidget(summary_wrap, 0, Qt.AlignTop)
            card_layout.addStretch(1)
            target_height = max(112, card.sizeHint().height())

            log_wrap = QFrame(row)
            log_wrap.setObjectName("PanelCard")
            log_wrap.setMinimumHeight(target_height)
            log_wrap.setMaximumHeight(target_height)
            log_wrap.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            log_layout = QVBoxLayout(log_wrap)
            log_layout.setContentsMargins(14, 14, 14, 14)
            log_layout.setSpacing(8)
            log_text = QTextEdit(log_wrap)
            log_text.setReadOnly(True)
            log_text.setProperty("logView", True)
            log_text.setPlainText("".join(self.log_buffers.get(key, [])))
            log_layout.addWidget(log_text, 1)

            control_wrap = QFrame(row)
            control_wrap.setObjectName("ActionStrip")
            control_wrap.setMinimumWidth(96)
            control_wrap.setMaximumWidth(110)
            control_wrap.setMinimumHeight(target_height)
            control_wrap.setMaximumHeight(target_height)
            control_wrap.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
            control_layout = QVBoxLayout(control_wrap)
            control_layout.setContentsMargins(10, 10, 10, 10)
            control_layout.setSpacing(8)
            auto_scroll_box = QCheckBox("自动滚动", control_wrap)
            auto_scroll_box.setChecked(bool(self.config.auto_scroll_log))
            export_btn = QPushButton("导出", control_wrap)
            clear_btn = QPushButton("清空", control_wrap)
            export_btn.clicked.connect(lambda _checked=False, target=key: self.export_logs_for_monitor(target))
            clear_btn.clicked.connect(lambda _checked=False, target=key: self.clear_logs_for_monitor(target))
            set_button_variant(export_btn, "warn")
            control_layout.addWidget(auto_scroll_box)
            control_layout.addWidget(export_btn)
            control_layout.addWidget(clear_btn)
            control_layout.addStretch(1)

            row_layout.addWidget(card, 0, Qt.AlignTop)
            row_layout.addWidget(log_wrap, 1)
            row_layout.addWidget(control_wrap)

            self.monitor_cards[key] = {
                "status_label": status_label,
                "action_btn": action_btn,
                "auto_scroll": auto_scroll_box,
            }
            self.log_widgets[key] = log_text
            self.monitor_rows_layout.addWidget(row)

        self.monitor_rows_layout.addStretch(1)

    def _append_log_entry(self, level: str, message: str, source: str) -> None:
        icons = {"INFO": "ℹ️", "SUCCESS": "✅", "WARNING": "⚠️", "ERROR": "❌"}
        line = f"[{time.strftime('%H:%M:%S')}] {icons.get(level, '•')} {message}\n"
        self.log_buffers.setdefault("global", []).append(line)
        self.log_buffers.setdefault(source, []).append(line)
        widget = self.log_widgets.get(source)
        if widget:
            widget.moveCursor(QTextCursor.End)
            widget.insertPlainText(line)
            if self._auto_scroll_enabled(source):
                widget.moveCursor(QTextCursor.End)
        try:
            self.log_bus.emit(level, message, source=source)
        except Exception:
            pass

    def _emit_log(self, level: str, message: str, source: str) -> None:
        self.log_signal.emit(level, message, source)

    def _auto_scroll_enabled(self, key: str) -> bool:
        card = self.monitor_cards.get(key, {})
        checkbox = card.get("auto_scroll")
        return bool(checkbox.isChecked()) if isinstance(checkbox, QCheckBox) else True

    def clear_logs_for_monitor(self, key: str) -> None:
        widget = self.log_widgets.get(key)
        if isinstance(widget, QTextEdit):
            widget.clear()
            self.log_buffers[key] = []

    def export_logs_for_monitor(self, key: str) -> None:
        widget = self.log_widgets.get(key)
        if not isinstance(widget, QTextEdit):
            return
        content = widget.toPlainText().strip()
        if not content:
            QMessageBox.warning(self, "提示", "当前日志为空，暂无可导出的内容")
            return
        export_dir = Path("exports")
        export_dir.mkdir(parents=True, exist_ok=True)
        default_name = f"folder_{key}_{time.strftime('%Y%m%d_%H%M%S')}.log"
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "导出监测日志",
            str((export_dir / default_name).resolve()),
            "Log Files (*.log);;Text Files (*.txt);;All Files (*.*)",
        )
        if not file_path:
            return
        Path(file_path).write_text(content + "\n", encoding="utf-8")
        QMessageBox.information(self, "导出成功", f"日志已导出到:\n{file_path}")

    def _apply_monitor_state(self, key: str, text: str, color: str, running: bool) -> None:
        _ = color
        card = self.monitor_cards.get(key, {})
        status_label = card.get("status_label")
        action_btn = card.get("action_btn")
        if isinstance(status_label, QLabel):
            status_label.setText(text)
            status_label.setProperty("pillTone", "success" if running else ("warning" if text == "停止中" else "neutral"))
            status_label.style().unpolish(status_label)
            status_label.style().polish(status_label)
            status_label.update()
        if isinstance(action_btn, QPushButton):
            if text == "停止中":
                action_btn.setText("停止中")
                action_btn.setEnabled(False)
                set_button_variant(action_btn, "warn")
            else:
                action_btn.setText("停止" if running else "启动")
                action_btn.setEnabled(True)
                set_button_variant(action_btn, "danger" if running else "primary")

    def toggle_single_monitor(self, key: str) -> None:
        runtime = self._get_runtime(key)
        if runtime.get("is_running"):
            self.stop_single_monitor(key)
        else:
            self.start_single_monitor(key)

    def start_single_monitor(self, key: str) -> None:
        runtime = self._get_runtime(key)
        if runtime.get("is_running"):
            return
        config = self._resolve_monitor_config(key)
        if not config:
            self._emit_log("ERROR", f"{key} 未找到有效配置", key)
            return
        if not config.get("enabled", False):
            QMessageBox.warning(self, "提示", f"{key} 未启用，请先在设置页启用该监测项")
            return

        folder_path = config["path"]
        webhook_url = str(config.get("webhook_url", "")).strip()
        webhook_alias = str(config.get("webhook_alias", "")).strip()
        if not folder_path.exists():
            self._emit_log("ERROR", f"文件夹不存在: {folder_path}", key)
            return
        if not webhook_url or not webhook_url.startswith("http"):
            self._emit_log("ERROR", "推送机器人配置无效", key)
            return

        observer = Observer()
        executor = ThreadPoolExecutor(max_workers=2, thread_name_prefix=f"{key}-monitor")
        with runtime["lock"]:
            runtime["inflight_events"].clear()
        handler = FolderMonitorHandler(
            callback=lambda path, event, monitor_key=key: self.on_file_event(monitor_key, path, event),
            log_handler=_QtFolderLogAdapter(self, key),
            source=key,
        )
        observer.schedule(handler, str(folder_path), recursive=False)
        observer.start()

        runtime["observer"] = observer
        runtime["executor"] = executor
        runtime["is_running"] = True
        runtime["config"] = config
        self.monitor_config_markers[key] = self._config_marker(config)
        self.state_signal.emit(key, "运行中", "#15803D", True)
        alias_text = f" (别名: {webhook_alias})" if webhook_alias else ""
        self._emit_log("INFO", f"启动监测: {folder_path}{alias_text}", key)
        self.scan_existing_files(key)

    def stop_single_monitor(self, key: str) -> None:
        runtime = self._get_runtime(key)
        if not runtime.get("is_running"):
            return
        self._emit_log("INFO", "正在停止监测...", key)
        runtime["is_running"] = False
        observer = runtime.get("observer")
        if observer:
            observer.stop()
            observer.join(timeout=self.OBSERVER_JOIN_TIMEOUT_SECONDS)
            if observer.is_alive():
                self._emit_log("WARNING", f"监测线程仍在退出中（>{self.OBSERVER_JOIN_TIMEOUT_SECONDS}秒），请稍候", key)
        executor = runtime.get("executor")
        if executor:
            executor.shutdown(wait=False, cancel_futures=True)
        with runtime["lock"]:
            runtime["inflight_events"].clear()
        runtime["observer"] = None
        runtime["executor"] = None
        runtime["config"] = None
        self.state_signal.emit(key, "已停止", "#64748B", False)
        self._emit_log("SUCCESS", "监测已停止", key)

    def start_all_monitors(self) -> None:
        for index in range(1, 4):
            self.start_single_monitor(self._monitor_id(index))

    def stop_all_monitors(self) -> None:
        for index in range(1, 4):
            self.stop_single_monitor(self._monitor_id(index))

    def on_file_event(self, key: str, file_path: str, event_type: str) -> None:
        runtime = self._get_runtime(key)
        if not runtime.get("is_running"):
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

    def _process_file_event(self, key: str, file_path: str, event_type: str, webhook_url: str, event_key: str) -> None:
        runtime = self._get_runtime(key)
        try:
            if event_type == "modified":
                time.sleep(1)
            path = Path(file_path)
            if not path.exists() or not path.is_file():
                self._emit_log("WARNING", f"文件不存在或不可用，跳过: {path}", key)
                return
            if not webhook_url:
                self._emit_log("ERROR", f"Webhook URL为空，无法处理文件: {path.name}", key)
                return
            if self.tracker.is_sent(path, webhook_url):
                self._emit_log("INFO", f"文件已处理，跳过: {path.name}", key)
                return
            self._emit_log("INFO", f"处理文件: {path.name} ({path.suffix}) [{event_type}]", key)
            file_id = send_file_via_webhook(
                path,
                webhook_url,
                event_callback=lambda level, message, src=key: self._emit_log(level, message, src),
            )
            self.tracker.mark_sent(path, webhook_url, file_id)
        except Exception as exc:
            self._emit_log("ERROR", f"处理失败: {exc}", key)
            self.state_signal.emit(key, "最近出错", "#B91C1C", runtime.get("is_running", False))
        finally:
            with runtime["lock"]:
                runtime["inflight_events"].discard(event_key)

    def scan_existing_files(self, key: str) -> None:
        runtime = self._get_runtime(key)
        config = runtime.get("config") or {}
        folder_path = config.get("path")
        webhook_url = str(config.get("webhook_url", "")).strip()
        if not folder_path or not webhook_url:
            return
        self._emit_log("INFO", "扫描现有文件...", key)
        try:
            for file_path in Path(folder_path).iterdir():
                if file_path.is_file() and not self.tracker.is_sent(file_path, webhook_url):
                    self._emit_log("INFO", f"发现未处理文件: {file_path.name} ({file_path.suffix})", key)
                    self.on_file_event(key, str(file_path), "existing")
        except Exception as exc:
            self._emit_log("ERROR", f"扫描文件夹失败 {folder_path}: {exc}", key)

    def apply_runtime_config_updates(self) -> None:
        for index in range(1, 4):
            key = self._monitor_id(index)
            runtime = self._get_runtime(key)
            new_config = self._resolve_monitor_config(key)
            new_marker = self._config_marker(new_config)
            old_marker = self.monitor_config_markers.get(key, ("", "", "", False))
            if not runtime.get("is_running"):
                self.monitor_config_markers[key] = new_marker
                continue
            if new_marker == old_marker:
                continue
            self._emit_log("INFO", "检测配置已热刷新", key)
            self.monitor_config_markers[key] = new_marker
            if new_config:
                runtime["config"] = new_config

    def on_page_activated(self) -> None:
        self.reload_monitors()

    def on_external_config_updated(self) -> None:
        self.reload_monitors()
        self.apply_runtime_config_updates()


class _QtFolderLogAdapter:
    def __init__(self, page: FolderMonitorPage, source: str) -> None:
        self.page = page
        self.source = source

    def log(self, level: str, message: str, source: str = "global"):
        self.page._emit_log(level, message, self.source)

    def info(self, message: str, source: str = "global"):
        self.page._emit_log("INFO", message, self.source)

    def success(self, message: str, source: str = "global"):
        self.page._emit_log("SUCCESS", message, self.source)

    def warning(self, message: str, source: str = "global"):
        self.page._emit_log("WARNING", message, self.source)

    def error(self, message: str, source: str = "global"):
        self.page._emit_log("ERROR", message, self.source)
