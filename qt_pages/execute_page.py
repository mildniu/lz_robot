from __future__ import annotations

import threading
import time
from datetime import datetime, timedelta
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

from mail_forwarder import MailProcessingService, load_config
from mail_forwarder.subject_attachment_rules import list_enabled_rules_with_slots
from qt_components import create_status_pill, set_button_variant

from .base import BasePage


class ExecutePage(BasePage):
    log_signal = Signal(str, str, str)
    state_signal = Signal(str, str, str, bool)
    result_signal = Signal(str, dict)
    structure_refresh_signal = Signal()

    def __init__(self, log_bus) -> None:
        super().__init__(log_bus, "邮件检测")
        self.config = load_config()
        self.rule_configs: dict[str, dict] = {}
        self.rule_cards: dict[str, dict[str, object]] = {}
        self.rule_runtimes: dict[str, dict] = {}
        self.log_buffers: dict[str, list[str]] = {"global": []}
        self.log_widgets: dict[str, QTextEdit] = {}
        self.rule_rows_layout: QVBoxLayout | None = None

        self.log_signal.connect(self._append_log_entry)
        self.state_signal.connect(self._apply_rule_state)
        self.result_signal.connect(self._apply_rule_result)
        self.structure_refresh_signal.connect(self.reload_rules)

        self._build_ui()
        self.reload_rules()

    def _build_ui(self) -> None:
        action_strip = QFrame(self)
        action_strip.setObjectName("ActionStrip")
        actions = QHBoxLayout(action_strip)
        actions.setContentsMargins(14, 12, 14, 12)
        actions.setSpacing(8)

        start_all_btn = QPushButton("全部启动", self)
        start_all_btn.clicked.connect(self.start_all_rules)
        stop_all_btn = QPushButton("全部停止", self)
        stop_all_btn.clicked.connect(self.stop_all_rules)
        test_all_btn = QPushButton("全部测试", self)
        test_all_btn.clicked.connect(self.test_all_rules)
        refresh_btn = QPushButton("刷新规则", self)
        refresh_btn.clicked.connect(self.reload_rules)

        for button in [start_all_btn, stop_all_btn, test_all_btn, refresh_btn]:
            button.setMinimumHeight(36)
            actions.addWidget(button)
        set_button_variant(start_all_btn, "primary")
        set_button_variant(stop_all_btn, "danger")
        set_button_variant(test_all_btn, "warn")
        actions.addStretch(1)
        self.layout.addWidget(action_strip)

        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        container = QWidget(scroll)
        self.rule_rows_layout = QVBoxLayout(container)
        self.rule_rows_layout.setContentsMargins(0, 0, 0, 0)
        self.rule_rows_layout.setSpacing(12)
        scroll.setWidget(container)
        self.layout.addWidget(scroll, 1)

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

    @staticmethod
    def format_wait_text(seconds: int) -> str:
        if seconds % 60 == 0 and seconds >= 60:
            return f"{seconds // 60} 分钟"
        return f"{seconds} 秒"

    def _rule_summary_items(self, rule: dict) -> list[str]:
        mailbox_alias = self._display_value(rule.get("mailbox_alias", ""), "未选邮箱")
        webhook_alias = self._display_value(rule.get("webhook_alias", ""), "未选机器人")
        interval_seconds = int(rule.get("poll_interval_seconds", 60) or 60)
        trigger_mode = str(rule.get("trigger_mode", "periodic")).strip() or "periodic"
        schedule_time = str(rule.get("schedule_time", "")).strip()
        mode = "脚本处理" if str(rule.get("script_path", "")).strip() else "直接推送"
        trigger_text = f"定时 {schedule_time or '--:--'}" if trigger_mode == "timed" else self.format_wait_text(interval_seconds)
        return [
            mailbox_alias,
            webhook_alias,
            trigger_text,
            mode,
        ]

    def _format_result_text(self, payload: dict | None) -> str:
        if not payload:
            return "结果: 暂无记录"
        status_map = {"processed": "成功", "skipped": "跳过", "not_found": "未找到", "error": "失败"}
        status = status_map.get(str(payload.get("status", "")), str(payload.get("status", "")))
        base = f"结果: {status} | 时间: {payload.get('time', '--')} | UID: {payload.get('uid', '-') or '-'} | 附件: {payload.get('file_count', 0)}"
        reason = str(payload.get("reason", "")).strip()
        return f"{base} | 说明: {reason}" if reason else base

    @staticmethod
    def _result_color(status: str) -> str:
        return {
            "processed": "#166534",
            "skipped": "#A16207",
            "not_found": "#475569",
            "error": "#B91C1C",
        }.get(status, "#475569")

    @staticmethod
    def _seconds_until_daily_time(schedule_time: str) -> tuple[int, str]:
        hour_text, minute_text = schedule_time.split(":", 1)
        hour = int(hour_text)
        minute = int(minute_text)
        now = datetime.now()
        target = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
        if now > target:
            target += timedelta(days=1)
        seconds = max(0, int((target - now).total_seconds()))
        return seconds, target.strftime("%Y-%m-%d %H:%M")

    def reload_rules(self) -> None:
        self.config = load_config()
        enabled_rule_map = {
            self._build_rule_id(slot_index): rule
            for slot_index, rule in list_enabled_rules_with_slots()
        }
        self.rule_configs = enabled_rule_map

        for rule_id in list(self.rule_runtimes.keys()):
            if rule_id not in enabled_rule_map:
                runtime = self.rule_runtimes.get(rule_id, {})
                runtime.setdefault("stop_event", threading.Event()).set()

        self._rebuild_rule_rows()

    def _rebuild_rule_rows(self) -> None:
        assert self.rule_rows_layout is not None
        while self.rule_rows_layout.count():
            item = self.rule_rows_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        self.rule_cards.clear()
        self.log_widgets.clear()

        for rule_id, rule in self.rule_configs.items():
            slot_index = self._slot_from_rule_id(rule_id)
            runtime = self.rule_runtimes.setdefault(
                rule_id,
                {
                    "thread": None,
                    "stop_event": threading.Event(),
                    "is_running": False,
                    "last_result": None,
                    "test_running": False,
                },
            )

            row = QWidget(self)
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
            title = QLabel(self._display_value(rule.get("keyword", ""), f"规则 {slot_index}"), card)
            title.setStyleSheet("font-size: 14px; font-weight: 700; color: #0F172A;")
            status = create_status_pill(card, "运行中" if runtime.get("is_running") else "待机中", "success" if runtime.get("is_running") else "neutral")
            slot_label = create_status_pill(card, f"规则 {slot_index}", "info")
            action_btn = QPushButton("停止" if runtime.get("is_running") else "启动", card)
            action_btn.setMinimumHeight(30)
            action_btn.setMinimumWidth(64)
            action_btn.clicked.connect(lambda _checked=False, target=rule_id: self.toggle_single_rule(target))
            set_button_variant(action_btn, "danger" if runtime.get("is_running") else "primary")
            test_btn = QPushButton("测试", card)
            test_btn.setMinimumHeight(30)
            test_btn.setMinimumWidth(58)
            test_btn.clicked.connect(lambda _checked=False, target=rule_id: self.test_single_rule(target))
            set_button_variant(test_btn, "warn")
            test_btn.setEnabled(not bool(runtime.get("is_running")))
            slot_label.setStyleSheet("font-size: 11px; padding: 1px 10px; min-height: 22px;")
            status.setStyleSheet("font-size: 11px; padding: 1px 10px; min-height: 22px;")
            header.addWidget(title)
            header.addWidget(slot_label)
            header.addStretch(1)
            header.addWidget(status)
            header.addWidget(test_btn)
            header.addWidget(action_btn)
            card_layout.addLayout(header)

            summary_wrap = QFrame(card)
            summary_wrap.setStyleSheet("background: transparent; border: none;")
            summary_wrap_layout = QGridLayout(summary_wrap)
            summary_wrap_layout.setContentsMargins(0, 2, 0, 2)
            summary_wrap_layout.setHorizontalSpacing(6)
            summary_wrap_layout.setVerticalSpacing(6)
            summary_wrap_layout.setColumnStretch(0, 1)
            summary_wrap_layout.setColumnStretch(1, 1)
            for item_index, value_text in enumerate(self._rule_summary_items(rule)):
                info_card = QFrame(summary_wrap)
                info_card.setStyleSheet("background: #F7FAFE; border: none; border-radius: 10px;")
                info_layout = QVBoxLayout(info_card)
                info_layout.setContentsMargins(10, 8, 10, 8)
                info_layout.setSpacing(0)
                value_label = QLabel(value_text, info_card)
                value_label.setWordWrap(True)
                value_label.setStyleSheet("color: #17304B; font-size: 11px; font-weight: 600;")
                info_layout.addWidget(value_label)
                summary_wrap_layout.addWidget(info_card, item_index // 2, item_index % 2)
            card_layout.addWidget(summary_wrap)

            result = QLabel(self._format_result_text(runtime.get("last_result")), card)
            result.setWordWrap(False)
            result.setMinimumHeight(34)
            result.setMaximumHeight(34)
            result.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            result_status = str((runtime.get("last_result") or {}).get("status", ""))
            result.setStyleSheet(
                f"color: {self._result_color(result_status)}; font-size: 11px; background: #F8FBFF; border-radius: 10px; padding: 6px 10px;"
            )
            card_layout.addWidget(result)
            target_height = max(120, card.sizeHint().height())

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
            log_text.setPlainText("".join(self.log_buffers.get(rule_id, [])))
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
            export_btn.clicked.connect(lambda _checked=False, target=rule_id: self.export_logs_for_rule(target))
            clear_btn.clicked.connect(lambda _checked=False, target=rule_id: self.clear_logs_for_rule(target))
            set_button_variant(export_btn, "warn")
            control_layout.addWidget(auto_scroll_box)
            control_layout.addWidget(export_btn)
            control_layout.addWidget(clear_btn)
            control_layout.addStretch(1)

            row_layout.addWidget(card, 0, Qt.AlignTop)
            row_layout.addWidget(log_wrap, 1)
            row_layout.addWidget(control_wrap)

            self.rule_cards[rule_id] = {
                "status_label": status,
                "result_label": result,
                "action_btn": action_btn,
                "test_btn": test_btn,
                "auto_scroll": auto_scroll_box,
            }
            self.log_widgets[rule_id] = log_text
            self.rule_rows_layout.addWidget(row)

        self.rule_rows_layout.addStretch(1)

    def _append_log_entry(self, level: str, message: str, source: str) -> None:
        icons = {"INFO": "ℹ️", "SUCCESS": "✅", "WARNING": "⚠️", "ERROR": "❌"}
        line = f"[{datetime.now().strftime('%H:%M:%S')}] {icons.get(level, '•')} {message}\n"
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

    def _apply_rule_state(self, rule_id: str, text: str, color: str, running: bool) -> None:
        _ = color
        card = self.rule_cards.get(rule_id, {})
        status = card.get("status_label")
        if isinstance(status, QLabel):
            status.setText(text)
            status.setProperty("pillTone", "success" if running else ("warning" if text == "停止中" else "neutral"))
            status.style().unpolish(status)
            status.style().polish(status)
            status.update()
        action_btn = card.get("action_btn")
        test_btn = card.get("test_btn")
        if isinstance(action_btn, QPushButton):
            if text == "停止中":
                action_btn.setText("停止中")
                action_btn.setEnabled(False)
                set_button_variant(action_btn, "warn")
            else:
                action_btn.setText("停止" if running else "启动")
                action_btn.setEnabled(True)
                set_button_variant(action_btn, "danger" if running else "primary")
        if isinstance(test_btn, QPushButton):
            test_btn.setEnabled(not running)

    def toggle_single_rule(self, rule_id: str) -> None:
        runtime = self.rule_runtimes.setdefault(rule_id, {})
        if runtime.get("is_running"):
            self.stop_single_rule(rule_id)
        else:
            self.start_single_rule(rule_id)

    def _apply_rule_result(self, rule_id: str, payload: dict) -> None:
        runtime = self.rule_runtimes.setdefault(rule_id, {})
        runtime["last_result"] = payload
        card = self.rule_cards.get(rule_id, {})
        label = card.get("result_label")
        if isinstance(label, QLabel):
            label.setText(self._format_result_text(payload))
            label.setStyleSheet(
                f"color: {self._result_color(str(payload.get('status', '')))}; font-size: 12px; background: #F8FBFF; border-radius: 12px; padding: 10px 12px;"
            )

    def _auto_scroll_enabled(self, rule_id: str) -> bool:
        card = self.rule_cards.get(rule_id, {})
        checkbox = card.get("auto_scroll")
        return bool(checkbox.isChecked()) if isinstance(checkbox, QCheckBox) else True

    def clear_logs_for_rule(self, rule_id: str) -> None:
        widget = self.log_widgets.get(rule_id)
        if isinstance(widget, QTextEdit):
            widget.clear()
            self.log_buffers[rule_id] = []

    def export_logs_for_rule(self, rule_id: str) -> None:
        widget = self.log_widgets.get(rule_id)
        if not isinstance(widget, QTextEdit):
            return
        content = widget.toPlainText().strip()
        if not content:
            QMessageBox.warning(self, "提示", "当前日志为空，暂无可导出的内容")
            return
        export_dir = Path("exports")
        export_dir.mkdir(parents=True, exist_ok=True)
        default_name = f"mail_rule_{self._slot_from_rule_id(rule_id)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "导出规则日志",
            str((export_dir / default_name).resolve()),
            "Log Files (*.log);;Text Files (*.txt);;All Files (*.*)",
        )
        if not file_path:
            return
        Path(file_path).write_text(content + "\n", encoding="utf-8")
        QMessageBox.information(self, "导出成功", f"日志已导出到:\n{file_path}")

    def _emit_log(self, level: str, message: str, source: str) -> None:
        self.log_signal.emit(level, message, source)

    def _run_rule_once(self, rule_id: str, *, force: bool, update_state: bool):
        rule = self.rule_configs.get(rule_id)
        if not rule:
            raise RuntimeError("规则不存在或已失效")
        service = MailProcessingService(load_config())
        return service.process_single_rule(
            rule,
            force=force,
            update_state=update_state,
            event_callback=lambda level, message, src=rule_id: self._emit_log(level, message, src),
        )

    def _remember_result(self, rule_id: str, result) -> None:
        payload = {
            "time": datetime.now().strftime("%H:%M:%S"),
            "status": getattr(result, "status", ""),
            "uid": getattr(result, "uid", ""),
            "file_count": len(getattr(result, "files", []) or []),
            "reason": getattr(result, "reason", ""),
        }
        self.result_signal.emit(rule_id, payload)

    def _log_rule_result(self, rule_id: str, result) -> None:
        self._remember_result(rule_id, result)
        mailbox_text = f"[{getattr(result, 'mailbox_alias', '')}] " if getattr(result, "mailbox_alias", "") else ""
        folder_text = f"(文件夹: {getattr(result, 'mailbox_folder', '')}) " if getattr(result, "mailbox_folder", "") else ""
        keyword_text = f"规则“{getattr(result, 'rule_keyword', '')}”" if getattr(result, "rule_keyword", "") else "规则"
        if result.status == "processed":
            self._emit_log("SUCCESS", f"{mailbox_text}{folder_text}{keyword_text}完成: uid={result.uid}, 附件数量={len(result.files)}", rule_id)
        elif result.status == "not_found":
            self._emit_log("INFO", f"{mailbox_text}{folder_text}{keyword_text}: {result.reason}", rule_id)
        elif result.status == "skipped":
            self._emit_log("WARNING", f"{mailbox_text}{folder_text}{keyword_text}: {result.reason}", rule_id)
        else:
            self._emit_log("ERROR", f"{mailbox_text}{folder_text}{keyword_text}: {result.reason}", rule_id)

    def start_single_rule(self, rule_id: str) -> None:
        runtime = self.rule_runtimes.get(rule_id)
        rule = self.rule_configs.get(rule_id)
        if not runtime or not rule:
            return
        if runtime.get("is_running"):
            return
        worker = runtime.get("thread")
        if worker and worker.is_alive():
            return

        stop_event = threading.Event()
        runtime["stop_event"] = stop_event
        runtime["is_running"] = True
        self.state_signal.emit(rule_id, "运行中", "#15803D", True)

        keyword = self._display_value(rule.get("keyword", ""), rule_id)
        self._emit_log("INFO", f"启动规则检测: {keyword}", rule_id)

        def runner() -> None:
            cycle_count = 0
            try:
                while not stop_event.is_set():
                    current_rule = self.rule_configs.get(rule_id)
                    if not current_rule:
                        break
                    trigger_mode = str(current_rule.get("trigger_mode", "periodic")).strip() or "periodic"

                    if trigger_mode == "timed":
                        schedule_time = str(current_rule.get("schedule_time", "")).strip()
                        if not schedule_time:
                            self._emit_log("ERROR", "定时时刻为空，规则已停止", rule_id)
                            break
                        wait_seconds, target_text = self._seconds_until_daily_time(schedule_time)
                        self._emit_log("INFO", f"下次定时执行: {target_text}", rule_id)
                        for _remaining in range(wait_seconds):
                            if stop_event.is_set():
                                break
                            time.sleep(1)
                        if stop_event.is_set():
                            break

                    cycle_count += 1
                    self._emit_log("INFO", f"第 {cycle_count} 轮开始检查", rule_id)
                    started_at = time.time()
                    try:
                        result = self._run_rule_once(rule_id, force=False, update_state=True)
                        self._log_rule_result(rule_id, result)
                    except Exception as exc:
                        self._emit_log("ERROR", f"执行出错: {exc}", rule_id)
                    finally:
                        self._emit_log("INFO", f"第 {cycle_count} 轮结束, 耗时 {time.time() - started_at:.1f} 秒", rule_id)

                    if stop_event.is_set():
                        break
                    current_rule = self.rule_configs.get(rule_id)
                    if not current_rule:
                        break
                    trigger_mode = str(current_rule.get("trigger_mode", "periodic")).strip() or "periodic"
                    if trigger_mode == "timed":
                        continue
                    wait_seconds = int(current_rule.get("poll_interval_seconds", 60) or 60)
                    self._emit_log("INFO", f"进入等待: {self.format_wait_text(wait_seconds)} 后开始下一轮", rule_id)
                    for _remaining in range(wait_seconds):
                        if stop_event.is_set():
                            break
                        time.sleep(1)
            finally:
                runtime["is_running"] = False
                runtime["thread"] = None
                self.state_signal.emit(rule_id, "已停止", "#64748B", False)
                self._emit_log("INFO", "规则线程已停止", rule_id)

        thread = threading.Thread(target=runner, daemon=True)
        runtime["thread"] = thread
        thread.start()

    def stop_single_rule(self, rule_id: str) -> None:
        runtime = self.rule_runtimes.get(rule_id)
        if not runtime:
            return
        stop_event = runtime.get("stop_event")
        if isinstance(stop_event, threading.Event):
            stop_event.set()
        runtime["is_running"] = False
        self.state_signal.emit(rule_id, "停止中", "#D97706", False)
        self._emit_log("INFO", "正在停止规则...", rule_id)

    def _run_single_test(self, rule_id: str) -> None:
        try:
            result = self._run_rule_once(rule_id, force=True, update_state=False)
            self._log_rule_result(rule_id, result)
            if result.status == "processed":
                self._emit_log("SUCCESS", "测试完成！规则已成功处理邮件", rule_id)
            else:
                self._emit_log("WARNING", "测试完成，但本次没有成功处理邮件", rule_id)
        except Exception as exc:
            self._emit_log("ERROR", f"测试过程出错: {exc}", rule_id)
        finally:
            runtime = self.rule_runtimes.get(rule_id, {})
            runtime["test_running"] = False

    def test_single_rule(self, rule_id: str) -> None:
        runtime = self.rule_runtimes.get(rule_id)
        if not runtime:
            return
        if runtime.get("is_running"):
            QMessageBox.warning(self, "提示", "规则正在运行中，请先停止后再测试")
            return
        if runtime.get("test_running"):
            return
        runtime["test_running"] = True
        self._emit_log("INFO", f"开始测试规则: {self.rule_configs.get(rule_id, {}).get('keyword', rule_id)}", rule_id)
        threading.Thread(target=lambda: self._run_single_test(rule_id), daemon=True).start()

    def start_all_rules(self) -> None:
        for rule_id in list(self.rule_configs.keys()):
            self.start_single_rule(rule_id)

    def stop_all_rules(self) -> None:
        for rule_id in list(self.rule_configs.keys()):
            self.stop_single_rule(rule_id)

    def test_all_rules(self) -> None:
        for rule_id in list(self.rule_configs.keys()):
            self.test_single_rule(rule_id)

    def on_page_activated(self) -> None:
        self.reload_rules()

    def on_external_config_updated(self) -> None:
        self.reload_rules()
