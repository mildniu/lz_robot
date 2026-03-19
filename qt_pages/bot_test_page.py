from __future__ import annotations

import threading
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Signal
from PySide6.QtGui import QTextCursor
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
    QComboBox,
)

from app_services import load_webhook_aliases
from mail_forwarder.processing_service import build_webhook_client, send_file_via_webhook
from qt_components import set_button_variant

from .base import BasePage


class BotTestPage(BasePage):
    log_signal = Signal(str, str)
    sending_state_signal = Signal(bool)

    NO_ALIAS_LABEL = "(未配置)"

    def __init__(self, log_bus) -> None:
        super().__init__(log_bus, "机器人测试")
        self.alias_map: dict[str, str] = {}
        self._sending = False

        self.alias_combo: QComboBox | None = None
        self.text_input: QLineEdit | None = None
        self.file_input: QLineEdit | None = None
        self.log_text: QTextEdit | None = None
        self.send_text_btn: QPushButton | None = None
        self.send_file_btn: QPushButton | None = None

        self.log_signal.connect(self._append_log)
        self.sending_state_signal.connect(self._apply_sending_state)

        self._build_ui()
        self.refresh_aliases(log_result=False)

    def _build_ui(self) -> None:
        alias_card = QFrame(self)
        alias_card.setObjectName("ActionStrip")
        alias_layout = QHBoxLayout(alias_card)
        alias_layout.setContentsMargins(14, 12, 14, 12)
        alias_layout.setSpacing(10)
        label = QLabel("推送机器人", self)
        label.setObjectName("FieldLabel")
        alias_layout.addWidget(label)
        self.alias_combo = QComboBox(self)
        self.alias_combo.setMinimumWidth(280)
        self.alias_combo.setMaximumWidth(360)
        alias_layout.addWidget(self.alias_combo)
        refresh_btn = QPushButton("重新读取", self)
        set_button_variant(refresh_btn, "warn")
        refresh_btn.clicked.connect(self.refresh_aliases)
        alias_layout.addWidget(refresh_btn)
        alias_layout.addStretch(1)
        self.layout.addWidget(alias_card)

        text_card = QFrame(self)
        text_card.setObjectName("PanelCard")
        text_layout = QVBoxLayout(text_card)
        text_layout.setContentsMargins(18, 16, 18, 16)
        text_layout.setSpacing(10)
        text_title = QLabel("文字测试", self)
        text_title.setObjectName("SectionTitle")
        text_layout.addWidget(text_title)
        self.text_input = QLineEdit(self)
        self.text_input.setPlaceholderText("输入测试文本")
        self.text_input.setText("这是一条机器人文字测试消息。")
        text_layout.addWidget(self.text_input)
        self.send_text_btn = QPushButton("发送文字测试", self)
        set_button_variant(self.send_text_btn, "primary")
        self.send_text_btn.clicked.connect(self.send_text_test)
        text_layout.addWidget(self.send_text_btn)
        self.layout.addWidget(text_card)

        file_card = QFrame(self)
        file_card.setObjectName("PanelCard")
        file_layout = QVBoxLayout(file_card)
        file_layout.setContentsMargins(18, 16, 18, 16)
        file_layout.setSpacing(10)
        file_title = QLabel("文件/图片测试", self)
        file_title.setObjectName("SectionTitle")
        file_layout.addWidget(file_title)
        file_row = QHBoxLayout()
        file_row.setSpacing(8)
        self.file_input = QLineEdit(self)
        self.file_input.setPlaceholderText("请选择本地文件路径")
        file_row.addWidget(self.file_input, 1)
        browse_btn = QPushButton("选择文件", self)
        set_button_variant(browse_btn, "warn")
        browse_btn.clicked.connect(self.choose_file)
        file_row.addWidget(browse_btn)
        file_layout.addLayout(file_row)
        self.send_file_btn = QPushButton("发送文件/图片测试", self)
        set_button_variant(self.send_file_btn, "primary")
        self.send_file_btn.clicked.connect(self.send_file_test)
        file_layout.addWidget(self.send_file_btn)
        self.layout.addWidget(file_card)

        log_card = QFrame(self)
        log_card.setObjectName("PanelCard")
        log_layout = QVBoxLayout(log_card)
        log_layout.setContentsMargins(18, 16, 18, 16)
        log_layout.setSpacing(10)
        log_title = QLabel("测试日志", self)
        log_title.setObjectName("SectionTitle")
        log_layout.addWidget(log_title)
        log_actions = QHBoxLayout()
        clear_btn = QPushButton("清空日志", self)
        set_button_variant(clear_btn, "warn")
        clear_btn.clicked.connect(self.clear_logs)
        export_btn = QPushButton("导出日志", self)
        set_button_variant(export_btn, "warn")
        export_btn.clicked.connect(self.export_logs)
        log_actions.addWidget(clear_btn)
        log_actions.addWidget(export_btn)
        log_actions.addStretch(1)
        log_layout.addLayout(log_actions)
        self.log_text = QTextEdit(self)
        self.log_text.setReadOnly(True)
        self.log_text.setProperty("logView", True)
        log_layout.addWidget(self.log_text, 1)
        self.layout.addWidget(log_card, 1)

    def refresh_aliases(self, log_result: bool = True) -> None:
        config = load_webhook_aliases()
        aliases = config.get("aliases", {})
        self.alias_map = {k: v.strip() for k, v in aliases.items() if str(k).strip() and str(v).strip()}

        values = [self.NO_ALIAS_LABEL] + sorted(self.alias_map.keys())
        assert self.alias_combo is not None
        current = self.alias_combo.currentText().strip()
        self.alias_combo.blockSignals(True)
        self.alias_combo.clear()
        self.alias_combo.addItems(values)
        if current in values:
            self.alias_combo.setCurrentText(current)
        elif len(values) > 1:
            self.alias_combo.setCurrentIndex(1)
        else:
            self.alias_combo.setCurrentIndex(0)
        self.alias_combo.blockSignals(False)
        if log_result:
            self.log_signal.emit("INFO", f"已加载机器人别名数量: {len(self.alias_map)}")

    def choose_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(self, "选择测试文件", "", "所有文件 (*.*)")
        if file_path and self.file_input:
            self.file_input.setText(file_path)

    def _selected_webhook(self) -> tuple[str, str]:
        assert self.alias_combo is not None
        alias = self.alias_combo.currentText().strip()
        if not alias or alias == self.NO_ALIAS_LABEL:
            raise ValueError("请先选择机器人别名")
        webhook_url = self.alias_map.get(alias, "").strip()
        if not webhook_url:
            raise ValueError(f"机器人别名无效: {alias}")
        if not webhook_url.startswith(("http://", "https://")):
            raise ValueError(f"Webhook 地址格式无效: {webhook_url}")
        return alias, webhook_url

    def _run_async(self, action: str, worker) -> None:
        if self._sending:
            self.log_signal.emit("WARNING", "当前有测试任务执行中，请稍后")
            return

        self.sending_state_signal.emit(True)
        self.log_signal.emit("INFO", f"开始执行: {action}")

        def task() -> None:
            try:
                worker()
            except Exception as exc:
                self.log_signal.emit("ERROR", f"{action}失败: {exc}")
            finally:
                self.sending_state_signal.emit(False)

        threading.Thread(target=task, daemon=True).start()

    def send_text_test(self) -> None:
        assert self.text_input is not None
        content = self.text_input.text().strip()
        if not content:
            self.log_signal.emit("WARNING", "文字内容为空，请先填写测试文本")
            return
        try:
            alias, webhook_url = self._selected_webhook()
        except Exception as exc:
            self.log_signal.emit("ERROR", str(exc))
            return

        def worker() -> None:
            self.log_signal.emit("INFO", f"目标机器人: {alias}")
            webhook = build_webhook_client(webhook_url)
            webhook.send_text_alert(content)
            self.log_signal.emit("SUCCESS", "文字消息发送成功")

        self._run_async("文字测试", worker)

    def send_file_test(self) -> None:
        assert self.file_input is not None
        raw_path = self.file_input.text().strip()
        if not raw_path:
            self.log_signal.emit("WARNING", "请先选择测试文件")
            return
        file_path = Path(raw_path)
        if not file_path.exists() or not file_path.is_file():
            self.log_signal.emit("ERROR", f"文件不存在: {file_path}")
            return
        try:
            alias, webhook_url = self._selected_webhook()
        except Exception as exc:
            self.log_signal.emit("ERROR", str(exc))
            return

        def worker() -> None:
            self.log_signal.emit("INFO", f"目标机器人: {alias}")
            self.log_signal.emit("INFO", f"发送文件: {file_path.name}")
            send_file_via_webhook(
                file_path,
                webhook_url,
                event_callback=lambda level, message: self.log_signal.emit(level, message),
            )
            self.log_signal.emit("SUCCESS", "文件/图片测试发送完成")

        self._run_async("文件/图片测试", worker)

    def _append_log(self, level: str, message: str) -> None:
        icons = {"INFO": "ℹ️", "SUCCESS": "✅", "WARNING": "⚠️", "ERROR": "❌"}
        line = f"[{datetime.now().strftime('%H:%M:%S')}] {icons.get(level, '•')} {message}\n"
        assert self.log_text is not None
        self.log_text.moveCursor(QTextCursor.End)
        self.log_text.insertPlainText(line)
        self.log_text.moveCursor(QTextCursor.End)
        try:
            self.log_bus.emit(level, message, source="bot_test")
        except Exception:
            pass

    def clear_logs(self) -> None:
        if self.log_text:
            self.log_text.clear()

    def export_logs(self) -> None:
        if not self.log_text:
            return
        content = self.log_text.toPlainText().strip()
        if not content:
            QMessageBox.warning(self, "提示", "当前日志为空，暂无可导出的内容")
            return
        export_dir = Path("exports")
        export_dir.mkdir(parents=True, exist_ok=True)
        default_name = f"bot_test_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "导出测试日志",
            str((export_dir / default_name).resolve()),
            "Log Files (*.log);;Text Files (*.txt);;All Files (*.*)",
        )
        if not file_path:
            return
        Path(file_path).write_text(content + "\n", encoding="utf-8")
        QMessageBox.information(self, "导出成功", f"日志已导出到:\n{file_path}")

    def _apply_sending_state(self, sending: bool) -> None:
        self._sending = sending
        state = not sending
        if self.send_text_btn:
            self.send_text_btn.setEnabled(state)
        if self.send_file_btn:
            self.send_file_btn.setEnabled(state)

    def on_page_activated(self) -> None:
        self.refresh_aliases(log_result=False)

    def on_external_config_updated(self) -> None:
        self.refresh_aliases(log_result=True)
