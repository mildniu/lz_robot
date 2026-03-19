from __future__ import annotations

import json
import threading
from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from app_services import load_webhook_aliases, save_webhook_aliases
from mail_forwarder import load_config
from mail_forwarder.config import upsert_env_file
from mail_forwarder.imap_client import ImapMailClient
from mail_forwarder.mailbox_store import load_mailbox_configs, save_mailbox_configs
from mail_forwarder.subject_attachment_rules import (
    load_subject_attachment_rules,
    parse_filename_keywords_input,
    parse_types_input,
    save_subject_attachment_rules,
)
from qt_components import create_field_label, create_status_pill, set_button_variant

from .base import BasePage


class SettingsPage(BasePage):
    mailbox_test_finished = Signal(int, bool, str)
    config_changed = Signal()

    ALIAS_SLOT_COUNT = 5
    MAILBOX_SLOT_COUNT = 5
    RULE_SLOT_COUNT = 5
    FOLDER_SLOT_COUNT = 3

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
    TRIGGER_VALUE_TO_LABEL = {
        "periodic": "周期检测",
        "timed": "定时检测",
    }

    def __init__(self, log_bus) -> None:
        super().__init__(log_bus, "设置")
        self.config = load_config()
        self.alias_rows: list[dict[str, object]] = []
        self.mailbox_rows: list[dict[str, object]] = []
        self.rule_cards: dict[int, dict[str, object]] = {}
        self.folder_cards: dict[int, dict[str, object]] = {}
        self.path_widgets: dict[str, object] = {}
        self.ui_widgets: dict[str, object] = {}

        self._build_ui()
        self.mailbox_test_finished.connect(self._on_mailbox_test_finished)
        self.reload_all()

    def _build_ui(self) -> None:
        tabs = QTabWidget(self)
        tabs.setDocumentMode(True)
        tabs.setUsesScrollButtons(False)
        self.layout.addWidget(tabs, 1)

        tabs.addTab(self._build_rule_tab(), "邮箱检测规则")
        tabs.addTab(self._build_mailbox_tab(), "邮箱配置")
        tabs.addTab(self._build_alias_tab(), "机器人别名")
        tabs.addTab(self._build_folder_tab(), "文件夹检测")
        tabs.addTab(self._build_path_tab(), "路径设置")
        tabs.addTab(self._build_ui_tab(), "界面设置")
        tabs.tabBar().setTabTextColor(0, QColor("#C26A12"))
        tabs.setCurrentIndex(0)

    def _build_scroll_container(self, parent: QWidget) -> tuple[QScrollArea, QWidget, QVBoxLayout]:
        scroll = QScrollArea(parent)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        container = QWidget(scroll)
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)
        scroll.setWidget(container)
        return scroll, container, layout

    def _build_action_strip(self, parent: QWidget, button_text: str, callback) -> QFrame:
        wrap = QFrame(parent)
        wrap.setObjectName("ActionStrip")
        layout = QHBoxLayout(wrap)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(8)
        button = QPushButton(button_text, parent)
        set_button_variant(button, "primary")
        button.clicked.connect(callback)
        layout.addWidget(button)
        layout.addStretch(1)
        return wrap

    def _create_rule_section(self, parent: QWidget, title: str) -> tuple[QFrame, QVBoxLayout]:
        card = QFrame(parent)
        card.setObjectName("InnerPanelCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(10)
        if str(title).strip():
            title_label = QLabel(title, card)
            title_label.setObjectName("MicroTitle")
            layout.addWidget(title_label)
        return card, layout

    def _create_form_grid(self, parent: QWidget | None = None) -> QGridLayout:
        layout = QGridLayout(parent)
        layout.setHorizontalSpacing(12)
        layout.setVerticalSpacing(10)
        layout.setColumnMinimumWidth(0, 92)
        layout.setColumnMinimumWidth(2, 92)
        layout.setColumnStretch(1, 1)
        layout.setColumnStretch(3, 1)
        return layout

    def _build_line_action_row(self, parent: QWidget, line_edit: QLineEdit, button: QPushButton) -> QWidget:
        wrap = QWidget(parent)
        layout = QHBoxLayout(wrap)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        layout.addWidget(line_edit, 1)
        layout.addWidget(button)
        return wrap

    def _toggle_rule_card(self, index: int, checked: bool) -> None:
        card = self.rule_cards.get(index, {})
        content = card.get("content")
        button = card.get("collapse_btn")
        if isinstance(content, QWidget):
            content.setVisible(bool(checked))
        if isinstance(button, QPushButton):
            button.setText("收起" if checked else "展开")

    @staticmethod
    def _add_form_row(layout: QGridLayout, row: int, left_label: str, left_widget: QWidget, right_label: str, right_widget: QWidget) -> None:
        parent = layout.parentWidget() or left_widget
        layout.addWidget(create_field_label(left_label, parent), row, 0)
        layout.addWidget(left_widget, row, 1)
        layout.addWidget(create_field_label(right_label, parent), row, 2)
        layout.addWidget(right_widget, row, 3)

    def _build_alias_tab(self) -> QWidget:
        page = QWidget(self)
        layout = QVBoxLayout(page)
        layout.setSpacing(12)
        scroll, container, container_layout = self._build_scroll_container(page)

        header_card = QFrame(container)
        header_card.setObjectName("PanelCardSoft")
        header_layout = QGridLayout(header_card)
        header_layout.setContentsMargins(16, 10, 16, 10)
        header_layout.setHorizontalSpacing(12)
        header_layout.setVerticalSpacing(0)
        header_layout.setColumnMinimumWidth(0, 88)
        header_layout.setColumnStretch(1, 0)
        header_layout.setColumnMinimumWidth(2, 88)
        header_layout.setColumnStretch(3, 1)
        header_layout.setColumnStretch(4, 1)
        header_layout.addWidget(QLabel("", header_card), 0, 0)
        header_layout.addWidget(create_field_label("别名", header_card), 0, 1)
        header_layout.addWidget(QLabel("", header_card), 0, 2)
        header_layout.addWidget(create_field_label("Webhook URL", header_card), 0, 3)
        header_layout.addWidget(QLabel("", header_card), 0, 4)
        container_layout.addWidget(header_card)

        for index in range(1, self.ALIAS_SLOT_COUNT + 1):
            card = QFrame(container)
            card.setObjectName("PanelCard")
            card_layout = QGridLayout(card)
            card_layout.setContentsMargins(16, 14, 16, 14)
            card_layout.setHorizontalSpacing(12)
            card_layout.setVerticalSpacing(10)
            card_layout.setColumnMinimumWidth(0, 88)
            card_layout.setColumnStretch(1, 0)
            card_layout.setColumnMinimumWidth(2, 88)
            card_layout.setColumnStretch(3, 1)
            badge = create_status_pill(card, f"别名 {index}", "info")
            alias_edit = QLineEdit(card)
            alias_edit.setPlaceholderText("例如：日报群")
            url_edit = QLineEdit(card)
            url_edit.setPlaceholderText("请输入 webhook URL")
            card_layout.addWidget(badge, 0, 0)
            card_layout.addWidget(alias_edit, 0, 1, 1, 2)
            card_layout.addWidget(url_edit, 0, 3, 1, 2)
            card_layout.setColumnStretch(4, 1)

            self.alias_rows.append({"alias": alias_edit, "url": url_edit})
            container_layout.addWidget(card)

        container_layout.addStretch(1)
        layout.addWidget(scroll, 1)
        layout.addWidget(self._build_action_strip(page, "保存机器人别名", self.save_aliases))
        return page

    def _build_mailbox_tab(self) -> QWidget:
        page = QWidget(self)
        layout = QVBoxLayout(page)
        layout.setSpacing(12)
        scroll, container, container_layout = self._build_scroll_container(page)

        header_card = QFrame(container)
        header_card.setObjectName("PanelCardSoft")
        header_layout = QGridLayout(header_card)
        header_layout.setContentsMargins(14, 8, 14, 8)
        header_layout.setHorizontalSpacing(8)
        header_layout.setVerticalSpacing(0)
        header_layout.setColumnMinimumWidth(0, 76)
        header_layout.setColumnMinimumWidth(1, 72)
        header_layout.setColumnMinimumWidth(2, 66)
        header_layout.setColumnMinimumWidth(3, 176)
        header_layout.setColumnMinimumWidth(4, 58)
        header_layout.setColumnMinimumWidth(5, 224)
        header_layout.setColumnMinimumWidth(6, 78)
        header_layout.setColumnMinimumWidth(7, 58)
        header_layout.setColumnStretch(0, 0)
        header_layout.setColumnStretch(1, 0)
        header_layout.setColumnStretch(2, 0)
        header_layout.setColumnStretch(3, 0)
        header_layout.setColumnStretch(4, 0)
        header_layout.setColumnStretch(5, 1)
        header_layout.setColumnStretch(6, 0)
        header_layout.setColumnStretch(7, 0)
        header_layout.setColumnMinimumWidth(8, 0)
        header_layout.addWidget(QLabel("", header_card), 0, 0)
        header_layout.addWidget(QLabel("", header_card), 0, 1)
        header_layout.addWidget(create_field_label("别名", header_card), 0, 2)
        header_layout.addWidget(create_field_label("服务器", header_card), 0, 3)
        header_layout.addWidget(create_field_label("端口", header_card), 0, 4)
        header_layout.addWidget(create_field_label("邮箱账号", header_card), 0, 5)
        header_layout.addWidget(create_field_label("密码", header_card), 0, 6)
        header_layout.addWidget(create_field_label("文件夹", header_card), 0, 7)
        container_layout.addWidget(header_card)

        for index in range(1, self.MAILBOX_SLOT_COUNT + 1):
            card = QFrame(container)
            card.setObjectName("PanelCard")
            card_layout = QGridLayout(card)
            card_layout.setContentsMargins(14, 10, 14, 10)
            card_layout.setHorizontalSpacing(8)
            card_layout.setVerticalSpacing(0)
            card_layout.setColumnMinimumWidth(0, 76)
            card_layout.setColumnMinimumWidth(1, 72)
            card_layout.setColumnMinimumWidth(2, 66)
            card_layout.setColumnMinimumWidth(3, 176)
            card_layout.setColumnMinimumWidth(4, 58)
            card_layout.setColumnMinimumWidth(5, 224)
            card_layout.setColumnMinimumWidth(6, 78)
            card_layout.setColumnMinimumWidth(7, 58)
            card_layout.setColumnStretch(0, 0)
            card_layout.setColumnStretch(1, 0)
            card_layout.setColumnStretch(2, 0)
            card_layout.setColumnStretch(3, 0)
            card_layout.setColumnStretch(4, 0)
            card_layout.setColumnStretch(5, 1)
            card_layout.setColumnStretch(6, 0)
            card_layout.setColumnStretch(7, 0)

            badge = create_status_pill(card, f"邮箱 {index}", "info")
            alias_edit = QLineEdit(card)
            host_edit = QLineEdit(card)
            port_edit = QLineEdit(card)
            username_edit = QLineEdit(card)
            password_edit = QLineEdit(card)
            password_edit.setEchoMode(QLineEdit.Password)
            mailbox_edit = QLineEdit(card)

            alias_edit.setPlaceholderText("邮箱别名")
            host_edit.setPlaceholderText("imap.example.com")
            port_edit.setPlaceholderText("993")
            username_edit.setPlaceholderText("邮箱账号")
            password_edit.setPlaceholderText("密码或授权码")
            mailbox_edit.setPlaceholderText("默认 INBOX")
            for widget in [alias_edit, host_edit, port_edit, username_edit, password_edit, mailbox_edit]:
                widget.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            badge.setMinimumWidth(88)
            badge.setMaximumWidth(88)
            alias_edit.setMaximumWidth(72)
            host_edit.setMaximumWidth(176)
            port_edit.setMaximumWidth(60)
            username_edit.setMaximumWidth(224)
            password_edit.setMaximumWidth(78)
            mailbox_edit.setMaximumWidth(64)

            test_btn = QPushButton("测试", card)
            set_button_variant(test_btn, "warn")
            test_btn.clicked.connect(lambda _checked=False, row_index=index - 1: self.test_mailbox_connection(row_index))
            test_btn.setMinimumWidth(64)
            test_btn.setMaximumWidth(64)

            card_layout.addWidget(badge, 0, 0)
            card_layout.addWidget(test_btn, 0, 1)
            card_layout.addWidget(alias_edit, 0, 2)
            card_layout.addWidget(host_edit, 0, 3)
            card_layout.addWidget(port_edit, 0, 4)
            card_layout.addWidget(username_edit, 0, 5)
            card_layout.addWidget(password_edit, 0, 6)
            card_layout.addWidget(mailbox_edit, 0, 7)

            self.mailbox_rows.append(
                {
                    "alias": alias_edit,
                    "host": host_edit,
                    "port": port_edit,
                    "username": username_edit,
                    "password": password_edit,
                    "mailbox": mailbox_edit,
                    "badge": badge,
                }
            )
            container_layout.addWidget(card)

        container_layout.addStretch(1)
        layout.addWidget(scroll, 1)
        layout.addWidget(self._build_action_strip(page, "保存邮箱配置", self.save_mailboxes))
        return page

    def _build_rule_tab(self) -> QWidget:
        page = QWidget(self)
        outer = QVBoxLayout(page)
        outer.setSpacing(12)
        scroll, container, container_layout = self._build_scroll_container(page)

        for index in range(1, self.RULE_SLOT_COUNT + 1):
            card = QFrame(container)
            card.setObjectName("PanelCard")
            card_layout = QVBoxLayout(card)
            card_layout.setContentsMargins(16, 14, 16, 14)
            card_layout.setSpacing(10)

            header_row = QHBoxLayout()
            header_row.setSpacing(8)
            collapse_btn = QPushButton("展开", card)
            collapse_btn.setCheckable(True)
            collapse_btn.setChecked(False)
            collapse_btn.setMinimumWidth(64)
            set_button_variant(collapse_btn, "warn")
            collapse_btn.clicked.connect(lambda checked=False, rule_index=index: self._toggle_rule_card(rule_index, checked))
            header_row.addWidget(collapse_btn)
            header_row.addWidget(create_status_pill(card, f"规则 {index}", "info"))
            enabled_box = QCheckBox("启用该规则", card)
            header_row.addWidget(enabled_box)
            header_row.addStretch(1)
            summary_state = create_status_pill(card, "未保存", "neutral")
            header_row.addWidget(summary_state)
            card_layout.addLayout(header_row)

            summary_label = QLabel("未保存", card)
            summary_label.setWordWrap(True)
            summary_label.setObjectName("SectionHint")
            card_layout.addWidget(summary_label)

            content = QWidget(card)
            content.setVisible(False)
            content_layout = QVBoxLayout(content)
            content_layout.setContentsMargins(0, 0, 0, 0)
            content_layout.setSpacing(10)

            keyword_edit = QLineEdit(card)
            type_edit = QLineEdit(card)
            name_keyword_edit = QLineEdit(card)
            mailbox_combo = QComboBox(card)
            webhook_combo = QComboBox(card)
            trigger_combo = QComboBox(card)
            trigger_combo.addItems(["周期检测", "定时检测"])
            interval_edit = QLineEdit(card)
            schedule_edit = QLineEdit(card)
            script_edit = QLineEdit(card)
            output_dir_edit = QLineEdit(card)
            max_size_edit = QLineEdit(card)

            keyword_edit.setPlaceholderText("例如：衡水装维营销日报")
            type_edit.setPlaceholderText("例如：xlsx")
            name_keyword_edit.setPlaceholderText("可留空")
            interval_edit.setPlaceholderText("默认 1")
            schedule_edit.setPlaceholderText("例如：08:30")
            script_edit.setPlaceholderText("可选：选择 .py 或 .exe")
            output_dir_edit.setPlaceholderText("脚本输出目录")
            max_size_edit.setPlaceholderText("默认跟随系统配置")

            match_card, match_layout = self._create_rule_section(card, "匹配条件")
            match_card.setObjectName("PanelCardSoft")
            match_form = self._create_form_grid()
            self._add_form_row(match_form, 0, "主题关键字", keyword_edit, "附件格式", type_edit)
            self._add_form_row(match_form, 1, "附件文件名关键字", name_keyword_edit, "所属邮箱", mailbox_combo)
            match_layout.addLayout(match_form)
            content_layout.addWidget(match_card)

            action_card, action_layout = self._create_rule_section(card, "执行动作")
            action_card.setObjectName("PanelCardSoft")
            action_form = self._create_form_grid()
            self._add_form_row(action_form, 0, "推送机器人", webhook_combo, "检测方式", trigger_combo)
            self._add_form_row(action_form, 1, "轮询间隔(min)", interval_edit, "定时时刻(HH:MM)", schedule_edit)

            script_browse_btn = QPushButton("选择程序", card)
            script_browse_btn.clicked.connect(
                lambda _checked=False, target=script_edit: self._choose_file(target, "选择处理程序", "处理程序 (*.py *.exe)")
            )
            output_browse_btn = QPushButton("输出目录", card)
            output_browse_btn.clicked.connect(lambda _checked=False, target=output_dir_edit: self._choose_directory(target))

            action_form.addWidget(create_field_label("处理程序", card), 2, 0)
            action_form.addWidget(self._build_line_action_row(card, script_edit, script_browse_btn), 2, 1)
            action_form.addWidget(create_field_label("输出目录", card), 2, 2)
            action_form.addWidget(self._build_line_action_row(card, output_dir_edit, output_browse_btn), 2, 3)
            action_form.addWidget(create_field_label("最大附件(MB)", card), 3, 0)
            action_form.addWidget(max_size_edit, 3, 1)
            action_layout.addLayout(action_form)
            content_layout.addWidget(action_card)

            content_layout.addWidget(self._build_action_strip(card, "保存当前规则", lambda checked=False, rule_index=index: self.save_single_rule(rule_index)))
            card_layout.addWidget(content)
            container_layout.addWidget(card)

            self.rule_cards[index] = {
                "collapse_btn": collapse_btn,
                "content": content,
                "enabled": enabled_box,
                "summary": summary_label,
                "summary_state": summary_state,
                "keyword": keyword_edit,
                "types": type_edit,
                "name_keyword": name_keyword_edit,
                "mailbox_alias": mailbox_combo,
                "webhook_alias": webhook_combo,
                "trigger_mode": trigger_combo,
                "interval": interval_edit,
                "schedule_time": schedule_edit,
                "script_path": script_edit,
                "output_dir": output_dir_edit,
                "max_size": max_size_edit,
            }

        container_layout.addStretch(1)
        outer.addWidget(scroll, 1)
        return page

    def _build_folder_tab(self) -> QWidget:
        page = QWidget(self)
        outer = QVBoxLayout(page)
        outer.setSpacing(12)
        scroll, container, container_layout = self._build_scroll_container(page)

        for index in range(1, self.FOLDER_SLOT_COUNT + 1):
            card = QFrame(container)
            card.setObjectName("PanelCard")
            card_layout = QVBoxLayout(card)
            card_layout.setContentsMargins(16, 14, 16, 14)
            card_layout.setSpacing(10)

            header_row = QHBoxLayout()
            header_row.setSpacing(8)
            header_row.addWidget(create_status_pill(card, f"检测 {index}", "info"))
            enabled_box = QCheckBox("启用该检测项", card)
            header_row.addWidget(enabled_box)
            header_row.addStretch(1)
            card_layout.addLayout(header_row)

            form_card, form_layout = self._create_rule_section(card, "")
            form_card.setObjectName("PanelCardSoft")
            form = self._create_form_grid()
            path_edit = QLineEdit(card)
            path_edit.setPlaceholderText("选择需要监测的本地目录")
            alias_combo = QComboBox(card)
            alias_combo.setMinimumWidth(220)
            alias_combo.setMaximumWidth(280)
            browse_btn = QPushButton("选择目录", card)
            browse_btn.clicked.connect(lambda _checked=False, target=path_edit: self._choose_directory(target))
            form.addWidget(create_field_label("监测路径", card), 0, 0)
            form.addWidget(self._build_line_action_row(card, path_edit, browse_btn), 0, 1)
            form.addWidget(create_field_label("推送机器人", card), 0, 2)
            form.addWidget(alias_combo, 0, 3)
            form_layout.addLayout(form)
            card_layout.addWidget(form_card)

            self.folder_cards[index] = {
                "enabled": enabled_box,
                "path": path_edit,
                "webhook_alias": alias_combo,
            }
            container_layout.addWidget(card)

        container_layout.addStretch(1)
        outer.addWidget(scroll, 1)
        outer.addWidget(self._build_action_strip(page, "保存文件夹检测配置", self.save_folder_settings))
        return page

    def _build_ui_tab(self) -> QWidget:
        page = QWidget(self)
        outer = QVBoxLayout(page)
        outer.setSpacing(12)

        card = QGroupBox("界面参数", page)
        card.setObjectName("FormSection")
        card_layout = self._create_form_grid(card)

        appearance_combo = QComboBox(card)
        appearance_combo.addItems(["light", "dark", "system"])
        theme_combo = QComboBox(card)
        theme_combo.addItems(list(self.THEME_LABEL_TO_VALUE.keys()))
        start_page_combo = QComboBox(card)
        start_page_combo.addItems(list(self.START_PAGE_LABEL_TO_VALUE.keys()))
        auto_scroll_box = QCheckBox("日志自动滚动", card)
        width_edit = QLineEdit(card)
        height_edit = QLineEdit(card)
        sidebar_edit = QLineEdit(card)
        poll_edit = QLineEdit(card)
        scale_edit = QLineEdit(card)
        timeout_edit = QLineEdit(card)

        self._add_form_row(card_layout, 0, "界面模式", appearance_combo, "颜色主题", theme_combo)
        self._add_form_row(card_layout, 1, "启动页", start_page_combo, "日志自动滚动", auto_scroll_box)
        self._add_form_row(card_layout, 2, "窗口宽度(px)", width_edit, "窗口高度(px)", height_edit)
        self._add_form_row(card_layout, 3, "侧栏宽度(px)", sidebar_edit, "日志刷新(ms)", poll_edit)
        self._add_form_row(card_layout, 4, "界面缩放", scale_edit, "脚本超时(s)", timeout_edit)

        outer.addWidget(card)
        outer.addWidget(self._build_action_strip(page, "保存界面设置", self.save_ui_settings))
        outer.addStretch(1)

        self.ui_widgets = {
            "appearance": appearance_combo,
            "color_theme": theme_combo,
            "start_page": start_page_combo,
            "auto_scroll_log": auto_scroll_box,
            "window_width": width_edit,
            "window_height": height_edit,
            "sidebar_width": sidebar_edit,
            "ui_log_poll_ms": poll_edit,
            "ui_scale": scale_edit,
            "script_timeout_seconds": timeout_edit,
        }
        return page

    def _build_path_tab(self) -> QWidget:
        page = QWidget(self)
        outer = QVBoxLayout(page)
        outer.setSpacing(12)

        card = QGroupBox("运行路径", page)
        card.setObjectName("FormSection")
        card_layout = QGridLayout(card)
        card_layout.setHorizontalSpacing(12)
        card_layout.setVerticalSpacing(10)
        card_layout.setColumnMinimumWidth(0, 96)
        card_layout.setColumnStretch(1, 1)

        download_edit = QLineEdit(card)
        download_btn = QPushButton("选择目录", card)
        download_btn.clicked.connect(lambda _checked=False, target=download_edit: self._choose_directory(target))
        state_edit = QLineEdit(card)
        state_btn = QPushButton("选择文件", card)
        state_btn.clicked.connect(lambda _checked=False, target=state_edit: self._choose_save_file(target, "选择状态文件"))

        card_layout.addWidget(create_field_label("下载目录", card), 0, 0)
        card_layout.addWidget(self._build_line_action_row(card, download_edit, download_btn), 0, 1)
        card_layout.addWidget(create_field_label("状态文件", card), 1, 0)
        card_layout.addWidget(self._build_line_action_row(card, state_edit, state_btn), 1, 1)

        outer.addWidget(card)
        outer.addWidget(self._build_action_strip(page, "保存路径设置", self.save_path_settings))
        outer.addStretch(1)

        self.path_widgets = {"downloads": download_edit, "state": state_edit}
        return page

    @staticmethod
    def _choose_file(target: QLineEdit, title: str, filter_text: str) -> None:
        file_path, _ = QFileDialog.getOpenFileName(None, title, "", filter_text)
        if file_path:
            target.setText(file_path)

    @staticmethod
    def _choose_directory(target: QLineEdit) -> None:
        directory = QFileDialog.getExistingDirectory(None, "选择输出目录")
        if directory:
            target.setText(directory)

    @staticmethod
    def _choose_save_file(target: QLineEdit, title: str) -> None:
        file_path, _ = QFileDialog.getSaveFileName(None, title, target.text().strip() or "", "JSON Files (*.json);;All Files (*.*)")
        if file_path:
            target.setText(file_path)

    def reload_all(self) -> None:
        self.config = load_config()
        self._reload_aliases()
        self._reload_mailboxes()
        self._reload_rules()
        self._reload_folders()
        self._reload_path_settings()
        self._reload_ui_settings()

    def _reload_aliases(self) -> None:
        aliases = load_webhook_aliases().get("aliases", {})
        items = list(aliases.items())
        for index, row in enumerate(self.alias_rows):
            name = items[index][0] if index < len(items) else ""
            url = items[index][1] if index < len(items) else ""
            self._set_line(row.get("alias"), name)
            self._set_line(row.get("url"), url)

    def _reload_mailboxes(self) -> None:
        mailboxes = load_mailbox_configs().get("mailboxes", [])
        for index, row in enumerate(self.mailbox_rows):
            mailbox = mailboxes[index] if index < len(mailboxes) else {}
            self._set_line(row.get("alias"), str(mailbox.get("alias", "")))
            self._set_line(row.get("host"), str(mailbox.get("host", "")))
            self._set_line(row.get("port"), str(mailbox.get("port", "")))
            self._set_line(row.get("username"), str(mailbox.get("username", "")))
            self._set_line(row.get("password"), str(mailbox.get("password", "")))
            self._set_line(row.get("mailbox"), str(mailbox.get("mailbox", "")))
            self._set_status_pill(row.get("badge"), f"邮箱 {index + 1}", "info")

    def _reload_rules(self) -> None:
        alias_names = [""] + sorted(load_webhook_aliases().get("aliases", {}).keys())
        mailbox_names = [""] + sorted(item.get("alias", "") for item in load_mailbox_configs().get("mailboxes", []) if item.get("alias", ""))
        rules = load_subject_attachment_rules().get("rules", [])

        for index in range(1, self.RULE_SLOT_COUNT + 1):
            card = self.rule_cards[index]
            rule = rules[index - 1] if index <= len(rules) else {}
            mailbox_combo = card["mailbox_alias"]
            webhook_combo = card["webhook_alias"]
            trigger_combo = card["trigger_mode"]
            assert isinstance(mailbox_combo, QComboBox)
            assert isinstance(webhook_combo, QComboBox)
            assert isinstance(trigger_combo, QComboBox)

            self._reset_combo(mailbox_combo, mailbox_names, str(rule.get("mailbox_alias", "")))
            self._reset_combo(webhook_combo, alias_names, str(rule.get("webhook_alias", "")))
            self._reset_combo(trigger_combo, ["周期检测", "定时检测"], self.TRIGGER_VALUE_TO_LABEL.get(str(rule.get("trigger_mode", "periodic")), "周期检测"))
            self._set_checkbox(card["enabled"], bool(rule.get("enabled", False)))
            self._set_line(card["keyword"], str(rule.get("keyword", "")))
            self._set_line(card["types"], (rule.get("types", [""]) or [""])[0])
            self._set_line(card["name_keyword"], (rule.get("filename_keywords", [""]) or [""])[0])
            self._set_line(card["interval"], str(max(1, int((rule.get("poll_interval_seconds") or 60) / 60))))
            self._set_line(card["schedule_time"], str(rule.get("schedule_time", "")))
            self._set_line(card["script_path"], str(rule.get("script_path", "")))
            self._set_line(card["output_dir"], str(rule.get("script_output_dir", "")))
            self._set_line(card["max_size"], str(rule.get("max_attachment_size_mb") or self.config.max_attachment_size_mb))
            self._set_summary(card["summary"], self._build_rule_summary(rule))
            self._set_status_pill(card.get("summary_state"), "已保存" if rule else "未保存", "success" if rule and rule.get("enabled") else "neutral")

    def _reload_folders(self) -> None:
        alias_names = [""] + sorted(load_webhook_aliases().get("aliases", {}).keys())
        config_file = Path("settings/folder_monitor_config.json")
        if config_file.exists():
            try:
                folder_config = json.loads(config_file.read_text(encoding="utf-8"))
            except Exception:
                folder_config = {}
        else:
            folder_config = {}

        for index in range(1, self.FOLDER_SLOT_COUNT + 1):
            payload = folder_config.get(f"folder_{index}", {})
            card = self.folder_cards[index]
            alias_combo = card["webhook_alias"]
            assert isinstance(alias_combo, QComboBox)
            self._set_checkbox(card["enabled"], bool(payload.get("enabled", False)))
            self._set_line(card["path"], str(payload.get("path", "")))
            self._reset_combo(alias_combo, alias_names, str(payload.get("webhook_alias", "")))

    def _reload_ui_settings(self) -> None:
        appearance = self.ui_widgets.get("appearance")
        color_theme = self.ui_widgets.get("color_theme")
        start_page = self.ui_widgets.get("start_page")
        auto_scroll = self.ui_widgets.get("auto_scroll_log")

        if isinstance(appearance, QComboBox):
            self._reset_combo(appearance, ["light", "dark", "system"], self.config.ui_appearance)
        if isinstance(color_theme, QComboBox):
            theme_label = next((label for label, value in self.THEME_LABEL_TO_VALUE.items() if value == self.config.ui_color_theme), "标准蓝色")
            self._reset_combo(color_theme, list(self.THEME_LABEL_TO_VALUE.keys()), theme_label)
        if isinstance(start_page, QComboBox):
            start_label = next((label for label, value in self.START_PAGE_LABEL_TO_VALUE.items() if value == self.config.start_page), "邮件检测")
            self._reset_combo(start_page, list(self.START_PAGE_LABEL_TO_VALUE.keys()), start_label)
        if isinstance(auto_scroll, QCheckBox):
            auto_scroll.setChecked(bool(self.config.auto_scroll_log))

        self._set_line(self.ui_widgets.get("window_width"), str(self.config.window_width))
        self._set_line(self.ui_widgets.get("window_height"), str(self.config.window_height))
        self._set_line(self.ui_widgets.get("sidebar_width"), str(self.config.sidebar_width))
        self._set_line(self.ui_widgets.get("ui_log_poll_ms"), str(self.config.ui_log_poll_ms))
        self._set_line(self.ui_widgets.get("ui_scale"), str(self.config.ui_scale))
        self._set_line(self.ui_widgets.get("script_timeout_seconds"), str(self.config.script_timeout_seconds))

    def _reload_path_settings(self) -> None:
        self._set_line(self.path_widgets.get("downloads"), str(self.config.download_dir))
        self._set_line(self.path_widgets.get("state"), str(self.config.state_file))

    @staticmethod
    def _set_checkbox(widget: object, value: bool) -> None:
        if isinstance(widget, QCheckBox):
            widget.setChecked(value)

    @staticmethod
    def _set_line(widget: object, value: str) -> None:
        if isinstance(widget, QLineEdit):
            widget.setText(value)

    @staticmethod
    def _reset_combo(widget: QComboBox, values: list[str], current: str) -> None:
        widget.blockSignals(True)
        widget.clear()
        widget.addItems(values)
        if current in values:
            widget.setCurrentText(current)
        elif values:
            widget.setCurrentIndex(0)
        widget.blockSignals(False)

    @staticmethod
    def _set_summary(widget: object, value: str) -> None:
        if isinstance(widget, QLabel):
            widget.setText(value)

    @staticmethod
    def _set_status_pill(widget: object, text: str, tone: str) -> None:
        if isinstance(widget, QLabel):
            widget.setText(text)
            widget.setProperty("pillTone", tone)
            widget.style().unpolish(widget)
            widget.style().polish(widget)
            widget.update()

    def _build_rule_summary(self, rule: dict) -> str:
        if not rule:
            return "未保存"
        status = "已启用" if rule.get("enabled") else "未启用"
        keyword = str(rule.get("keyword", "")).strip() or "未填写主题"
        mailbox_alias = str(rule.get("mailbox_alias", "")).strip() or "未选邮箱"
        webhook_alias = str(rule.get("webhook_alias", "")).strip() or "未选机器人"
        trigger_mode = str(rule.get("trigger_mode", "periodic")).strip() or "periodic"
        trigger_text = f"定时 {str(rule.get('schedule_time', '')).strip() or '--:--'}" if trigger_mode == "timed" else f"周期 {max(1, int(rule.get('poll_interval_seconds', 60) or 60) // 60)} 分钟"
        mode_text = "脚本处理" if str(rule.get("script_path", "")).strip() else "直接推送"
        return f"{status} | {keyword} | {mailbox_alias} -> {webhook_alias} | {trigger_text} | {mode_text}"

    def _collect_mailbox_row_values(self, row: int) -> list[str]:
        item = self.mailbox_rows[row]
        return [
            self._line_text(item.get("alias")),
            self._line_text(item.get("host")),
            self._line_text(item.get("port")),
            self._line_text(item.get("username")),
            self._line_text(item.get("password")),
            self._line_text(item.get("mailbox")),
        ]

    def test_mailbox_connection(self, row: int) -> None:
        alias, host, port_text, username, password, mailbox = self._collect_mailbox_row_values(row)
        if not any([alias, host, port_text, username, password, mailbox]):
            self._show_error("当前卡片还是空的，请先填写邮箱连接参数")
            return
        if not alias or not host or not port_text or not username or not password:
            self._show_error("请先填写完整的邮箱别名、服务器、端口、邮箱账号和密码/授权码")
            return
        try:
            port = int(port_text)
            if port <= 0:
                raise ValueError
        except ValueError:
            self._show_error("端口必须是大于 0 的整数")
            return

        badge = self.mailbox_rows[row].get("badge")
        self._set_status_pill(badge, f"邮箱 {row + 1}", "warning")

        def worker() -> None:
            try:
                with ImapMailClient(host=host, port=port, username=username, password=password, mailbox=mailbox or "INBOX", timeout_seconds=15):
                    pass
                self.mailbox_test_finished.emit(row, True, f"邮箱“{alias}”连接成功")
            except Exception as exc:
                self.mailbox_test_finished.emit(row, False, f"邮箱“{alias}”连接失败：{exc}")

        threading.Thread(target=worker, daemon=True).start()

    def _on_mailbox_test_finished(self, row: int, success: bool, message: str) -> None:
        badge = self.mailbox_rows[row].get("badge")
        self._set_status_pill(badge, f"邮箱 {row + 1}", "success" if success else "danger")
        if success:
            self._show_info(message)
        else:
            self._show_error(message)

    def save_aliases(self) -> None:
        aliases: dict[str, str] = {}
        for row in self.alias_rows:
            name = self._line_text(row.get("alias"))
            url = self._line_text(row.get("url"))
            if not name and not url:
                continue
            if not name or not url:
                self._show_error("别名和 URL 必须同时填写")
                return
            aliases[name] = url

        if not aliases:
            self._show_error("请至少填写一个机器人别名")
            return

        save_webhook_aliases(aliases, next(iter(aliases.keys()), ""))
        self.reload_all()
        self.config_changed.emit()
        self._show_info("机器人别名已保存")

    def save_mailboxes(self) -> None:
        mailboxes: list[dict[str, object]] = []
        for row_index in range(self.MAILBOX_SLOT_COUNT):
            alias, host, port_text, username, password, mailbox = self._collect_mailbox_row_values(row_index)
            if not any([alias, host, port_text, username, password, mailbox]):
                continue
            if not alias or not host or not port_text or not username or not password:
                self._show_error(f"邮箱配置第 {row_index + 1} 项未填写完整")
                return
            try:
                port = int(port_text)
                if port <= 0:
                    raise ValueError
            except ValueError:
                self._show_error(f"邮箱配置第 {row_index + 1} 项端口无效")
                return
            mailboxes.append(
                {
                    "alias": alias,
                    "host": host,
                    "port": port,
                    "username": username,
                    "password": password,
                    "mailbox": mailbox or "INBOX",
                }
            )

        if not mailboxes:
            self._show_error("请至少填写一个邮箱配置")
            return

        save_mailbox_configs(mailboxes)
        self.reload_all()
        self.config_changed.emit()
        self._show_info("邮箱配置已保存")

    def save_single_rule(self, index: int) -> None:
        card = self.rule_cards[index]
        enabled = self._checked(card["enabled"])
        keyword = self._line_text(card["keyword"])
        types_text = self._line_text(card["types"])
        name_keyword = self._line_text(card["name_keyword"])
        mailbox_alias = self._combo_text(card["mailbox_alias"])
        webhook_alias = self._combo_text(card["webhook_alias"])
        trigger_label = self._combo_text(card["trigger_mode"]) or "周期检测"
        trigger_mode = "timed" if trigger_label == "定时检测" else "periodic"
        interval_text = self._line_text(card["interval"])
        schedule_time = self._line_text(card["schedule_time"])
        script_path = self._line_text(card["script_path"])
        output_dir = self._line_text(card["output_dir"])
        max_size_text = self._line_text(card["max_size"])

        is_empty = not any([enabled, keyword, types_text, name_keyword, mailbox_alias, webhook_alias, schedule_time, script_path, output_dir, max_size_text])
        current_rules = list(load_subject_attachment_rules().get("rules", []))
        while len(current_rules) < index:
            current_rules.append({})

        if is_empty:
            current_rules[index - 1] = {}
            save_subject_attachment_rules(current_rules)
            self.reload_all()
            self.config_changed.emit()
            self._show_info(f"规则 {index} 已清空并保存")
            return

        if not keyword or not types_text:
            self._show_error("主题关键字和附件格式必须同时填写")
            return

        parsed_types = parse_types_input(types_text)
        parsed_keywords = parse_filename_keywords_input(name_keyword)
        if not parsed_types:
            self._show_error("附件格式无效，例如：xlsx 或 png")
            return
        try:
            interval_minutes = max(1, int(interval_text or "1"))
            max_size = max(1, int(max_size_text or str(self.config.max_attachment_size_mb)))
        except ValueError:
            self._show_error("轮询间隔和最大附件必须是整数")
            return

        if trigger_mode == "timed" and schedule_time and (len(schedule_time) != 5 or schedule_time[2] != ":"):
            self._show_error("定时时刻格式应为 HH:MM")
            return

        aliases = load_webhook_aliases().get("aliases", {})
        webhook_url = aliases.get(webhook_alias, "").strip()
        if enabled and not webhook_url:
            self._show_error("当前规则已启用，但未选择有效机器人")
            return

        if script_path:
            suffix = Path(script_path).suffix.lower()
            if suffix not in {".py", ".exe"}:
                self._show_error("处理程序仅支持 .py 或 .exe")
                return
            if output_dir.strip() == "":
                self._show_error("选择处理程序后，必须填写输出目录")
                return

        current_rules[index - 1] = {
            "enabled": enabled,
            "keyword": keyword,
            "types": parsed_types,
            "filename_keywords": parsed_keywords,
            "mailbox_alias": mailbox_alias,
            "webhook_alias": webhook_alias,
            "webhook_url": webhook_url,
            "script_path": script_path,
            "script_output_dir": output_dir,
            "trigger_mode": trigger_mode,
            "schedule_time": schedule_time if trigger_mode == "timed" else "",
            "poll_interval_seconds": interval_minutes * 60,
            "max_attachment_size_mb": max_size,
        }

        save_subject_attachment_rules(current_rules)
        self.reload_all()
        self.config_changed.emit()
        self._show_info(f"规则 {index} 已保存")

    def save_folder_settings(self) -> None:
        aliases = load_webhook_aliases().get("aliases", {})
        folder_monitor_config: dict[str, dict[str, object]] = {}
        warnings: list[str] = []
        errors: list[str] = []

        for index in range(1, self.FOLDER_SLOT_COUNT + 1):
            card = self.folder_cards[index]
            enabled = self._checked(card["enabled"])
            path_text = self._line_text(card["path"])
            webhook_alias = self._combo_text(card["webhook_alias"])
            webhook_url = aliases.get(webhook_alias, "").strip()

            if not any([enabled, path_text, webhook_alias]):
                continue
            if enabled:
                if not path_text:
                    warnings.append(f"检测 {index}: 已启用但未设置文件夹路径")
                elif not Path(path_text).exists():
                    warnings.append(f"检测 {index}: 文件夹不存在 - {path_text}")
                if not webhook_alias:
                    errors.append(f"检测 {index}: 已启用但未选择推送机器人")
                elif not webhook_url:
                    errors.append(f"检测 {index}: 推送机器人无效 - {webhook_alias}")

            folder_monitor_config[f"folder_{index}"] = {
                "path": path_text,
                "webhook_alias": webhook_alias,
                "webhook_url": webhook_url,
                "enabled": enabled,
            }

        if errors:
            self._show_error("\n".join(errors))
            return

        config_file = Path("settings/folder_monitor_config.json")
        config_file.parent.mkdir(parents=True, exist_ok=True)
        tmp_file = config_file.with_suffix(config_file.suffix + ".tmp")
        tmp_file.write_text(json.dumps(folder_monitor_config, indent=2, ensure_ascii=False), encoding="utf-8")
        tmp_file.replace(config_file)

        message = f"文件夹检测配置已保存，当前启用 {sum(1 for item in folder_monitor_config.values() if item.get('enabled'))} 项"
        if warnings:
            message += "\n\n警告：\n" + "\n".join(warnings)
        self.reload_all()
        self.config_changed.emit()
        self._show_info(message)

    def save_ui_settings(self) -> None:
        appearance = self._combo_text(self.ui_widgets.get("appearance"))
        theme_label = self._combo_text(self.ui_widgets.get("color_theme"))
        start_page_label = self._combo_text(self.ui_widgets.get("start_page"))
        auto_scroll = self._checked(self.ui_widgets.get("auto_scroll_log"))

        try:
            window_width = int(self._line_text(self.ui_widgets.get("window_width")))
            window_height = int(self._line_text(self.ui_widgets.get("window_height")))
            sidebar_width = int(self._line_text(self.ui_widgets.get("sidebar_width")))
            ui_log_poll_ms = int(self._line_text(self.ui_widgets.get("ui_log_poll_ms")))
            script_timeout_seconds = int(self._line_text(self.ui_widgets.get("script_timeout_seconds")))
            ui_scale = float(self._line_text(self.ui_widgets.get("ui_scale")))
        except ValueError:
            self._show_error("窗口宽高、侧栏宽度、日志刷新、脚本超时必须为整数，界面缩放必须为数字")
            return

        if min(window_width, window_height, sidebar_width, ui_log_poll_ms, script_timeout_seconds) <= 0 or ui_scale <= 0:
            self._show_error("界面参数必须大于 0")
            return
        if appearance not in {"light", "dark", "system"}:
            self._show_error("界面模式仅支持 light / dark / system")
            return

        start_page_value = self.START_PAGE_LABEL_TO_VALUE.get(start_page_label, "execute")
        theme_value = self.THEME_LABEL_TO_VALUE.get(theme_label, "blue")
        upsert_env_file(
            Path("settings/app_config.json"),
            {
                "UI_APPEARANCE": appearance,
                "UI_COLOR_THEME": theme_value,
                "START_PAGE": start_page_value,
                "WINDOW_WIDTH": str(window_width),
                "WINDOW_HEIGHT": str(window_height),
                "SIDEBAR_WIDTH": str(sidebar_width),
                "UI_LOG_POLL_MS": str(ui_log_poll_ms),
                "AUTO_SCROLL_LOG": "true" if auto_scroll else "false",
                "UI_SCALE": str(ui_scale),
                "SCRIPT_TIMEOUT_SECONDS": str(script_timeout_seconds),
            },
        )
        self.reload_all()
        self.config_changed.emit()
        self._show_info("界面设置已保存，重启后全部生效")

    def save_path_settings(self) -> None:
        download_dir = self._line_text(self.path_widgets.get("downloads"))
        state_file = self._line_text(self.path_widgets.get("state"))
        if not download_dir:
            self._show_error("下载目录不能为空")
            return
        if not state_file:
            self._show_error("状态文件不能为空")
            return
        upsert_env_file(Path("settings/app_config.json"), {"DOWNLOAD_DIR": download_dir, "STATE_FILE": state_file})
        self.reload_all()
        self.config_changed.emit()
        self._show_info("路径设置已保存")

    @staticmethod
    def _checked(widget: object) -> bool:
        return bool(widget.isChecked()) if isinstance(widget, QCheckBox) else False

    @staticmethod
    def _line_text(widget: object) -> str:
        return widget.text().strip() if isinstance(widget, QLineEdit) else ""

    @staticmethod
    def _combo_text(widget: object) -> str:
        return widget.currentText().strip() if isinstance(widget, QComboBox) else ""

    def _show_info(self, message: str) -> None:
        QMessageBox.information(self, "保存成功", message)

    def _show_error(self, message: str) -> None:
        QMessageBox.critical(self, "验证失败", message)

    def on_page_activated(self) -> None:
        self.reload_all()
