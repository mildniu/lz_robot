from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QPushButton, QVBoxLayout
from qt_components import create_status_pill, set_button_variant

from .base import BasePage


class AboutPage(BasePage):
    def __init__(self, log_bus) -> None:
        super().__init__(log_bus, "关于", "PySide6 版关于页，显示工具简介、使用说明和配置状态。")
        summary_card = QFrame(self)
        summary_card.setObjectName("SectionCard")
        summary_layout = QVBoxLayout(summary_card)
        summary_layout.setContentsMargins(20, 18, 20, 18)
        summary_layout.setSpacing(10)

        top_row = QHBoxLayout()
        top_row.addWidget(create_status_pill(summary_card, "v5.1", "info"))
        top_row.addStretch(1)
        summary_layout.addLayout(top_row)

        summary_title = QLabel("工具简介与使用说明", self)
        summary_title.setObjectName("SectionTitle")
        summary_layout.addWidget(summary_title)
        self.summary_label = QLabel(self)
        self.summary_label.setWordWrap(True)
        self.summary_label.setObjectName("SectionHint")
        summary_layout.addWidget(self.summary_label)
        self.layout.addWidget(summary_card)

        status_card = QFrame(self)
        status_card.setObjectName("PanelCard")
        status_layout = QVBoxLayout(status_card)
        status_layout.setContentsMargins(20, 18, 20, 18)
        status_layout.setSpacing(10)
        status_title = QLabel("配置文件状态", self)
        status_title.setObjectName("SectionTitle")
        status_layout.addWidget(status_title)
        self.status_label = QLabel(self)
        self.status_label.setWordWrap(True)
        self.status_label.setObjectName("SectionHint")
        status_layout.addWidget(self.status_label)
        self.layout.addWidget(status_card)

        refresh_btn = QPushButton("刷新信息", self)
        set_button_variant(refresh_btn, "primary")
        refresh_btn.clicked.connect(self.refresh_info)
        self.layout.addWidget(refresh_btn)
        self.layout.addStretch(1)

        self.refresh_info()

    def refresh_info(self) -> None:
        summary = "\n".join(
            [
                "量子推送机器人 v5.1",
                "",
                "支持邮件检测与文件夹检测两种模式，可长期运行。",
                "支持按邮箱别名、机器人别名、附件格式、脚本处理程序进行规则化处理。",
                "",
                "推荐使用顺序",
                "1. 先配置机器人别名。",
                "2. 再配置邮箱别名，并测试连接。",
                "3. 在邮箱检测规则中逐条保存规则。",
                "4. 如需文件夹推送，再配置文件夹检测。",
                "5. 最后在邮件检测或文件夹检测页启动并观察日志。",
            ]
        )

        state_file = Path("state/mail_state.json")
        file_state_file = Path("state/file_sent_state.json")
        folder_monitor_file = Path("settings/folder_monitor_config.json")
        alias_file = Path("settings/webhook_aliases.json")
        mailbox_file = Path("settings/mailbox_aliases.json")
        app_config_file = Path("settings/app_config.json")
        subject_rule_file = Path("settings/subject_attachment_rules.json")
        status = "\n".join(
            [
                f"主配置: {'存在' if app_config_file.exists() else '不存在'} ({app_config_file})",
                f"邮箱别名配置: {'存在' if mailbox_file.exists() else '不存在'} ({mailbox_file})",
                f"机器人别名配置: {'存在' if alias_file.exists() else '不存在'} ({alias_file})",
                f"邮箱检测规则: {'存在' if subject_rule_file.exists() else '不存在'} ({subject_rule_file})",
                f"文件夹检测配置: {'存在' if folder_monitor_file.exists() else '不存在'} ({folder_monitor_file})",
                "",
                f"邮件状态: {'存在' if state_file.exists() else '不存在'} ({state_file})",
                f"文件发送状态: {'存在' if file_state_file.exists() else '不存在'} ({file_state_file})",
            ]
        )

        self.summary_label.setText(summary)
        self.status_label.setText(status)

    def on_page_activated(self) -> None:
        self.refresh_info()

    def on_external_config_updated(self) -> None:
        self.refresh_info()
