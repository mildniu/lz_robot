from pathlib import Path

import customtkinter as ctk

class AboutPage(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.summary_value = None
        self.status_value = None
        self.setup_ui()
        self.refresh_info()

    def setup_ui(self):
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=25, pady=20)

        title = ctk.CTkLabel(
            container,
            text="工具简介",
            font=ctk.CTkFont(size=28, weight="bold"),
        )
        title.pack(anchor="w", pady=(0, 10))

        subtitle = ctk.CTkLabel(
            container,
            text="使用说明",
            font=ctk.CTkFont(size=16),
            text_color=("gray30", "gray70"),
        )
        subtitle.pack(anchor="w", pady=(0, 16))

        info_card = ctk.CTkFrame(container, fg_color=("gray95", "gray20"))
        info_card.pack(fill="x", pady=(0, 12))

        self.summary_value = ctk.CTkLabel(
            info_card,
            text="",
            justify="left",
            anchor="w",
            font=ctk.CTkFont(size=13),
        )
        self.summary_value.pack(fill="x", padx=16, pady=16)

        status_card = ctk.CTkFrame(container, fg_color=("gray95", "gray20"))
        status_card.pack(fill="x", pady=(0, 12))

        self.status_value = ctk.CTkLabel(
            status_card,
            text="",
            justify="left",
            anchor="w",
            font=ctk.CTkFont(size=13),
        )
        self.status_value.pack(fill="x", padx=16, pady=16)

        refresh_btn = ctk.CTkButton(
            container,
            text="刷新信息",
            command=self.refresh_info,
            width=120,
            height=36,
        )
        refresh_btn.pack(anchor="w")

    def refresh_info(self):
        try:
            summary = "\n".join(
                [
                    "【工具简介】",
                    "量子推送机器人 v5.1",
                    "支持邮件检测与文件夹检测两种模式，可长期运行。",
                    "邮件规则支持按所属邮箱、附件格式、机器人、处理程序(.py/.exe)进行细分处理。",
                    "可直接推送原附件，也可先由脚本处理后推送文字、图片、文件。",
                    "",
                    "【使用说明】",
                    "1. 先在“机器人别名”中维护 webhook 地址。",
                    "2. 再到“邮箱配置”维护 IMAP 邮箱别名，并测试连接。",
                    "3. 在“邮箱检测规则”中选择所属邮箱、推送机器人和处理方式。",
                    "4. 如需文件夹推送，再到“文件夹检测”配置监测目录和目标机器人。",
                    "5. 最后在“邮件检测”或“文件夹检测”页点击测试或启动，查看日志结果。",
                ]
            )
        except Exception as exc:
            summary = "\n".join(
                [
                    "【工具简介】",
                    "量子推送机器人 v5.1",
                    "支持邮件检测与文件夹检测两种模式，可长期运行。",
                    "支持脚本处理、机器人推送和文件留存。",
                    "",
                    f"配置读取失败: {exc}",
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
                "【配置文件状态】",
                f"- 主配置: {'存在' if app_config_file.exists() else '不存在'} ({app_config_file})",
                f"- 邮箱别名配置: {'存在' if mailbox_file.exists() else '不存在'} ({mailbox_file})",
                f"- 机器人别名配置: {'存在' if alias_file.exists() else '不存在'} ({alias_file})",
                f"- 邮箱检测规则: {'存在' if subject_rule_file.exists() else '不存在'} ({subject_rule_file})",
                f"- 文件夹检测配置: {'存在' if folder_monitor_file.exists() else '不存在'} ({folder_monitor_file})",
                "",
                "【运行状态文件】",
                f"- 邮件状态: {'存在' if state_file.exists() else '不存在'} ({state_file})",
                f"- 文件发送状态: {'存在' if file_state_file.exists() else '不存在'} ({file_state_file})",
            ]
        )

        self.summary_value.configure(text=summary)
        self.status_value.configure(text=status)

    def on_page_activated(self):
        self.refresh_info()

    def on_external_config_updated(self):
        self.refresh_info()
