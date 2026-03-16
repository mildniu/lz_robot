import threading
import time
from typing import Dict

import customtkinter as ctk

from mail_forwarder import MailProcessingService, load_config
from mail_forwarder.mailbox_store import load_mailbox_configs

from .common import LogHandler, ModernButton


class ExecutePage(ctk.CTkFrame):
    def __init__(self, master, log_handler: LogHandler, auto_scroll_log: bool = True, **kwargs):
        super().__init__(master, **kwargs)
        self.log_handler = log_handler
        self.auto_scroll_log = auto_scroll_log
        self.worker_thread = None
        self.is_running = False
        self.cycle_count = 0
        self.setup_ui()
        self.log_handler.add_callback(self.append_log)

    def setup_ui(self):
        control_bar = ctk.CTkFrame(self, height=100)
        control_bar.pack(fill="x", padx=25, pady=(20, 10))

        left_panel = ctk.CTkFrame(control_bar, fg_color="transparent")
        left_panel.pack(side="left", fill="both", expand=True)

        title = ctk.CTkLabel(left_panel, text="📧 邮件检测", font=ctk.CTkFont(size=26, weight="bold"))
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
            text="停止",
            icon="⏹️",
            command=self.stop_worker,
            width=120,
            height=45,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color="#EF5350",
            hover_color="#C62828",
            state="disabled",
        )
        self.stop_btn.pack(side="left", padx=5)

        self.test_btn = ModernButton(
            right_panel,
            text="测试",
            icon="🧪",
            command=self.test_once,
            width=120,
            height=45,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color="#42A5F5",
            hover_color="#1976D2",
        )
        self.test_btn.pack(side="left", padx=5)

        log_container = ctk.CTkFrame(self)
        log_container.pack(fill="both", expand=True, padx=25, pady=(0, 20))

        log_header = ctk.CTkFrame(log_container, height=50)
        log_header.pack(fill="x")

        log_title = ctk.CTkLabel(log_header, text="📋 邮件检测日志", font=ctk.CTkFont(size=18, weight="bold"))
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

    @staticmethod
    def format_wait_text(seconds: int) -> str:
        if seconds % 60 == 0 and seconds >= 60:
            minutes = seconds // 60
            return f"{minutes} 分钟"
        return f"{seconds} 秒"

    def log_batch_result(self, batch_result):
        results = list(getattr(batch_result, "results", []) or [])
        processed_count = sum(1 for item in results if item.status == "processed")
        skipped_count = sum(1 for item in results if item.status == "skipped")
        not_found_count = sum(1 for item in results if item.status == "not_found")
        error_count = sum(1 for item in results if item.status == "error")
        self.log_handler.info(
            f"本轮规则统计: 成功 {processed_count}，跳过 {skipped_count}，未找到 {not_found_count}，失败 {error_count}"
        )
        for item in results:
            mailbox_text = f"[{item.mailbox_alias}] " if getattr(item, "mailbox_alias", "") else ""
            folder_text = f"(文件夹: {item.mailbox_folder}) " if getattr(item, "mailbox_folder", "") else ""
            keyword_text = f"规则“{item.rule_keyword}”" if getattr(item, "rule_keyword", "") else "规则"
            if item.status == "processed":
                self.log_handler.success(
                    f"{mailbox_text}{folder_text}{keyword_text}完成: uid={item.uid}, 附件数量={len(item.files)}"
                )
            elif item.status == "not_found":
                self.log_handler.info(f"{mailbox_text}{folder_text}{keyword_text}: {item.reason}")
            elif item.status == "skipped":
                self.log_handler.warning(f"{mailbox_text}{folder_text}{keyword_text}: {item.reason}")
            else:
                self.log_handler.error(f"{mailbox_text}{folder_text}{keyword_text}: {item.reason}")

    def append_log(self, log_entry: Dict[str, str]):
        icons = {"INFO": "ℹ️", "SUCCESS": "✅", "WARNING": "⚠️", "ERROR": "❌"}
        log_line = f"[{log_entry['time']}] {icons.get(log_entry['level'], '•')} {log_entry['message']}\n"
        self.log_text.insert("end", log_line)
        if self.auto_scroll_log:
            self.log_text.see("end")

    def clear_logs(self):
        self.log_text.delete("1.0", "end")

    def start_worker(self):
        if self.is_running:
            return

        if self.worker_thread and self.worker_thread.is_alive():
            self.log_handler.warning("上一次任务仍在退出中，请稍后再试")
            return

        try:
            config = load_config()
            mailbox_payload = load_mailbox_configs()
            self.is_running = True
            self.cycle_count = 0
            self.start_btn.configure(state="disabled")
            self.stop_btn.configure(state="normal")
            self.status_label.configure(text="🟢 运行中", text_color="#2CC985")

            self.log_handler.info("=" * 50)
            self.log_handler.info("量子机器人已启动")
            self.log_handler.info("轮询间隔: 按命中规则配置（规则页填写单位为分钟）")
            self.log_handler.info(f"已配置邮箱别名数量: {len(mailbox_payload.get('mailboxes', []))}")
            self.log_handler.info("邮箱检测规则: 以“邮箱检测规则”标签页已启用规则为准")
            self.log_handler.info(f"状态文件: {config.state_file}")
            self.log_handler.info("=" * 50)

            def run_worker():
                service = MailProcessingService(config)
                while self.is_running:
                    self.cycle_count += 1
                    cycle_started_at = time.time()
                    try:
                        self.log_handler.info("-" * 40)
                        self.log_handler.info(f"第 {self.cycle_count} 轮开始检查邮件")
                        batch_result = service.process_rule_batch(
                            update_state=True,
                            event_callback=lambda level, message: self.log_handler.log(level, message),
                        )
                        self.log_batch_result(batch_result)
                        next_wait_seconds = (
                            batch_result.next_poll_interval_seconds
                            if getattr(batch_result, "next_poll_interval_seconds", 0) > 0
                            else config.poll_interval_seconds
                        )
                    except Exception as exc:
                        self.log_handler.error(f"执行出错: {exc}")
                        next_wait_seconds = config.poll_interval_seconds
                    finally:
                        elapsed = time.time() - cycle_started_at
                        self.log_handler.info(f"第 {self.cycle_count} 轮结束, 耗时 {elapsed:.1f} 秒")

                    if self.is_running:
                        wait_seconds = next_wait_seconds
                        self.log_handler.info(
                            f"进入等待: {self.format_wait_text(wait_seconds)} 后开始下一轮"
                        )
                        for remaining in range(wait_seconds, 0, -1):
                            if not self.is_running:
                                break
                            if remaining in (60, 30, 10, 5, 3, 2, 1):
                                self.log_handler.info(
                                    f"距离下一轮还有 {self.format_wait_text(remaining)}"
                                )
                            time.sleep(1)

                self.log_handler.info("工作线程已停止")

            self.worker_thread = threading.Thread(target=run_worker, daemon=True)
            self.worker_thread.start()
        except Exception as exc:
            self.is_running = False
            self.start_btn.configure(state="normal")
            self.stop_btn.configure(state="disabled")
            self.status_label.configure(text="❌ 错误", text_color="#EF5350")
            self.log_handler.error(f"启动失败: {exc}")

    def stop_worker(self):
        if not self.is_running and not (self.worker_thread and self.worker_thread.is_alive()):
            return

        self.log_handler.info("正在停止...")
        self.is_running = False
        if self.worker_thread and self.worker_thread.is_alive():
            self.worker_thread.join(timeout=5)
        if self.worker_thread and self.worker_thread.is_alive():
            self.start_btn.configure(state="normal")
            self.stop_btn.configure(state="disabled")
            self.status_label.configure(text="🟡 停止中", text_color="#FFB300")
            self.log_handler.warning("线程仍在退出中，可能正等待网络请求超时")
            return

        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        self.status_label.configure(text="⏸️ 已停止", text_color="gray")
        self.log_handler.success("量子机器人已停止")

    def test_once(self):
        if self.is_running:
            self.log_handler.warning("机器人正在运行中，请先停止")
            return

        def run_test():
            try:
                self.log_handler.info("=" * 50)
                self.log_handler.info("开始测试 - 处理最新邮件")
                self.log_handler.info("=" * 50)

                self.log_handler.info("正在加载配置...")
                config = load_config()
                mailbox_payload = load_mailbox_configs()
                self.log_handler.success("配置加载成功")
                self.log_handler.info("初始化邮件处理服务...")
                self.log_handler.info(f"邮箱别名数量: {len(mailbox_payload.get('mailboxes', []))}")
                if not mailbox_payload.get("mailboxes", []):
                    self.log_handler.info("当前使用兼容模式：将 app_config.json 中的单邮箱视为“默认邮箱”")
                self.log_handler.info("邮箱检测规则: 以“邮箱检测规则”标签页已启用规则为准")

                service = MailProcessingService(config)
                batch_result = service.process_rule_batch(
                    force=True,
                    update_state=False,
                    event_callback=lambda level, message: self.log_handler.log(level, message),
                )
                self.log_batch_result(batch_result)
                processed_results = [item for item in batch_result.results if item.status == "processed"]
                if not processed_results:
                    self.log_handler.warning("测试完成，但本次没有规则成功处理邮件")
                    self.log_handler.info("排查建议：确认所属邮箱、文件夹、主题关键字、附件格式是否配置正确")
                    return
                for item in processed_results:
                    mailbox_text = f"[{item.mailbox_alias}] " if item.mailbox_alias else ""
                    self.log_handler.success(f"{mailbox_text}找到邮件 UID: {item.uid}")
                    self.log_handler.success(f"提取到 {len(item.files)} 个附件:")
                    for index, file_path in enumerate(item.files, 1):
                        size_mb = file_path.stat().st_size / (1024 * 1024)
                        self.log_handler.info(f"  {index}. {file_path.name} ({size_mb:.2f} MB)")

                self.log_handler.success("=" * 50)
                self.log_handler.success("测试完成！邮件已成功处理")
                self.log_handler.success("=" * 50)
            except Exception as exc:
                import traceback

                self.log_handler.error(f"测试过程出错: {exc}")
                self.log_handler.error(f"详细错误:\n{traceback.format_exc()}")

        threading.Thread(target=run_test, daemon=True).start()
