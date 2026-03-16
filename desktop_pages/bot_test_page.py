import threading
import tkinter as tk
import tkinter.messagebox as messagebox
from datetime import datetime
from pathlib import Path
from tkinter import filedialog

import customtkinter as ctk

from mail_forwarder.processing_service import build_webhook_client, send_file_via_webhook

from .common import ModernButton
from .webhook_alias_store import load_webhook_aliases


class BotTestPage(ctk.CTkFrame):
    NO_ALIAS_LABEL = "(未配置)"

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.alias_map = {}
        self.alias_var = tk.StringVar(value=self.NO_ALIAS_LABEL)
        self._sending = False
        self.setup_ui()
        self.refresh_aliases()

    def setup_ui(self):
        control_bar = ctk.CTkFrame(self, height=100)
        control_bar.pack(fill="x", padx=25, pady=(20, 10))

        left_panel = ctk.CTkFrame(control_bar, fg_color="transparent")
        left_panel.pack(side="left", fill="both", expand=True)

        title = ctk.CTkLabel(left_panel, text="🤖 机器人测试", font=ctk.CTkFont(size=26, weight="bold"))
        title.pack(pady=(5, 5), anchor="w")

        subtitle = ctk.CTkLabel(
            left_panel,
            text="选择已配置机器人，测试文字/图片/文件推送",
            font=ctk.CTkFont(size=14),
            text_color=("gray35", "gray70"),
        )
        subtitle.pack(anchor="w")

        alias_row = ctk.CTkFrame(self, fg_color="transparent")
        alias_row.pack(fill="x", padx=25, pady=(0, 10))

        ctk.CTkLabel(alias_row, text="推送机器人", width=90, anchor="w", font=ctk.CTkFont(size=13)).pack(
            side="left", padx=(0, 8)
        )
        self.alias_menu = ctk.CTkOptionMenu(
            alias_row,
            values=[self.NO_ALIAS_LABEL],
            variable=self.alias_var,
            width=260,
            height=34,
        )
        self.alias_menu.pack(side="left")

        refresh_btn = ctk.CTkButton(
            alias_row,
            text="重新读取",
            command=self.refresh_aliases,
            width=82,
            height=30,
            font=ctk.CTkFont(size=11),
            fg_color=("gray88", "gray24"),
            hover_color=("gray80", "gray30"),
            text_color=("gray25", "gray85"),
            border_width=1,
            border_color=("gray78", "gray38"),
        )
        refresh_btn.pack(side="left", padx=8)

        text_card = ctk.CTkFrame(self)
        text_card.pack(fill="x", padx=25, pady=(0, 10))
        ctk.CTkLabel(text_card, text="文字测试", font=ctk.CTkFont(size=16, weight="bold")).pack(
            anchor="w", padx=12, pady=(10, 6)
        )

        self.text_input = ctk.CTkTextbox(text_card, height=90, font=ctk.CTkFont(size=13))
        self.text_input.pack(fill="x", padx=12, pady=(0, 10))
        self.text_input.insert("1.0", "这是一条机器人文字测试消息。")

        self.send_text_btn = ModernButton(
            text_card,
            text="发送文字测试",
            icon="✉️",
            command=self.send_text_test,
            width=160,
            height=38,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#42A5F5",
            hover_color="#1976D2",
        )
        self.send_text_btn.pack(anchor="w", padx=12, pady=(0, 12))

        file_card = ctk.CTkFrame(self)
        file_card.pack(fill="x", padx=25, pady=(0, 10))
        ctk.CTkLabel(file_card, text="文件/图片测试", font=ctk.CTkFont(size=16, weight="bold")).pack(
            anchor="w", padx=12, pady=(10, 6)
        )

        file_row = ctk.CTkFrame(file_card, fg_color="transparent")
        file_row.pack(fill="x", padx=12, pady=(0, 10))
        self.file_entry = ctk.CTkEntry(file_row, placeholder_text="请选择本地文件路径", height=34)
        self.file_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        browse_btn = ctk.CTkButton(
            file_row,
            text="选择文件",
            command=self.choose_file,
            width=90,
            height=34,
            font=ctk.CTkFont(size=12),
        )
        browse_btn.pack(side="left")

        self.send_file_btn = ModernButton(
            file_card,
            text="发送文件/图片测试",
            icon="📎",
            command=self.send_file_test,
            width=180,
            height=38,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#2CC985",
            hover_color="#239B6D",
        )
        self.send_file_btn.pack(anchor="w", padx=12, pady=(0, 12))

        log_card = ctk.CTkFrame(self)
        log_card.pack(fill="both", expand=True, padx=25, pady=(0, 20))
        log_header = ctk.CTkFrame(log_card, height=40)
        log_header.pack(fill="x")
        ctk.CTkLabel(log_header, text="测试日志", font=ctk.CTkFont(size=16, weight="bold")).pack(
            side="left", padx=10, pady=8
        )
        ctk.CTkButton(
            log_header,
            text="清空",
            command=self.clear_log,
            width=70,
            height=30,
            font=ctk.CTkFont(size=12),
        ).pack(side="right", padx=10, pady=5)
        ctk.CTkButton(
            log_header,
            text="导出",
            command=self.export_log,
            width=70,
            height=30,
            font=ctk.CTkFont(size=12),
        ).pack(side="right", padx=(0, 8), pady=5)

        self.log_text = ctk.CTkTextbox(log_card, font=ctk.CTkFont(family="Consolas", size=12))
        self.log_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def append_log(self, level: str, message: str):
        icons = {"INFO": "ℹ️", "SUCCESS": "✅", "WARNING": "⚠️", "ERROR": "❌"}
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{timestamp}] {icons.get(level, '•')} {message}\n")
        self.log_text.see("end")

    def clear_log(self):
        self.log_text.delete("1.0", "end")

    def export_log(self):
        content = self.log_text.get("1.0", "end").strip()
        if not content:
            messagebox.showwarning("提示", "当前日志为空，暂无可导出的内容")
            return
        export_dir = Path("exports")
        export_dir.mkdir(parents=True, exist_ok=True)
        file_path = filedialog.asksaveasfilename(
            title="导出机器人测试日志",
            initialdir=str(export_dir.resolve()),
            initialfile=f"bot_test_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log",
            defaultextension=".log",
            filetypes=[("Log Files", "*.log"), ("Text Files", "*.txt"), ("All Files", "*.*")],
        )
        if not file_path:
            return
        Path(file_path).write_text(content + "\n", encoding="utf-8")
        messagebox.showinfo("导出成功", f"日志已导出到:\n{file_path}")

    def refresh_aliases(self, log_result: bool = True):
        config = load_webhook_aliases()
        aliases = config.get("aliases", {})
        self.alias_map = {k: v.strip() for k, v in aliases.items() if str(k).strip() and str(v).strip()}

        values = [self.NO_ALIAS_LABEL] + sorted(self.alias_map.keys())
        current = self.alias_var.get().strip()
        self.alias_menu.configure(values=values)
        if current not in values:
            self.alias_var.set(values[1] if len(values) > 1 else values[0])
        if log_result:
            self.append_log("INFO", f"已加载机器人别名数量: {len(self.alias_map)}")

    def on_page_activated(self):
        self.refresh_aliases(log_result=False)

    def on_external_config_updated(self):
        self.refresh_aliases(log_result=True)

    def choose_file(self):
        root = tk.Tk()
        root.withdraw()
        root.wm_attributes("-topmost", 1)
        file_path = filedialog.askopenfilename(title="选择测试文件")
        root.destroy()
        if file_path:
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, file_path)

    def _get_selected_webhook(self) -> tuple[str, str]:
        alias = self.alias_var.get().strip()
        if not alias or alias == self.NO_ALIAS_LABEL:
            raise ValueError("请先选择机器人别名")
        webhook_url = self.alias_map.get(alias, "").strip()
        if not webhook_url:
            raise ValueError(f"机器人别名无效: {alias}")
        if not webhook_url.startswith(("http://", "https://")):
            raise ValueError(f"Webhook 地址格式无效: {webhook_url}")
        return alias, webhook_url

    def _set_sending(self, sending: bool):
        self._sending = sending
        state = "disabled" if sending else "normal"
        self.send_text_btn.configure(state=state)
        self.send_file_btn.configure(state=state)

    def _run_async(self, action: str, worker):
        if self._sending:
            self.append_log("WARNING", "当前有测试任务执行中，请稍后")
            return

        self._set_sending(True)
        self.append_log("INFO", f"开始执行: {action}")

        def _task():
            try:
                worker()
            except Exception as exc:
                self.after(0, lambda: self.append_log("ERROR", f"{action}失败: {exc}"))
            finally:
                self.after(0, lambda: self._set_sending(False))

        threading.Thread(target=_task, daemon=True).start()

    def send_text_test(self):
        content = self.text_input.get("1.0", "end").strip()
        if not content:
            self.append_log("WARNING", "文字内容为空，请先填写测试文本")
            return

        try:
            alias, webhook_url = self._get_selected_webhook()
        except Exception as exc:
            self.append_log("ERROR", str(exc))
            return

        def worker():
            self.after(0, lambda: self.append_log("INFO", f"目标机器人: {alias}"))
            webhook = build_webhook_client(webhook_url)
            webhook.send_text_alert(content)
            self.after(0, lambda: self.append_log("SUCCESS", "文字消息发送成功"))

        self._run_async("文字测试", worker)

    def send_file_test(self):
        raw_path = self.file_entry.get().strip()
        if not raw_path:
            self.append_log("WARNING", "请先选择测试文件")
            return

        file_path = Path(raw_path)
        if not file_path.exists() or not file_path.is_file():
            self.append_log("ERROR", f"文件不存在: {file_path}")
            return

        try:
            alias, webhook_url = self._get_selected_webhook()
        except Exception as exc:
            self.append_log("ERROR", str(exc))
            return

        def worker():
            self.after(0, lambda: self.append_log("INFO", f"目标机器人: {alias}"))
            self.after(0, lambda: self.append_log("INFO", f"发送文件: {file_path.name}"))

            def event_callback(level: str, message: str):
                self.after(0, lambda: self.append_log(level, message))

            send_file_via_webhook(file_path, webhook_url, event_callback=event_callback)
            self.after(0, lambda: self.append_log("SUCCESS", "文件/图片测试发送完成"))

        self._run_async("文件/图片测试", worker)
