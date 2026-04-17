#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""量子推送机器人 - 桌面程序入口"""

from pathlib import Path
import os
import sys
import threading
import tkinter as tk

import customtkinter as ctk
try:
    from PIL import Image
except Exception:  # pragma: no cover - runtime fallback for missing pillow
    Image = None
try:
    import pystray
except Exception:  # pragma: no cover - runtime fallback for missing pystray
    pystray = None

from desktop_pages import AboutPage, BotTestPage, ExecutePage, FolderMonitorPage, LogHandler, SettingsPage
from mail_forwarder.config import load_config, upsert_env_file

APP_TITLE = "量子推送机器人 v5.2"
APP_FOOTER_TEXT = "v5.2\nby 不丢西瓜der"
WINDOWS_APP_ID = "QuantumTelecom.LZRobot.5.2"


def runtime_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def ensure_stable_working_directory() -> Path:
    base_dir = runtime_base_dir()
    try:
        os.chdir(base_dir)
    except OSError:
        pass
    return base_dir


def resource_path(relative_path: str) -> Path:
    base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base_dir / relative_path


def apply_windows_app_id() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes

        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(WINDOWS_APP_ID)
    except Exception:
        pass


class ModernApp(ctk.CTk):
    """桌面应用壳，只负责窗口、导航和页面切换。"""

    def __init__(self):
        super().__init__()
        self.runtime_base_dir = ensure_stable_working_directory()
        self.config = load_config()
        self._initial_geometry = f"{self.config.window_width}x{self.config.window_height}"
        self.title(APP_TITLE)
        self._last_normal_size = (self.config.window_width, self.config.window_height)
        self._window_icon_ref = None
        self._size_tracking_enabled = False
        self._tray_icon = None
        self._tray_thread = None
        self._tray_notice_shown = False
        self._force_exit = False

        # Use a CJK-friendly default UI font on Windows for better Chinese rendering.
        try:
            self.option_add("*Font", "{Microsoft YaHei UI} 10")
        except Exception:
            pass

        ctk.set_appearance_mode(self.config.ui_appearance)
        try:
            ctk.set_default_color_theme(self.config.ui_color_theme)
        except Exception:
            ctk.set_default_color_theme("blue")
        try:
            ctk.set_widget_scaling(self.config.ui_scale)
        except Exception:
            ctk.set_widget_scaling(1.0)
        self.geometry(self._initial_geometry)

        self.email_log_handler = LogHandler()
        self.folder_log_handler = LogHandler()
        self.pages = {}
        self.logo_image = None

        self.create_ui()
        self._apply_window_icon()
        self._setup_tray_icon()
        self.bind("<Configure>", self._on_configure)
        self.after(0, self._finalize_initial_geometry)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.process_logs()

    def _finalize_initial_geometry(self):
        try:
            # Ensure startup layout/scale changes do not pollute persisted size.
            self.geometry(self._initial_geometry)
            self.update_idletasks()
        except Exception:
            pass
        finally:
            # Delay enabling tracking to skip startup layout jitter.
            self.after(600, self._enable_size_tracking)

    def _enable_size_tracking(self):
        self._size_tracking_enabled = True

    def _read_current_geometry_size(self):
        try:
            geometry_text = str(self.geometry()).split("+", 1)[0]
            width_text, height_text = geometry_text.split("x", 1)
            width = int(width_text)
            height = int(height_text)
            if width > 0 and height > 0:
                return width, height
        except Exception:
            return None
        return None

    def _apply_window_icon(self):
        ico_path = resource_path("icon/ico_quantum_telecom.ico")
        png_path = resource_path("icon/logo_quantum_telecom.png")

        if sys.platform.startswith("win") and ico_path.exists():
            try:
                self.iconbitmap(str(ico_path))
            except Exception:
                pass

        if png_path.exists():
            try:
                self._window_icon_ref = tk.PhotoImage(file=str(png_path))
                self.iconphoto(True, self._window_icon_ref)
            except Exception:
                self._window_icon_ref = None

    def _on_configure(self, _event=None):
        try:
            if not self._size_tracking_enabled:
                return
            # 只记录 normal 状态，避免最小化/最大化导致的尺寸污染
            if self.state() != "normal":
                return
            size = self._read_current_geometry_size()
            if size:
                self._last_normal_size = size
        except Exception:
            pass

    def _load_tray_image(self):
        if Image is None:
            return None
        icon_path = resource_path("icon/logo_quantum_telecom.png")
        if not icon_path.exists():
            return None
        try:
            return Image.open(icon_path).convert("RGBA")
        except Exception:
            return None

    def _setup_tray_icon(self):
        if pystray is None:
            return
        tray_image = self._load_tray_image()
        if tray_image is None:
            return
        try:
            menu = pystray.Menu(
                pystray.MenuItem("Show Window", lambda icon, item: self.after(0, self.restore_from_tray), default=True),
                pystray.MenuItem("Exit", lambda icon, item: self.after(0, self.quit_from_tray)),
            )
            self._tray_icon = pystray.Icon("QuantumBotTray", tray_image, APP_TITLE, menu)
            self._tray_thread = threading.Thread(target=self._tray_icon.run, daemon=True)
            self._tray_thread.start()
        except Exception:
            self._tray_icon = None
            self._tray_thread = None

    def restore_from_tray(self):
        try:
            self.deiconify()
            self.lift()
            self.focus_force()
        except Exception:
            pass

    def _show_tray_notice_once(self):
        if self._tray_notice_shown or self._tray_icon is None:
            return
        try:
            self._tray_icon.notify("QuantumBot is still running in the system tray.", APP_TITLE)
            self._tray_notice_shown = True
        except Exception:
            pass

    def quit_from_tray(self):
        self._force_exit = True
        self.on_close()

    def _shutdown_tray_icon(self):
        tray_icon = self._tray_icon
        self._tray_icon = None
        if tray_icon is None:
            return
        try:
            tray_icon.stop()
        except Exception:
            pass

    def _save_window_geometry(self):
        try:
            if self.state() == "normal":
                size = self._read_current_geometry_size()
                if size:
                    self._last_normal_size = size

            width, height = self._last_normal_size
            width = max(640, int(width))
            height = max(480, int(height))
            if width == self.config.window_width and height == self.config.window_height:
                return
            upsert_env_file(
                Path("settings/app_config.json"),
                {
                    "WINDOW_WIDTH": str(width),
                    "WINDOW_HEIGHT": str(height),
                },
            )
        except Exception:
            # 窗口大小持久化失败不影响主流程退出
            pass

    def create_ui(self):
        main_container = ctk.CTkFrame(self)
        main_container.pack(fill="both", expand=True)

        sidebar = ctk.CTkFrame(main_container, width=self.config.sidebar_width, corner_radius=0)
        sidebar.pack(side="left", fill="y")
        self.create_sidebar(sidebar)

        self.content_area = ctk.CTkFrame(main_container)
        self.content_area.pack(side="right", fill="both", expand=True)

        self.create_pages()
        self.show_page(self.config.start_page)

    def create_sidebar(self, sidebar):
        logo_frame = ctk.CTkFrame(sidebar, height=150)
        logo_frame.pack(fill="x", padx=10, pady=(30, 20))

        logo_path = resource_path("icon/logo_quantum_telecom.png")
        if Image and logo_path.exists():
            logo = Image.open(logo_path)
            self.logo_image = ctk.CTkImage(light_image=logo, dark_image=logo, size=(72, 72))
            ctk.CTkLabel(logo_frame, text="", image=self.logo_image).pack()
        else:
            ctk.CTkLabel(logo_frame, text="📧", font=ctk.CTkFont(size=50)).pack()
        ctk.CTkLabel(
            logo_frame,
            text="量子机器人",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).pack(pady=(10, 0))

        nav_buttons = [
            ("📧", "邮件检测", "execute"),
            ("📁", "文件夹检测", "folder"),
            ("🤖", "机器人测试", "bot_test"),
            ("⚙️", "设置", "settings"),
            ("ℹ️", "关于", "about"),
        ]
        for icon, text, page_id in nav_buttons:
            button = ctk.CTkButton(
                sidebar,
                text=f"{icon}  {text}",
                command=lambda target=page_id: self.show_page(target),
                height=50,
                font=ctk.CTkFont(size=15),
                corner_radius=10,
                fg_color="transparent",
                text_color=("gray10", "#DCE4EE"),
                hover_color=("gray70", "gray25"),
            )
            button.pack(fill="x", padx=15, pady=8)
            setattr(self, f"{page_id}_btn", button)

        ctk.CTkLabel(
            sidebar,
            text=APP_FOOTER_TEXT,
            font=ctk.CTkFont(size=12),
            text_color="gray",
        ).pack(side="bottom", pady=20)

    def create_pages(self):
        self.pages["execute"] = ExecutePage(
            self.content_area,
            self.email_log_handler,
            auto_scroll_log=self.config.auto_scroll_log,
        )
        self.pages["folder"] = FolderMonitorPage(
            self.content_area,
            self.folder_log_handler,
            auto_scroll_log=self.config.auto_scroll_log,
        )
        self.pages["bot_test"] = BotTestPage(self.content_area)
        self.pages["settings"] = SettingsPage(
            self.content_area,
            on_config_changed=self.notify_config_changed,
        )
        self.pages["about"] = AboutPage(self.content_area)

    def show_page(self, page_id: str):
        for page in self.pages.values():
            page.pack_forget()

        current_page = self.pages[page_id]
        current_page.pack(fill="both", expand=True)
        page_activated = getattr(current_page, "on_page_activated", None)
        if callable(page_activated):
            page_activated()

        for pid in ["execute", "folder", "bot_test", "settings", "about"]:
            button = getattr(self, f"{pid}_btn", None)
            if not button:
                continue
            if pid == page_id:
                button.configure(
                    fg_color=("gray70", "gray25"),
                    text_color=("gray10", "white"),
                )
            else:
                button.configure(
                    fg_color="transparent",
                    text_color=("gray10", "#DCE4EE"),
                )

    def notify_config_changed(self):
        self.config = load_config()
        for page in self.pages.values():
            refresh_func = getattr(page, "on_external_config_updated", None)
            if callable(refresh_func):
                refresh_func()

    def process_logs(self):
        self.email_log_handler.dispatch_pending()
        self.folder_log_handler.dispatch_pending()
        self.after(self.config.ui_log_poll_ms, self.process_logs)

    def on_close(self):
        if not self._force_exit and self._tray_icon is not None:
            self._save_window_geometry()
            try:
                self.withdraw()
            except Exception:
                self.iconify()
            self._show_tray_notice_once()
            return

        self._save_window_geometry()

        execute_page = self.pages.get("execute")
        folder_page = self.pages.get("folder")

        if execute_page and getattr(execute_page, "is_running", False):
            execute_page.stop_worker()
        if folder_page and getattr(folder_page, "is_running", False):
            folder_page.stop_monitor()

        self._shutdown_tray_icon()
        self.destroy()


def main():
    try:
        import customtkinter  # noqa: F401
    except ImportError:
        print("错误: 未安装 CustomTkinter")
        print("请运行: pip install customtkinter")
        return

    apply_windows_app_id()
    app = ModernApp()
    app.mainloop()


if __name__ == "__main__":
    main()
