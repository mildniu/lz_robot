#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""量子推送机器人 - PySide6 桌面入口"""

from __future__ import annotations

import os
import sys
from pathlib import Path

from PySide6.QtGui import QFont, QGuiApplication, QIcon
from PySide6.QtWidgets import QApplication, QHBoxLayout, QMainWindow, QMessageBox, QStackedWidget, QWidget

from app_services import LogBus
from mail_forwarder import load_config
from mail_forwarder.config import upsert_env_file
from qt_components import NavigationSidebar
from qt_pages import AboutPage, BotTestPage, ExecutePage, FolderMonitorPage, SettingsPage

APP_TITLE = "量子推送机器人 v5.1 (PySide6)"
WINDOWS_APP_ID = "QuantumTelecom.LZRobot.5.1.Qt"


def project_root() -> Path:
    return Path(__file__).resolve().parent


def runtime_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return project_root()


def ensure_stable_working_directory() -> Path:
    base_dir = runtime_base_dir()
    try:
        os.chdir(base_dir)
    except OSError:
        pass
    return base_dir


def apply_windows_app_id() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes

        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(WINDOWS_APP_ID)
    except Exception:
        pass


def resource_path(relative_path: str) -> Path:
    base_dir = Path(getattr(sys, "_MEIPASS", runtime_base_dir()))
    return base_dir / relative_path


def build_app_stylesheet() -> str:
    return """
    QWidget {
        background: #EEF3F9;
        color: #10233A;
        font-size: 13px;
    }
    QMainWindow {
        background: #EEF4FA;
    }
    QLabel {
        background: transparent;
    }
    QWidget#PageShell {
        background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #F8FBFE, stop:0.52 #EFF4FA, stop:1 #E9F0F8);
    }
    QFrame#PageHeaderCard, QFrame#SectionCard, QFrame#PanelCard {
        border-radius: 20px;
    }
    QFrame#InnerPanelCard {
        border-radius: 16px;
        background: #F8FBFF;
        border: 1px solid #DEE8F4;
    }
    QFrame#MetricCard {
        border-radius: 18px;
        background: #FFFFFF;
        border: 1px solid #DCE7F3;
    }
    QFrame#MetricCard[metricTone="info"] {
        background: #F3F8FF;
        border: 1px solid #D8E5FB;
    }
    QFrame#MetricCard[metricTone="success"] {
        background: #F2FBF5;
        border: 1px solid #D5ECD9;
    }
    QFrame#MetricCard[metricTone="warning"] {
        background: #FFF9EF;
        border: 1px solid #F1DEC0;
    }
    QFrame#ActionStrip {
        background: #FFFFFF;
        border: 1px solid #DDE7F2;
        border-radius: 16px;
    }
    QFrame#PageHeaderCard {
        background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #FFFFFF, stop:1 #F4F8FE);
        border: 1px solid #D9E5F3;
    }
    QFrame#SectionCard {
        background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #F7FAFF, stop:1 #EDF4FF);
        border: 1px solid #D9E5F6;
    }
    QFrame#PanelCard {
        background: #FFFFFF;
        border: 1px solid #DDE7F2;
    }
    QFrame#PanelCardSoft {
        background: #F9FBFE;
        border: 1px solid #DEE8F4;
        border-radius: 16px;
    }
    QLabel#PageTitle {
        color: #10233A;
        font-size: 31px;
        font-weight: 700;
    }
    QLabel#PageSubtitle {
        color: #61748E;
        font-size: 13px;
    }
    QLabel#SectionTitle {
        color: #11253D;
        font-size: 16px;
        font-weight: 700;
    }
    QLabel#SectionHint {
        color: #667A92;
        font-size: 12px;
        line-height: 1.55;
    }
    QLabel#MicroTitle {
        color: #18304C;
        font-size: 13px;
        font-weight: 700;
    }
    QLabel#MetricTitle {
        color: #6A7C92;
        font-size: 12px;
        font-weight: 600;
    }
    QLabel#MetricValue {
        color: #0F2640;
        font-size: 22px;
        font-weight: 700;
    }
    QLabel#CardTitle {
        color: #142B45;
        font-size: 15px;
        font-weight: 700;
    }
    QLabel#FieldLabel {
        color: #31465F;
        font-size: 12px;
        font-weight: 600;
        padding-right: 8px;
    }
    QLabel#StatusPill {
        min-height: 24px;
        padding: 1px 11px;
        border-radius: 13px;
        font-size: 11px;
        font-weight: 700;
    }
    QLabel#StatusPill[pillTone="neutral"] {
        background: #EEF4FB;
        color: #5F738B;
        border: 1px solid #D8E3F0;
    }
    QLabel#StatusPill[pillTone="success"] {
        background: #EBF8F0;
        color: #157347;
        border: 1px solid #CBEBD8;
    }
    QLabel#StatusPill[pillTone="warning"] {
        background: #FFF6E8;
        color: #B86A10;
        border: 1px solid #F0D5A8;
    }
    QLabel#StatusPill[pillTone="danger"] {
        background: #FFF1F2;
        color: #C13943;
        border: 1px solid #F2CDD2;
    }
    QLabel#StatusPill[pillTone="info"] {
        background: #EAF2FF;
        color: #2259C8;
        border: 1px solid #CBDAFA;
    }
    QGroupBox#RuleCard, QGroupBox#MonitorCard, QGroupBox#FormSection {
        font-weight: 700;
        border: 1px solid #D8E4F0;
        border-radius: 18px;
        margin-top: 14px;
        padding-top: 18px;
        background: #FFFFFF;
    }
    QGroupBox#RuleCard::title, QGroupBox#MonitorCard::title, QGroupBox#FormSection::title {
        subcontrol-origin: margin;
        left: 18px;
        padding: 0 8px;
        color: #10233A;
    }
    QPushButton {
        border: 1px solid #D3DDE9;
        border-radius: 12px;
        background: #FFFFFF;
        color: #10233A;
        padding: 8px 14px;
        font-weight: 600;
        font-size: 12px;
    }
    QPushButton:hover {
        background: #F1F7FF;
        border-color: #8DB5F2;
    }
    QPushButton:pressed {
        background: #E2EEFF;
    }
    QPushButton:disabled {
        color: #94A3B8;
        background: #F8FAFC;
        border-color: #E2E8F0;
    }
    QPushButton[variant="primary"] {
        background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #1E5FE8, stop:1 #2F7CF6);
        color: white;
        border-color: #2466E8;
    }
    QPushButton[variant="primary"]:hover {
        background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #1A54D6, stop:1 #266EEA);
        border-color: #1A54D6;
    }
    QPushButton[variant="danger"] {
        background: #FFF7F7;
        color: #BA2E2E;
        border-color: #F0CCCC;
    }
    QPushButton[variant="warn"] {
        background: #FFF8ED;
        color: #B86811;
        border-color: #F3D3A7;
    }
    QLineEdit, QTextEdit, QComboBox, QTabWidget::pane {
        background: #FFFFFF;
        border: 1px solid #D8E3F0;
        border-radius: 12px;
    }
    QLineEdit, QTextEdit, QComboBox {
        min-height: 24px;
        padding: 8px 12px;
        selection-background-color: #2563EB;
    }
    QLineEdit:focus, QTextEdit:focus, QComboBox:focus {
        border: 1px solid #6BA1F2;
    }
    QTextEdit {
        padding: 10px 12px;
        line-height: 1.45;
    }
    QTextEdit[logView="true"] {
        background: #FBFDFF;
        border: 1px solid #DCE7F3;
        border-radius: 14px;
        padding: 12px 14px;
        font-family: Consolas, "Microsoft YaHei UI";
        font-size: 12px;
        color: #19314D;
    }
    QComboBox {
        padding-right: 34px;
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #FFFFFF, stop:1 #F8FBFF);
    }
    QComboBox::drop-down {
        subcontrol-origin: padding;
        subcontrol-position: top right;
        width: 32px;
        border: none;
        border-left: 1px solid #E2EAF4;
        background: #F3F8FF;
        border-top-right-radius: 12px;
        border-bottom-right-radius: 12px;
    }
    QComboBox::down-arrow {
        width: 9px;
        height: 9px;
    }
    QComboBox:on {
        border-color: #8FB5F3;
    }
    QComboBox QAbstractItemView {
        background: #FFFFFF;
        border: 1px solid #D8E3F0;
        border-radius: 14px;
        padding: 10px 8px;
        selection-background-color: #E7F0FF;
        selection-color: #12345A;
        outline: 0;
        show-decoration-selected: 1;
    }
    QAbstractItemView::item {
        min-height: 30px;
        padding: 6px 10px;
        border-radius: 8px;
        margin: 2px 4px;
    }
    QTabWidget::pane {
        margin-top: 14px;
        padding: 14px;
        background: #FFFFFF;
        border: 1px solid #D8E4F0;
        border-radius: 18px;
    }
    QTabWidget::tab-bar {
        left: 2px;
    }
    QTabBar::tab {
        background: #E8F0FA;
        border: 1px solid #D8E4F2;
        padding: 12px 22px;
        margin-right: 8px;
        border-radius: 13px;
        color: #445771;
        font-weight: 600;
    }
    QTabBar::tab:selected {
        background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #1E5FE8, stop:1 #2E7AF3);
        color: white;
        border-color: #1E5FE8;
    }
    QTabBar::tab:hover:!selected {
        background: #EEF4FD;
        border-color: #C7D8EE;
    }
    QHeaderView::section {
        background: #EDF4FF;
        color: #3A4E68;
        border: none;
        border-bottom: 1px solid #DAE5F3;
        padding: 11px 8px;
        font-weight: 700;
    }
    QTableWidget {
        background: #FFFFFF;
        alternate-background-color: #F8FBFF;
        border: 1px solid #D8E4F0;
        border-radius: 16px;
        gridline-color: #E7EEF6;
    }
    QTableWidget::item {
        padding: 8px;
    }
    QTableCornerButton::section {
        background: #EDF4FF;
        border: none;
        border-bottom: 1px solid #DAE5F3;
    }
    QScrollArea, QSplitter, QGroupBox, QFrame {
        background: transparent;
    }
    QSplitter::handle {
        background: transparent;
        width: 12px;
        height: 12px;
    }
    QGroupBox {
        font-weight: 700;
        border: 1px solid #D8E4F0;
        border-radius: 16px;
        margin-top: 12px;
        padding-top: 16px;
        background: #FFFFFF;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        left: 16px;
        padding: 0 8px;
        color: #10233A;
    }
    QCheckBox {
        color: #30455E;
        spacing: 8px;
    }
    QCheckBox::indicator {
        width: 16px;
        height: 16px;
        border-radius: 5px;
        border: 1px solid #B8C7DA;
        background: #FFFFFF;
    }
    QCheckBox::indicator:checked {
        background: #2A72EF;
        border: 1px solid #2A72EF;
    }
    QScrollBar:vertical {
        background: transparent;
        width: 12px;
        margin: 2px 2px 2px 2px;
    }
    QScrollBar::handle:vertical {
        background: #C7D6E6;
        min-height: 36px;
        border-radius: 6px;
    }
    QScrollBar::handle:vertical:hover {
        background: #A9C0DD;
    }
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical,
    QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
        background: none;
        border: none;
        height: 0px;
    }
    """


class QuantumMainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.runtime_base_dir = ensure_stable_working_directory()
        self.config = load_config()
        self.log_bus = LogBus()
        self.pages: dict[str, QWidget] = {}
        self._last_normal_size = (self.config.window_width, self.config.window_height)
        self._size_tracking_enabled = False

        self.setWindowTitle(self.config.app_title or APP_TITLE)
        self.resize(self.config.window_width, self.config.window_height)
        self._apply_icons()
        self._build_ui()
        self._initial_setup_window()

    def _apply_icons(self) -> None:
        ico_path = resource_path("icon/ico_quantum_telecom.ico")
        if ico_path.exists():
            icon = QIcon(str(ico_path))
            self.setWindowIcon(icon)
            QGuiApplication.setWindowIcon(icon)

    def _initial_setup_window(self) -> None:
        self.resize(self.config.window_width, self.config.window_height)
        self._last_normal_size = (self.config.window_width, self.config.window_height)
        self._size_tracking_enabled = False

    def _build_ui(self) -> None:
        root = QWidget(self)
        layout = QHBoxLayout(root)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.sidebar = NavigationSidebar(
            items=[
                ("execute", "邮件检测"),
                ("folder", "文件夹检测"),
                ("bot_test", "机器人测试"),
                ("settings", "设置"),
                ("about", "关于"),
            ],
            on_selected=self.show_page,
            width=self.config.sidebar_width,
            footer_text=self.config.app_footer_text or "v5.1\nPySide6",
        )
        layout.addWidget(self.sidebar)

        self.stack = QStackedWidget(root)
        layout.addWidget(self.stack, 1)
        self.setCentralWidget(root)

        self.pages = {
            "execute": ExecutePage(self.log_bus),
            "folder": FolderMonitorPage(self.log_bus),
            "bot_test": BotTestPage(self.log_bus),
            "settings": SettingsPage(self.log_bus),
            "about": AboutPage(self.log_bus),
        }
        settings_page = self.pages.get("settings")
        if settings_page and hasattr(settings_page, "config_changed"):
            settings_page.config_changed.connect(self._notify_external_config_updated)
        for page_id, page in self.pages.items():
            self.stack.addWidget(page)
            self.sidebar.register_page(page_id)

        self.show_page(self.config.start_page)

    def showEvent(self, event) -> None:  # type: ignore[override]
        super().showEvent(event)
        if not self._size_tracking_enabled:
            self._size_tracking_enabled = True

    def closeEvent(self, event) -> None:  # type: ignore[override]
        self._save_window_geometry()
        super().closeEvent(event)

    def show_page(self, page_id: str) -> None:
        page = self.pages.get(page_id)
        if not page:
            return
        self.stack.setCurrentWidget(page)
        self.sidebar.set_active(page_id)
        activated = getattr(page, "on_page_activated", None)
        if callable(activated):
            activated()

    def _notify_external_config_updated(self) -> None:
        self.config = load_config()
        self.setWindowTitle(self.config.app_title or APP_TITLE)
        for page_id, page in self.pages.items():
            if page_id == "settings":
                continue
            updated = getattr(page, "on_external_config_updated", None)
            if callable(updated):
                updated()

    def resizeEvent(self, event) -> None:  # type: ignore[override]
        super().resizeEvent(event)
        if not self._size_tracking_enabled:
            return
        if self.isMaximized() or self.isMinimized():
            return
        width = self.width()
        height = self.height()
        if width > 0 and height > 0:
            self._last_normal_size = (width, height)

    def _save_window_geometry(self) -> None:
        try:
            if self.isMinimized():
                return
            if not self.isMaximized():
                width = max(640, int(self.width()))
                height = max(480, int(self.height()))
                self._last_normal_size = (width, height)
            width, height = self._last_normal_size
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
            pass


def main() -> int:
    ensure_stable_working_directory()
    apply_windows_app_id()
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setFont(QFont("Microsoft YaHei UI", 10))
    app.setStyleSheet(build_app_stylesheet())
    window = QuantumMainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except ModuleNotFoundError as exc:
        QMessageBox.critical(None, "缺少依赖", f"缺少 PySide6 依赖：{exc}")
        raise
