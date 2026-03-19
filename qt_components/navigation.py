from __future__ import annotations

from collections.abc import Callable

from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QPushButton, QVBoxLayout


class NavigationSidebar(QFrame):
    def __init__(
        self,
        items: list[tuple[str, str]],
        on_selected: Callable[[str], None],
        *,
        width: int = 220,
        footer_text: str = "v5.1\nPySide6",
    ) -> None:
        super().__init__()
        self.on_selected = on_selected
        self.buttons: dict[str, QPushButton] = {}
        self.item_meta: dict[str, tuple[str, str]] = {
            "execute": ("01", "邮件检测"),
            "folder": ("02", "文件夹检测"),
            "bot_test": ("03", "机器人测试"),
            "settings": ("04", "设置"),
            "about": ("05", "关于"),
        }
        self.setFixedWidth(width)
        self.setObjectName("NavigationSidebar")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 24, 18, 20)
        layout.setSpacing(14)

        brand = QFrame(self)
        brand.setObjectName("SidebarBrand")
        brand_layout = QVBoxLayout(brand)
        brand_layout.setContentsMargins(18, 20, 18, 20)
        brand_layout.setSpacing(10)
        layout.addWidget(brand)

        logo = QLabel(brand)
        logo.setFixedSize(64, 64)
        logo.setScaledContents(True)
        logo_path = self._resource_path("icon/logo_quantum_telecom.png")
        if logo_path.exists():
            logo.setPixmap(QPixmap(str(logo_path)))
        brand_layout.addWidget(logo, alignment=Qt.AlignCenter)

        title = QLabel("量子机器人", brand)
        title.setAlignment(Qt.AlignCenter)
        title.setObjectName("SidebarTitle")
        brand_layout.addWidget(title)

        for page_id, text in items:
            button = QPushButton(text, self)
            seq, label = self.item_meta.get(page_id, ("--", text))
            button.setText(f"{seq}  {label}")
            button.clicked.connect(lambda _checked=False, pid=page_id: self.on_selected(pid))
            button.setCheckable(True)
            button.setMinimumHeight(48)
            button.setObjectName("SidebarButton")
            self.buttons[page_id] = button
            layout.addWidget(button)

        divider = QFrame(self)
        divider.setObjectName("SidebarDivider")
        divider.setFixedHeight(1)
        layout.addWidget(divider)

        layout.addStretch(1)

        footer = QLabel(footer_text, self)
        footer.setAlignment(Qt.AlignLeft | Qt.AlignBottom)
        footer.setObjectName("SidebarFooter")
        layout.addWidget(footer)

        self.setStyleSheet(
            """
            QFrame#NavigationSidebar {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #12345B, stop:0.48 #163B65, stop:1 #1A436F);
                border-right: 1px solid rgba(255, 255, 255, 0.08);
            }
            QFrame#SidebarBrand {
                background: rgba(255, 255, 255, 0.10);
                border: 1px solid rgba(255, 255, 255, 0.13);
                border-radius: 20px;
            }
            QLabel#SidebarTitle { font-size: 20px; font-weight: 700; color: #FFFFFF; }
            QLabel#SidebarFooter { font-size: 12px; color: rgba(226, 236, 246, 0.70); padding-left: 6px; }
            QFrame#SidebarDivider { background: rgba(255, 255, 255, 0.12); margin: 6px 4px 0 4px; }
            QPushButton#SidebarButton {
                border: none;
                border-radius: 15px;
                background: transparent;
                text-align: left;
                padding: 13px 16px;
                color: #EDF4FD;
                font-size: 14px;
                font-weight: 600;
                letter-spacing: 0.3px;
            }
            QPushButton#SidebarButton:hover {
                background: rgba(255, 255, 255, 0.11);
            }
            QPushButton#SidebarButton:checked {
                background: #F8FBFF;
                color: #143055;
                font-weight: 700;
                padding-left: 21px;
                border-left: 4px solid #2A72EF;
            }
            QPushButton#SidebarButton:checked:hover {
                background: #F8FBFF;
            }
            """
        )

    def register_page(self, page_id: str) -> None:
        if page_id not in self.buttons:
            return

    @staticmethod
    def _resource_path(relative_path: str):
        from pathlib import Path

        return Path(__file__).resolve().parents[1] / relative_path

    def set_active(self, page_id: str) -> None:
        for current_id, button in self.buttons.items():
            button.setChecked(current_id == page_id)
