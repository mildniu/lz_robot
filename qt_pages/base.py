from __future__ import annotations

from PySide6.QtWidgets import QVBoxLayout, QWidget

from app_services import LogBus


class BasePage(QWidget):
    def __init__(self, log_bus: LogBus, title: str, subtitle: str = "") -> None:
        super().__init__()
        self.log_bus = log_bus
        self.setObjectName("PageShell")

        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(28, 24, 28, 24)
        self.layout.setSpacing(18)

    def on_page_activated(self) -> None:
        pass

    def on_external_config_updated(self) -> None:
        pass
