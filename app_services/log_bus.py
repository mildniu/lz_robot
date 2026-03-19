from __future__ import annotations

from PySide6.QtCore import QObject, Signal


class LogBus(QObject):
    log_emitted = Signal(str, str, str)

    def emit(self, level: str, message: str, source: str = "global") -> None:
        self.log_emitted.emit(level, message, source)
