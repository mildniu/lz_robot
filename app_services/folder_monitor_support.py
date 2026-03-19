from __future__ import annotations

import hashlib
import json
import threading
from datetime import datetime
from pathlib import Path

from watchdog.events import FileSystemEventHandler


class FileSentTracker:
    def __init__(self, state_file: str = "state/file_sent_state.json"):
        self.state_file = Path(state_file)
        self.state_file.parent.mkdir(parents=True, exist_ok=True)
        self._lock = threading.Lock()
        self.sent_files = self.load_state()

    def load_state(self) -> dict[str, dict[str, str]]:
        if self.state_file.exists():
            try:
                return json.loads(self.state_file.read_text(encoding="utf-8"))
            except Exception:
                return {}
        return {}

    def save_state(self) -> None:
        temp_file = self.state_file.with_suffix(self.state_file.suffix + ".tmp")
        temp_file.write_text(
            json.dumps(self.sent_files, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        temp_file.replace(self.state_file)

    def get_file_hash(self, file_path: Path) -> str:
        hasher = hashlib.md5()
        with file_path.open("rb") as file_handle:
            for chunk in iter(lambda: file_handle.read(1024 * 1024), b""):
                hasher.update(chunk)
        return hasher.hexdigest()

    def is_sent(self, file_path: Path, webhook_url: str) -> bool:
        with self._lock:
            file_key = str(file_path)
            if file_key not in self.sent_files:
                return False

            record = self.sent_files[file_key]
            if record.get("webhook_url") != webhook_url:
                return False

            return record.get("file_hash") == self.get_file_hash(file_path)

    def mark_sent(self, file_path: Path, webhook_url: str, file_id: str) -> None:
        with self._lock:
            self.sent_files[str(file_path)] = {
                "file_hash": self.get_file_hash(file_path),
                "webhook_url": webhook_url,
                "file_id": file_id,
                "sent_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            self.save_state()


class FolderMonitorHandler(FileSystemEventHandler):
    def __init__(self, callback, log_handler=None, logger=None, source: str = "global"):
        super().__init__()
        self.callback = callback
        self.log_handler = log_handler or logger
        self.source = source

    def _emit_log(self, level: str, message: str) -> None:
        if not self.log_handler:
            return
        if callable(self.log_handler):
            self.log_handler(level, message, self.source)
            return
        log_method = getattr(self.log_handler, level.lower(), None)
        if callable(log_method):
            log_method(message, source=self.source)
            return
        generic_log = getattr(self.log_handler, "log", None)
        if callable(generic_log):
            generic_log(level, message, source=self.source)

    def on_created(self, event) -> None:
        if event.is_directory:
            return
        self._emit_log("INFO", f"检测到新文件: {event.src_path}")
        if self.callback:
            self.callback(event.src_path, "created")

    def on_modified(self, event) -> None:
        if event.is_directory:
            return
        if self.callback:
            self.callback(event.src_path, "modified")
