import hashlib
import json
import queue
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Dict

import customtkinter as ctk
from watchdog.events import FileSystemEventHandler


class LogHandler:
    """Log queue consumed by the Tk main thread."""

    def __init__(self):
        self.queue = queue.Queue()
        self.callbacks = []

    def add_callback(self, callback):
        self.callbacks.append(callback)

    def log(self, level: str, message: str, source: str = "global"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.queue.put({"time": timestamp, "level": level, "message": message, "source": source})

    def info(self, message: str, source: str = "global"):
        self.log("INFO", message, source=source)

    def success(self, message: str, source: str = "global"):
        self.log("SUCCESS", message, source=source)

    def warning(self, message: str, source: str = "global"):
        self.log("WARNING", message, source=source)

    def error(self, message: str, source: str = "global"):
        self.log("ERROR", message, source=source)

    def dispatch_pending(self):
        while True:
            try:
                log_entry = self.queue.get_nowait()
            except queue.Empty:
                break

            for callback in list(self.callbacks):
                try:
                    callback(log_entry)
                except Exception:
                    pass


class ModernButton(ctk.CTkButton):
    def __init__(self, *args, icon: str = "", **kwargs):
        super().__init__(*args, **kwargs)
        if icon:
            self.configure(text=f"{icon} {self.cget('text')}")


class FileSentTracker:
    def __init__(self, state_file: str = "state/file_sent_state.json"):
        self.state_file = Path(state_file)
        self.state_file.parent.mkdir(parents=True, exist_ok=True)
        self._lock = threading.Lock()
        self.sent_files = self.load_state()

    def load_state(self) -> Dict[str, Dict[str, str]]:
        if self.state_file.exists():
            try:
                return json.loads(self.state_file.read_text(encoding="utf-8"))
            except Exception:
                return {}
        return {}

    def save_state(self):
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

    def mark_sent(self, file_path: Path, webhook_url: str, file_id: str):
        with self._lock:
            self.sent_files[str(file_path)] = {
                "file_hash": self.get_file_hash(file_path),
                "webhook_url": webhook_url,
                "file_id": file_id,
                "sent_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            self.save_state()


class FolderMonitorHandler(FileSystemEventHandler):
    def __init__(self, callback, log_handler: LogHandler, source: str = "global"):
        super().__init__()
        self.callback = callback
        self.log_handler = log_handler
        self.source = source

    def on_created(self, event):
        if not event.is_directory:
            self.log_handler.info(f"检测到新文件: {event.src_path}", source=self.source)
            if self.callback:
                self.callback(event.src_path, "created")

    def on_modified(self, event):
        if not event.is_directory:
            if self.callback:
                self.callback(event.src_path, "modified")
