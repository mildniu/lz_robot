import json
from pathlib import Path
from typing import Dict


class JsonStateStore:
    def __init__(self, path: Path) -> None:
        self.path = path

    def read(self) -> Dict[str, str]:
        if not self.path.exists():
            return {}
        try:
            return json.loads(self.path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return {}

    def write(self, data: Dict[str, str]) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        tmp_path = self.path.with_suffix(self.path.suffix + ".tmp")
        tmp_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        tmp_path.replace(self.path)

    def get_last_sent_uid(self, key: str = "last_sent_uid") -> str:
        return str(self.read().get(key, "")).strip()

    def set_last_sent_uid(self, uid: str, key: str = "last_sent_uid") -> None:
        payload = self.read()
        payload[key] = uid
        self.write(payload)
