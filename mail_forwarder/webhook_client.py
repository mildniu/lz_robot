from pathlib import Path
from urllib.parse import parse_qs, urlencode, urlparse, urlunparse

import requests


class QuantumWebhookClient:
    def __init__(self, send_url: str, upload_url: str) -> None:
        self.send_url = send_url
        self.upload_url = upload_url

    def send_text_alert(self, content: str) -> None:
        payload = {"type": "text", "textMsg": {"content": content}}
        response = requests.post(self.send_url, json=payload, timeout=20)
        response.raise_for_status()

    def _build_upload_url(self, upload_type: int) -> str:
        parsed = urlparse(self.upload_url)
        query = parse_qs(parsed.query, keep_blank_values=True)
        query["type"] = [str(upload_type)]
        return urlunparse(parsed._replace(query=urlencode(query, doseq=True)))

    def upload_file_with_meta(self, file_path: Path, upload_type: int = 2) -> tuple[str, dict]:
        upload_url = self._build_upload_url(upload_type)
        with file_path.open("rb") as file_handle:
            response = requests.post(
                upload_url,
                files={"file": (file_path.name, file_handle)},
                timeout=60,
            )
        response.raise_for_status()
        body = response.json()
        if not body.get("ok", False):
            raise RuntimeError(f"Upload failed: {body}")

        file_id = ((body.get("data") or {}).get("id")) or body.get("id")
        if not file_id:
            raise RuntimeError(f"Upload response missing file id: {body}")
        return str(file_id), body

    def upload_file(self, file_path: Path) -> str:
        file_id, _ = self.upload_file_with_meta(file_path, upload_type=2)
        return file_id

    def send_file_message(self, file_id: str) -> None:
        payload = {"type": "file", "fileMsg": {"fileId": file_id}}
        response = requests.post(self.send_url, json=payload, timeout=20)
        response.raise_for_status()
        body = response.json()
        if not body.get("ok", False):
            raise RuntimeError(f"Send file message failed: {body}")

    def send_image_message(self, file_id: str, height: int, width: int) -> None:
        payload = {
            "type": "image",
            "imageMsg": {
                "fileId": file_id,
                "height": int(height),
                "width": int(width),
            },
        }
        response = requests.post(self.send_url, json=payload, timeout=20)
        response.raise_for_status()
        body = response.json()
        if not body.get("ok", False):
            raise RuntimeError(f"Send image message failed: {body}")
