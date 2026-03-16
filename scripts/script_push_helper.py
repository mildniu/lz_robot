#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""规则脚本推送助手。

给规则脚本使用，自动从环境变量中读取当前规则选中的机器人信息:
    LZ_WEBHOOK_URL
    LZ_WEBHOOK_UPLOAD_URL
    LZ_WEBHOOK_ALIAS

典型用法:
    from script_push_helper import ScriptPushClient
    client = ScriptPushClient.from_env()
    client.send_text("处理完成")
    client.send_file(Path("result.xlsx"))
    client.send_image(Path("result.png"))
"""

from __future__ import annotations

import json
import os
import struct
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import requests


def parse_upload_url(send_url: str) -> str:
    parsed = urlparse(send_url)
    key = parse_qs(parsed.query).get("key", [""])[0]
    if not key:
        raise ValueError("Webhook URL 缺少 key 参数")
    return f"{parsed.scheme}://{parsed.netloc}/im-external/v1/webhook/upload-attachment?key={key}&type=2"


def is_image_file(file_path: Path) -> bool:
    return file_path.suffix.lower() in {".png", ".jpg", ".jpeg"}


def get_image_dimensions(file_path: Path) -> tuple[int, int]:
    suffix = file_path.suffix.lower()
    if suffix == ".png":
        with file_path.open("rb") as file_handle:
            header = file_handle.read(24)
        if len(header) < 24 or header[:8] != b"\x89PNG\r\n\x1a\n":
            raise RuntimeError(f"Invalid PNG file: {file_path.name}")
        width, height = struct.unpack(">II", header[16:24])
        return int(width), int(height)

    if suffix in {".jpg", ".jpeg"}:
        with file_path.open("rb") as file_handle:
            data = file_handle.read()
        if len(data) < 4 or data[0:2] != b"\xff\xd8":
            raise RuntimeError(f"Invalid JPEG file: {file_path.name}")

        index = 2
        data_len = len(data)
        while index + 9 < data_len:
            if data[index] != 0xFF:
                index += 1
                continue
            marker = data[index + 1]
            index += 2
            if marker in {0xD8, 0xD9, 0x01} or 0xD0 <= marker <= 0xD7:
                continue
            if index + 2 > data_len:
                break
            segment_length = struct.unpack(">H", data[index : index + 2])[0]
            if segment_length < 2 or index + segment_length > data_len:
                break
            if marker in {0xC0, 0xC1, 0xC2, 0xC3, 0xC9, 0xCA, 0xCB}:
                if index + 7 > data_len:
                    break
                height = struct.unpack(">H", data[index + 3 : index + 5])[0]
                width = struct.unpack(">H", data[index + 5 : index + 7])[0]
                return int(width), int(height)
            index += segment_length

        raise RuntimeError(f"Cannot parse JPEG dimensions: {file_path.name}")

    raise RuntimeError(f"Unsupported image type: {file_path.suffix}")


class ScriptPushClient:
    def __init__(self, send_url: str, upload_url: str, alias: str = "") -> None:
        self.send_url = send_url
        self.upload_url = upload_url
        self.alias = alias

    @classmethod
    def from_env(cls) -> "ScriptPushClient":
        send_url = os.environ.get("LZ_WEBHOOK_URL", "").strip()
        upload_url = os.environ.get("LZ_WEBHOOK_UPLOAD_URL", "").strip()
        alias = os.environ.get("LZ_WEBHOOK_ALIAS", "").strip()
        if not send_url:
            raise RuntimeError("当前规则未传入机器人 webhook，无法推送")
        if not upload_url:
            upload_url = parse_upload_url(send_url)
        return cls(send_url=send_url, upload_url=upload_url, alias=alias)

    def _request_json(self, payload: dict) -> dict:
        response = requests.post(self.send_url, json=payload, timeout=30)
        response.raise_for_status()
        body = response.json()
        if not body.get("ok", False):
            raise RuntimeError(json.dumps(body, ensure_ascii=False))
        return body

    def send_text(self, content: str) -> None:
        self._request_json({"type": "text", "textMsg": {"content": content}})

    def upload_file(self, file_path: Path, upload_type: int = 2) -> str:
        file_path = Path(file_path).expanduser().resolve()
        upload_url = self.upload_url.replace("type=2", f"type={upload_type}")
        with file_path.open("rb") as file_handle:
            response = requests.post(
                upload_url,
                files={"file": (file_path.name, file_handle)},
                timeout=120,
            )
        response.raise_for_status()
        body = response.json()
        if not body.get("ok", False):
            raise RuntimeError(json.dumps(body, ensure_ascii=False))
        file_id = ((body.get("data") or {}).get("id")) or body.get("id")
        if not file_id:
            raise RuntimeError(f"上传返回缺少 file id: {body}")
        return str(file_id)

    def send_file(self, file_path: Path) -> str:
        file_id = self.upload_file(file_path, upload_type=2)
        self._request_json({"type": "file", "fileMsg": {"fileId": file_id}})
        return file_id

    def send_image(self, file_path: Path) -> str:
        file_path = Path(file_path).expanduser().resolve()
        file_id = self.upload_file(file_path, upload_type=1)
        width, height = get_image_dimensions(file_path)
        self._request_json(
            {
                "type": "image",
                "imageMsg": {"fileId": file_id, "height": int(height), "width": int(width)},
            }
        )
        return file_id
