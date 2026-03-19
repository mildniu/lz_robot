import email
import imaplib
import socket
import time
import base64
from datetime import datetime
from email.message import Message
from typing import Optional, Tuple

from .mime_utils import decode_mime_text


def _encode_imap_utf7(value: str) -> bytes:
    """Encode mailbox names using IMAP modified UTF-7 for non-ASCII folders."""
    if not value:
        return b"INBOX"

    chunks: list[bytes] = []
    buffer: list[str] = []

    def flush_buffer() -> None:
        if not buffer:
            return
        encoded = "".join(buffer).encode("utf-16-be")
        modified = base64.b64encode(encoded).rstrip(b"=").replace(b"/", b",")
        chunks.append(b"&" + modified + b"-")
        buffer.clear()

    for ch in value:
        code = ord(ch)
        if 0x20 <= code <= 0x7E:
            flush_buffer()
            if ch == "&":
                chunks.append(b"&-")
            else:
                chunks.append(ch.encode("ascii"))
        else:
            buffer.append(ch)

    flush_buffer()
    return b"".join(chunks)


class ImapMailClient:
    def __init__(
        self,
        host: str,
        port: int,
        username: str,
        password: str,
        mailbox: str,
        timeout_seconds: int = 30,
    ) -> None:
        self.host = host
        self.port = port
        self.username = username
        self.password = password
        self.mailbox = mailbox
        self.timeout_seconds = timeout_seconds
        self._imap: Optional[imaplib.IMAP4_SSL] = None

    def __enter__(self) -> "ImapMailClient":
        # Prevent network hangs from blocking the worker forever.
        self._imap = imaplib.IMAP4_SSL(self.host, self.port, timeout=self.timeout_seconds)
        # The Qt settings page allows Chinese mailbox folder names, and some accounts may
        # also contain non-ASCII credentials. Switch the client to UTF-8 command encoding
        # before LOGIN/SELECT so imaplib does not fail early with ASCII encoding errors.
        if any(any(ord(ch) > 127 for ch in str(value)) for value in (self.username, self.password, self.mailbox)):
            try:
                self._imap._mode_utf8()
            except Exception:
                pass
        self._imap.login(self.username, self.password)
        status, _ = self._imap.select(_encode_imap_utf7(self.mailbox), readonly=True)
        if status != "OK":
            raise RuntimeError(f"Failed to select mailbox: {self.mailbox}")
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        if self._imap is not None:
            try:
                self._imap.logout()
            except (imaplib.IMAP4.error, OSError, socket.timeout):
                pass
            self._imap = None

    @property
    def imap(self) -> imaplib.IMAP4_SSL:
        if self._imap is None:
            raise RuntimeError("IMAP connection is not established")
        return self._imap

    def _fetch_header_and_internaldate(
        self, uid: bytes
    ) -> Tuple[Optional[str], Optional[datetime]]:
        status, data = self.imap.uid(
            "fetch", uid, "(BODY.PEEK[HEADER.FIELDS (SUBJECT)] INTERNALDATE)"
        )
        if status != "OK" or not data:
            return None, None

        subject_value = None
        internal_dt = None
        for item in data:
            if not isinstance(item, tuple):
                continue
            meta = item[0]
            raw_header = item[1]
            if raw_header:
                header_msg = email.message_from_bytes(raw_header)
                subject_value = decode_mime_text(header_msg.get("Subject"))
            if meta:
                internal_tuple = imaplib.Internaldate2tuple(meta)
                if internal_tuple:
                    internal_dt = datetime.fromtimestamp(time.mktime(internal_tuple))
        return subject_value, internal_dt

    def find_latest_uid(self, subject_keywords: list[str]) -> Optional[bytes]:
        """查找包含任一关键字的最新邮件UID"""
        return self.find_latest_uid_by_subject(subject_keywords, match_mode="contains")

    def find_latest_uid_by_subject(
        self,
        subject_keywords: list[str],
        *,
        match_mode: str = "contains",
    ) -> Optional[bytes]:
        """按收件时间查找主题匹配的最新邮件 UID"""
        status, data = self.imap.uid("search", None, "ALL")
        if status != "OK" or not data or not data[0]:
            return None

        latest_uid = None
        latest_dt = None
        for uid in data[0].split():
            subject, internal_dt = self._fetch_header_and_internaldate(uid)
            if not subject:
                continue

            # 检查主题是否包含任一关键字
            matched = False
            if subject_keywords:
                for keyword in subject_keywords:
                    if match_mode == "exact":
                        if subject == keyword:
                            matched = True
                            break
                    elif keyword in subject:
                        matched = True
                        break
            else:
                # 如果没有关键字，匹配所有邮件
                matched = True

            if not matched:
                continue

            dt = internal_dt or datetime.min
            if latest_dt is None or dt > latest_dt:
                latest_uid = uid
                latest_dt = dt
                continue

            if dt == latest_dt and latest_uid is not None:
                try:
                    if int(uid) > int(latest_uid):
                        latest_uid = uid
                except ValueError:
                    pass
        return latest_uid

    def fetch_message(self, uid: bytes) -> Message:
        status, data = self.imap.uid("fetch", uid, "(RFC822)")
        if status != "OK" or not data:
            raise RuntimeError(f"Failed to fetch message uid={uid.decode()}")

        for item in data:
            if isinstance(item, tuple) and item[1]:
                return email.message_from_bytes(item[1])
        raise RuntimeError(f"Message body is empty for uid={uid.decode()}")
