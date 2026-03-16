from email.header import decode_header
from typing import Optional


def decode_mime_text(value: Optional[str]) -> str:
    if not value:
        return ""
    chunks = []
    for part, charset in decode_header(value):
        if isinstance(part, bytes):
            chunks.append(part.decode(charset or "utf-8", errors="replace"))
        else:
            chunks.append(part)
    return "".join(chunks).strip()
