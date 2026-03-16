from email.message import Message
from pathlib import Path
from typing import List

from .mime_utils import decode_mime_text


MIME_EXTENSION_MAP = {
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "xlsx",
    "application/vnd.ms-excel": "xls",
    "image/png": "png",
    "image/jpeg": "jpeg",
    "image/jpg": "jpg",
}


def normalize_attachment_types(types: list[str]) -> list[str]:
    normalized = []
    for item in types:
        value = item.strip().lower().lstrip(".")
        if not value:
            continue
        normalized.append(value)

    # Keep compatibility between jpeg/jpg.
    values = set(normalized)
    if "jpeg" in values:
        values.add("jpg")
    if "jpg" in values:
        values.add("jpeg")
    return sorted(values)


def extract_attachments_by_types(
    message: Message,
    output_dir: Path,
    uid: bytes,
    allowed_types: list[str],
    filename_keywords: list[str] | None = None,
) -> List[Path]:
    saved_files: List[Path] = []
    output_dir.mkdir(parents=True, exist_ok=True)
    attachment_index = 1
    allowed_exts = set(normalize_attachment_types(allowed_types))
    normalized_name_keywords = [kw.strip().lower() for kw in (filename_keywords or []) if kw.strip()]
    if not allowed_exts:
        return saved_files

    for part in message.walk():
        if part.is_multipart():
            continue

        filename = decode_mime_text(part.get_filename())
        disposition = (part.get("Content-Disposition") or "").lower()
        content_type = (part.get_content_type() or "").lower()

        has_filename = bool(filename)
        is_attachment_like = "attachment" in disposition or has_filename
        name_ext = ""
        if filename and "." in filename:
            name_ext = filename.rsplit(".", 1)[-1].lower().strip()
        type_ext = MIME_EXTENSION_MAP.get(content_type, "").lower().strip()

        matches_by_name = bool(name_ext and name_ext in allowed_exts)
        matches_by_type = bool(type_ext and type_ext in allowed_exts)

        if not is_attachment_like:
            continue
        if not matches_by_name and not matches_by_type:
            continue
        if normalized_name_keywords:
            if not filename:
                continue
            filename_lower = filename.lower()
            if not any(keyword in filename_lower for keyword in normalized_name_keywords):
                continue

        payload = part.get_payload(decode=True)
        if not payload:
            continue

        # Prefer filename extension; fallback to MIME extension.
        final_ext = name_ext if matches_by_name else type_ext
        if not final_ext:
            final_ext = next(iter(allowed_exts))

        if not filename:
            filename = f"attachment_{uid.decode()}_{attachment_index}.{final_ext}"
        elif "." not in filename:
            filename = f"{filename}_{uid.decode()}.{final_ext}"
        elif name_ext not in allowed_exts and matches_by_type:
            filename = f"{filename}_{uid.decode()}.{final_ext}"
        else:
            stem, dot, suffix = filename.rpartition(".")
            if stem and dot and suffix:
                filename = f"{stem}_{uid.decode()}.{suffix}"
            else:
                filename = f"{filename}_{uid.decode()}"

        safe_name = filename.replace("\\", "_").replace("/", "_")
        save_path = output_dir / safe_name
        save_path.write_bytes(payload)
        saved_files.append(save_path)
        attachment_index += 1

    return saved_files


def extract_xlsx_attachments(message: Message, output_dir: Path, uid: bytes) -> List[Path]:
    return extract_attachments_by_types(message, output_dir, uid, ["xlsx"])
