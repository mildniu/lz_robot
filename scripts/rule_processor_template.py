#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""邮箱规则处理脚本模板。

调用方式:
    python rule_processor_template.py <attachment_path> <output_dir>

也可以把这个脚本单独打包成 exe 后给主程序调用。

程序会额外注入以下环境变量:
    LZ_ATTACHMENT_PATH
    LZ_OUTPUT_DIR
    LZ_MAIL_SUBJECT
    LZ_MAIL_SENDER
    LZ_MAIL_DATE
    LZ_RULE_KEYWORD
    LZ_MAILBOX_ALIAS
    LZ_WEBHOOK_ALIAS
    LZ_WEBHOOK_URL
    LZ_WEBHOOK_UPLOAD_URL

这个模板的默认行为:
1. 读取输入附件路径
2. 在输出目录中生成一个带时间戳的新文件
3. 如有需要，可使用当前规则选中的机器人直接发送文字/图片/文件

你可以在 `process_attachment()` 里替换成自己的业务逻辑，
例如:
- 读取 Excel 后生成新的报表
- 读取图片后裁剪/压缩/加水印
- 从原始附件中提取部分内容并另存
"""

from __future__ import annotations

import os
import shutil
import sys
from datetime import datetime
from pathlib import Path

from script_push_helper import ScriptPushClient


def build_output_name(input_path: Path) -> str:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    stem = input_path.stem
    suffix = input_path.suffix
    return f"{stem}_processed_{timestamp}{suffix}"


def process_attachment(input_path: Path, output_dir: Path) -> Path:
    """在这里替换为你自己的附件处理逻辑。"""
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / build_output_name(input_path)

    # 默认实现: 直接复制一份，方便先验证整条链路。
    shutil.copy2(input_path, output_path)
    return output_path


def print_context(input_path: Path, output_dir: Path) -> None:
    context = {
        "attachment_path": str(input_path),
        "output_dir": str(output_dir),
        "mail_subject": os.environ.get("LZ_MAIL_SUBJECT", ""),
        "mail_sender": os.environ.get("LZ_MAIL_SENDER", ""),
        "mail_date": os.environ.get("LZ_MAIL_DATE", ""),
        "rule_keyword": os.environ.get("LZ_RULE_KEYWORD", ""),
        "mailbox_alias": os.environ.get("LZ_MAILBOX_ALIAS", ""),
        "webhook_alias": os.environ.get("LZ_WEBHOOK_ALIAS", ""),
    }
    print("rule processor context:")
    for key, value in context.items():
        print(f"  {key}: {value}")


def main() -> int:
    if len(sys.argv) < 3:
        print("usage: python rule_processor_template.py <attachment_path> <output_dir>")
        return 2

    input_path = Path(sys.argv[1]).expanduser().resolve()
    output_dir = Path(sys.argv[2]).expanduser().resolve()

    if not input_path.exists() or not input_path.is_file():
        print(f"input attachment not found: {input_path}")
        return 1

    print_context(input_path, output_dir)
    output_path = process_attachment(input_path, output_dir)
    print(f"generated file: {output_path}")

    # 示例:
    # 如果希望脚本自己推送，可以取消下面注释。
    #
    # client = ScriptPushClient.from_env()
    # client.send_text(f"脚本处理完成: {output_path.name}")
    # if output_path.suffix.lower() in {'.png', '.jpg', '.jpeg'}:
    #     client.send_image(output_path)
    # else:
    #     client.send_file(output_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
