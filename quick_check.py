#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""量子推送机器人快速自检"""

from __future__ import annotations

import json
import sys
from pathlib import Path


def load_json(path: Path) -> dict:
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def print_section(title: str):
    print("\n" + "=" * 72)
    print(title)
    print("=" * 72)


def check_required_files():
    print_section("1. 核心文件检查")
    required_files = [
        Path("gui_app.py"),
        Path("start_gui.bat"),
        Path("README.md"),
        Path("QuantumBot.spec"),
        Path("settings/app_config.json"),
    ]
    for file_path in required_files:
        status = "OK" if file_path.exists() else "MISSING"
        print(f"[{status}] {file_path}")


def check_mail_rules():
    print_section("2. 邮件检测规则")
    rules_payload = load_json(Path("settings/subject_attachment_rules.json"))
    rules = list(rules_payload.get("rules", []))
    enabled_rules = [item for item in rules if item.get("enabled")]
    print(f"规则总数: {len(rules)}")
    print(f"启用规则: {len(enabled_rules)}")
    if not enabled_rules:
        print("[WARNING] 当前没有启用的邮件规则")
        return

    for index, rule in enumerate(rules, start=1):
        if not isinstance(rule, dict) or not rule:
            print(f"- 槽位{index}: (空)")
            continue
        keyword = str(rule.get("keyword", "")).strip() or "(未填写主题)"
        mailbox_alias = str(rule.get("mailbox_alias", "")).strip() or "(未选邮箱)"
        webhook_alias = str(rule.get("webhook_alias", "")).strip() or "(未选机器人)"
        interval_seconds = int(rule.get("poll_interval_seconds", 0) or 0)
        trigger_mode = str(rule.get("trigger_mode", "periodic")).strip() or "periodic"
        trigger_text = (
            f"定时={str(rule.get('schedule_time', '')).strip() or '--:--'}"
            if trigger_mode == "timed"
            else f"周期={interval_seconds}s"
        )
        mode = "脚本处理" if str(rule.get("script_path", "")).strip() else "直接推送"
        status = "启用" if rule.get("enabled") else "未启用"
        print(
            f"- 槽位{index}: {status} | {keyword} | 邮箱={mailbox_alias} | 机器人={webhook_alias} | "
            f"{trigger_text} | 模式={mode}"
        )


def check_folder_monitors():
    print_section("3. 文件夹检测")
    config = load_json(Path("settings/folder_monitor_config.json"))
    enabled_count = 0
    for i in range(1, 4):
        key = f"folder_{i}"
        cfg = config.get(key, {})
        enabled = bool(cfg.get("enabled", False))
        if enabled:
            enabled_count += 1
        path = str(cfg.get("path", "")).strip()
        alias = str(cfg.get("webhook_alias", "")).strip()
        print(
            f"- {key}: {'启用' if enabled else '未启用'} | "
            f"路径={path or '(空)'} | 机器人={alias or '(空)'}"
        )
    if enabled_count == 0:
        print("[WARNING] 当前没有启用的文件夹检测项")


def check_runtime_state():
    print_section("4. 运行状态文件")
    for file_path in [Path("state/mail_state.json"), Path("state/file_sent_state.json")]:
        if not file_path.exists():
            print(f"[INFO] {file_path} 不存在")
            continue
        try:
            payload = json.loads(file_path.read_text(encoding="utf-8"))
            print(f"[OK] {file_path} | 记录数={len(payload)}")
        except Exception as exc:
            print(f"[ERROR] {file_path} 读取失败: {exc}")


def check_packaging_assets():
    print_section("5. 打包资源检查")
    spec_files = [Path("QuantumBot.spec"), Path("QuantumBot.mac.spec")]
    forbidden_text = "GUI_V3_README.md"
    for spec_path in spec_files:
        if not spec_path.exists():
            print(f"[MISSING] {spec_path}")
            continue
        content = spec_path.read_text(encoding="utf-8")
        if forbidden_text in content:
            print(f"[WARNING] {spec_path} 仍包含旧文件名 {forbidden_text}")
        else:
            print(f"[OK] {spec_path} 未发现旧文件名引用")


def check_stability_settings():
    print_section("6. 稳定性配置")
    app_config = load_json(Path("settings/app_config.json"))
    timeout_seconds = str(app_config.get("SCRIPT_TIMEOUT_SECONDS", "")).strip() or "300"
    print(f"- 处理程序超时: {timeout_seconds}s")
    try:
        parsed = int(timeout_seconds)
        if parsed < 30:
            print("[WARNING] SCRIPT_TIMEOUT_SECONDS 过小，复杂脚本可能被提前中断")
        else:
            print("[OK] 处理程序超时配置有效")
    except ValueError:
        print("[ERROR] SCRIPT_TIMEOUT_SECONDS 不是整数")


def main():
    print("=" * 72)
    print("量子推送机器人 Quick Check")
    print("=" * 72)
    check_required_files()
    check_mail_rules()
    check_folder_monitors()
    check_runtime_state()
    check_packaging_assets()
    check_stability_settings()
    print("\n检查完成。")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"\n[ERROR] {exc}")
        raise
