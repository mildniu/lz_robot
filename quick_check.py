#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
快速诊断：检查文件夹监测配置
"""

import sys
import json
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

def quick_check():
    """快速检查配置"""
    print("=" * 70)
    print("Quick Configuration Check")
    print("=" * 70)

    # 读取配置
    config_file = Path("settings/folder_monitor_config.json")

    if not config_file.exists():
        print("\n[ERROR] Config file not found!")
        print("\nPlease:")
        print("1. Start GUI: python gui_app_v3.py")
        print("2. Go to Settings page")
        print("3. Configure folder monitor")
        print("4. Save config")
        return

    with open(config_file, 'r', encoding='utf-8') as f:
        config = json.load(f)

    print("\n[1] Checking folder monitors...")
    for i in range(1, 4):
        key = f"folder_{i}"
        if key in config:
            cfg = config[key]
            path = cfg.get("path", "")
            webhook = cfg.get("webhook_url", "")
            enabled = cfg.get("enabled", False)

            status = "[ENABLED]" if enabled else "[DISABLED]"
            print(f"\n{key} {status}")
            print(f"  Path: {path if path else '(empty)'}")

            # 显示webhook的前50个字符
            if webhook:
                display_webhook = webhook[:50] + "..." if len(webhook) > 50 else webhook
                print(f"  Webhook: {display_webhook}")
                print(f"  Full URL length: {len(webhook)} characters")
            else:
                print(f"  Webhook: (empty)")

    # 检查启用的监测
    print("\n[2] Checking enabled monitors...")
    enabled_monitors = []
    for key, cfg in config.items():
        if cfg.get("enabled", False):
            enabled_monitors.append(key)

    if not enabled_monitors:
        print("[WARNING] No monitors are enabled!")
        print("\nTo enable:")
        print("1. Open GUI: python gui_app_v3.py")
        print("2. Go to Settings")
        print("3. Check the enable checkbox")
        print("4. Save config")
    else:
        print(f"[OK] {len(enabled_monitors)} monitor(s) enabled:")
        for key in enabled_monitors:
            cfg = config[key]
            webhook = cfg.get("webhook_url", "")
            if webhook:
                print(f"  - {key}: Webhook configured")
            else:
                print(f"  - {key}: [WARNING] Webhook is empty!")

    print("\n" + "=" * 70)
    print("Quick Check Complete")
    print("=" * 70)
    print("\nIf Webhook URLs are configured:")
    print("  -> The GUI should display them in Settings")
    print("  -> Click in the input box to see full URL")
    print("  -> Use arrow keys to scroll through long URLs")
    print("\nTo test folder monitoring:")
    print("  1. Start GUI: python gui_app_v3.py")
    print("  2. Go to Monitor page")
    print("  3. Click Start button")
    print()

if __name__ == "__main__":
    try:
        quick_check()
    except Exception as e:
        print(f"\n[ERROR] {e}")
        import traceback
        traceback.print_exc()
