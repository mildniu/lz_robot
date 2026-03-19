from __future__ import annotations

from dataclasses import dataclass, field

from .log_bus import LogBus


@dataclass
class MailRuleRuntimeManager:
    log_bus: LogBus
    running_rule_ids: set[str] = field(default_factory=set)

    def start_rule(self, rule_id: str) -> None:
        self.running_rule_ids.add(rule_id)
        self.log_bus.emit("INFO", f"[Qt Preview] 启动邮件规则: {rule_id}", source=rule_id)

    def stop_rule(self, rule_id: str) -> None:
        self.running_rule_ids.discard(rule_id)
        self.log_bus.emit("INFO", f"[Qt Preview] 停止邮件规则: {rule_id}", source=rule_id)


@dataclass
class FolderMonitorRuntimeManager:
    log_bus: LogBus
    running_monitor_ids: set[str] = field(default_factory=set)

    def start_monitor(self, monitor_id: str) -> None:
        self.running_monitor_ids.add(monitor_id)
        self.log_bus.emit("INFO", f"[Qt Preview] 启动文件夹监测: {monitor_id}", source=monitor_id)

    def stop_monitor(self, monitor_id: str) -> None:
        self.running_monitor_ids.discard(monitor_id)
        self.log_bus.emit("INFO", f"[Qt Preview] 停止文件夹监测: {monitor_id}", source=monitor_id)
