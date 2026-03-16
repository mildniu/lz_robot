import time
from logging import Logger
from typing import Optional

from .config import AppConfig
from .logging_utils import get_logger
from .processing_service import MailProcessingService, build_webhook_client


class MailForwarderWorker:
    def __init__(self, config: AppConfig, logger: Optional[Logger] = None) -> None:
        self.config = config
        self.logger = logger or get_logger()
        self.service = MailProcessingService(config)
        self.webhook = build_webhook_client(config.webhook_send_url)

    def run_once(self) -> None:
        batch_result = self.service.process_rule_batch(update_state=True)
        if not batch_result.results:
            self.logger.info("No enabled mail rules found")
            return
        for item in batch_result.results:
            mailbox_text = f"[{item.mailbox_alias}] " if item.mailbox_alias else ""
            if item.status == "processed":
                self.logger.info(
                    "%sRule %s finished uid=%s total files=%d",
                    mailbox_text,
                    item.rule_keyword,
                    item.uid,
                    len(item.files),
                )
            else:
                self.logger.info("%sRule %s: %s", mailbox_text, item.rule_keyword, item.reason)

    def run_forever(self) -> None:
        self.logger.info(
            "Worker started, polling every %s seconds", self.config.poll_interval_seconds
        )
        while True:
            try:
                self.run_once()
            except Exception as exc:
                self.logger.exception("Run failed: %s", exc)
                try:
                    self.webhook.send_text_alert(f"【邮件推送失败】{exc}")
                except Exception as alert_exc:
                    self.logger.error("Failed to send alert via webhook: %s", alert_exc)
            time.sleep(self.config.poll_interval_seconds)
