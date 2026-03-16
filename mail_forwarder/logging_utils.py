import logging


def get_logger() -> logging.Logger:
    logger = logging.getLogger("imap_webhook_forwarder")
    logger.setLevel(logging.INFO)
    if logger.handlers:
        return logger

    handler = logging.StreamHandler()
    handler.setFormatter(
        logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S")
    )
    logger.addHandler(handler)
    return logger
