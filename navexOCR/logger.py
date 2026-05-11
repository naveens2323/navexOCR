import logging
import os
from logging.handlers import RotatingFileHandler

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))
LOG_DIR = os.path.join(BASE_DIR, "logs")
os.makedirs(LOG_DIR, exist_ok=True)

LOG_FILE = os.path.join(LOG_DIR, "navexocr.log")

def setup_logging(level=logging.INFO):
    logger = logging.getLogger()
    logger.setLevel(level)

    fmt = logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s")

    fh = RotatingFileHandler(LOG_FILE, maxBytes=10_000_000, backupCount=5)
    fh.setFormatter(fmt)
    fh.setLevel(level)
    logger.addHandler(fh)

    ch = logging.StreamHandler()
    ch.setFormatter(fmt)
    ch.setLevel(level)
    logger.addHandler(ch)

    return logger


_logger = setup_logging()

def get_logger(name=None):
    return logging.getLogger(name)
