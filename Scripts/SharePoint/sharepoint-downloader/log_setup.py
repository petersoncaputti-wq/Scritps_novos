"""Log historico rotativo e log completo separado por execucao."""
import logging
import uuid
from datetime import datetime, timedelta, timezone
from logging.handlers import RotatingFileHandler


class LocalFormatter(logging.Formatter):
    def formatTime(self, record, datefmt=None):
        return datetime.fromtimestamp(record.created, timezone(timedelta(hours=-3))).isoformat(timespec='milliseconds')


def configure_logging(logs):
    logs.mkdir(exist_ok=True)
    run_id = datetime.now(timezone(timedelta(hours=-3))).strftime('%Y%m%d_%H%M%S') + '_' + uuid.uuid4().hex[:8]
    run_path = logs / f'execucao_{run_id}.log'
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    formatter = LocalFormatter(f'%(asctime)s %(levelname)s execucao={run_id} thread=%(threadName)s %(message)s')
    handlers = [RotatingFileHandler(logs / 'sharepoint_downloader.log', maxBytes=5_000_000, backupCount=3, encoding='utf-8'),
                logging.FileHandler(run_path, encoding='utf-8')]
    for handler in handlers:
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    return run_path, handlers
