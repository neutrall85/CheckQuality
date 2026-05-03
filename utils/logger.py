import logging
import os
from logging.handlers import RotatingFileHandler

LOG_DIR = os.path.join(os.path.expanduser("~"), "AppData", "Roaming", "ExcelReporter", "logs")
os.makedirs(LOG_DIR, exist_ok=True)

def setup_logger(name: str = __name__) -> logging.Logger:
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)

    if not logger.handlers:
        # Файловый обработчик (общий лог)
        file_handler = RotatingFileHandler(
            os.path.join(LOG_DIR, "app.log"),
            maxBytes=10 * 1024 * 1024,  # 10 MB
            backupCount=7,
            encoding='utf-8'
        )
        file_handler.setLevel(logging.DEBUG)
        formatter = logging.Formatter('[%(asctime)s] [%(levelname)s] [%(name)s] %(message)s',
                                      datefmt='%Y-%m-%d %H:%M:%S')
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

        # Обработчик для критических ошибок
        critical_handler = RotatingFileHandler(
            os.path.join(LOG_DIR, "critical.log"),
            maxBytes=5 * 1024 * 1024,
            backupCount=3,
            encoding='utf-8'
        )
        critical_handler.setLevel(logging.ERROR)
        critical_handler.setFormatter(formatter)
        logger.addHandler(critical_handler)

    return logger