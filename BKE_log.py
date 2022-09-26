import logging
from logging.handlers import RotatingFileHandler
def setup_custom_logger(name):
    filename = "log.txt"

    logging.basicConfig(handlers=[RotatingFileHandler(filename, maxBytes=10485760, backupCount=5)],
                        level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s ')
    logger = logging.getLogger()
    return logger