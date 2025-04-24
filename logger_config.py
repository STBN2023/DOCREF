import logging
import sys

def get_logger(name=__name__):
    logger = logging.getLogger(name)
    if not logger.handlers:
        logger.setLevel(logging.INFO)
        fmt = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        handlers = [
            logging.StreamHandler(sys.stdout),
            logging.FileHandler('powerpoint_generator.log', encoding='utf-8')
        ]
        for h in handlers:
            h.setFormatter(logging.Formatter(fmt))
            logger.addHandler(h)
    return logger
