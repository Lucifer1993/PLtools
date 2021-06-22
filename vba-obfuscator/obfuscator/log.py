import logging
import sys


def configure_logging():
    # create a stdout handler
    handler = logging.StreamHandler(sys.stderr)
    handler.setLevel(logging.INFO)

    # create a logging format
    formatter = logging.Formatter('[%(levelname)s] %(message)s')
    handler.setFormatter(formatter)

    logging.basicConfig(
        handlers=[handler],
        level=logging.DEBUG,
    )
