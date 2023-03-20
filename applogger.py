import logging

from appresource import APP_LOG


class AppLogger(object):

    def __init__(self, filename=APP_LOG):
        self.filename = filename
        self.format = '[%(asctime)s-%(filename)s-%(levelname)s:%(message)s]'
        self.level = logging.DEBUG
        self.datefmt = '%Y-%m-%d%I:%M:%S %p'
        logging.basicConfig(filename=filename, format=self.format, level=self.level, datefmt=self.datefmt,
                            filemode='a', encoding='utf-8')

    @staticmethod
    def error(msg):
        logging.error(msg)

    @staticmethod
    def warning(msg):
        logging.warning(msg)

    @staticmethod
    def info(msg):
        logging.info(msg)

    @staticmethod
    def debug(msg):
        logging.debug(msg)

