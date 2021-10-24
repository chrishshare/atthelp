# -*- coding: UTF8 -*-
import json
import logging.config
import os


def init_logging(cf='logger.json', level='logging.DEBUG'):
    if os.path.exists(cf):
        with open(cf, 'r', encoding='utf-8') as f:
            config = json.load(f)
            logging.config.dictConfig(config)
    logger = logging.getLogger(__name__)
    return logger

