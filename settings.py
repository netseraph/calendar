"""程序运行设置"""

import logging
import os


def init_logging(level: int):
    """初始化日志系统

    0 NOTSET,

    1 DEBUG,

    2 INFO,

    3 WARNING,

    4 ERROR,

    5 CRITICAL,
    """
    _level_list = (
        logging.NOTSET,
        logging.DEBUG,
        logging.INFO,
        logging.WARNING,
        logging.ERROR,
        logging.CRITICAL,
    )
    if 0 <= level < len(_level_list):
        _level = level
    elif level < 0:
        _level = 0
    else:
        _level = len(_level_list) - 1
    _filename = f"{os.path.split(os.getcwd())[-1]}.log"
    _format = "%(asctime)s - %(filename)s - %(levelname)s - %(message)s"
    logging.basicConfig(
        filename=_filename, format=_format, level=_level, encoding="UTF-8"
    )
