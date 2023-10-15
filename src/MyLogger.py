import pathlib
import logging
import datetime

LogPath = pathlib.Path('./log')

def SetLogger(name):
    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)
    if not LogPath.exists():
        LogPath.mkdir(parents=True, exist_ok=True)

    formatter = logging.Formatter('%(asctime)s:%(levelname)s:%(name)s:%(message)s')
    file_name = LogPath / (datetime.datetime.today().strftime('%Y-%m-%d')+'.log')
    file_handler = logging.FileHandler(file_name, encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)
    logger.handlers.clear()  # 刪除舊handler
    logger.addHandler(file_handler)
    return logger