import logging
import os

log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
logger = logging.getLogger(__name__)
logger.setLevel('DEBUG')
file_handler = logging.FileHandler('Logs.log')
formatter = logging.Formatter(log_format)
file_handler.setFormatter(formatter)

logger.addHandler(file_handler)


def open_outlook(outlook_loc: str):
    try:
        logger.info('{}'.format(f"Attempting to open Outlook Application"))
        os.startfile(outlook_loc)
        logger.info('{}'.format(f"Successfully opened Outlook Application"))
        os.close
    except FileNotFoundError as e:
        logger.error('{}'.format(f"ERROR: {e.errno} - {e}"))
    except PermissionError as e:
        logger.error('{}'.format(f"ERROR: {e.errno} - {e}"))


def close_outlook():
    os.system('taskkill /F /IM outlook.exe')