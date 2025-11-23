import logging
import logging.config
import os

from .settings import (
    LOG_BACKUP_COUNT,
    LOG_CONSOLE_LEVEL,
    LOG_DIR,
    LOG_ERROR_FILE_ENABLED,
    LOG_ERROR_FILE_LEVEL,
    LOG_ERROR_FILENAME,
    LOG_INFO_FILE_ENABLED,
    LOG_INFO_FILE_LEVEL,
    LOG_INFO_FILENAME,
    LOG_MAX_BYTES,
    LOG_ROOT_LEVEL,
)


def setup_logging():
    logging_dir = str(LOG_DIR)
    logging_console_level = str(LOG_CONSOLE_LEVEL).upper()
    logging_info_file_level = str(LOG_INFO_FILE_LEVEL).upper()
    logging_info_filename = str(LOG_INFO_FILENAME)
    logging_error_file_level = str(LOG_ERROR_FILE_LEVEL).upper()
    logging_error_filename = str(LOG_ERROR_FILENAME)
    logging_root_level = str(LOG_ROOT_LEVEL).upper()

    if logging_dir:
        os.makedirs(logging_dir, exist_ok=True)

    handlers_config = {}
    root_handlers_list = []

    # Console Handler
    handlers_config['console'] = {
        'level': logging_console_level,
        'class': 'logging.StreamHandler',
        'formatter': 'standard',
    }
    root_handlers_list.append('console')

    # Info File Handler (condicional)
    if LOG_INFO_FILE_ENABLED and logging_dir and logging_info_filename:
        handlers_config['info_file'] = {
            'level': logging_info_file_level,
            'class': 'logging.handlers.RotatingFileHandler',
            'filename': os.path.join(logging_dir, logging_info_filename),
            'maxBytes': LOG_MAX_BYTES,
            'backupCount': LOG_BACKUP_COUNT,
            'mode': 'a',
            'encoding': 'utf-8',
            'formatter': 'standard',
        }
        root_handlers_list.append('info_file')

    # Error File Handler (condicional)
    if LOG_ERROR_FILE_ENABLED and logging_dir and logging_error_filename:
        handlers_config['error_file'] = {
            'level': logging_error_file_level,
            'class': 'logging.handlers.RotatingFileHandler',
            'filename': os.path.join(logging_dir, logging_error_filename),
            'maxBytes': LOG_MAX_BYTES,
            'backupCount': LOG_BACKUP_COUNT,
            'mode': 'a',
            'encoding': 'utf-8',
            'formatter': 'standard',
        }
        root_handlers_list.append('error_file')

    logging_config = {
        'version': 1,
        'disable_existing_loggers': False,
        'formatters': {
            'standard': {'format': '%(asctime)s - %(levelname)s - %(name)s - %(message)s'},
            'detailed': {
                'format': (
                    '%(asctime)s - %(levelname)s - %(name)s - %(module)s - %(funcName)s - %(lineno)d - %(message)s'
                )
            },
        },
        'handlers': handlers_config,
        'root': {'level': logging_root_level, 'handlers': root_handlers_list},
    }

    logging.config.dictConfig(logging_config)
    logging.getLogger(__name__).info(
        f'Logging configurado. Root level: {logging_root_level}. Handlers: {", ".join(root_handlers_list)}'
    )
