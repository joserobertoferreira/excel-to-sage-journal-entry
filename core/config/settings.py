import os
import sys
from datetime import date, datetime
from pathlib import Path

from core.utils.utils import load_config_from_env

# from decouple import config

# Build paths inside the project like this: BASE_DIR / 'subdir'.
if getattr(sys, 'frozen', False):
    # Se sim, o diretório base é a pasta onde o .exe está.
    BASE_DIR = Path(os.path.dirname(sys.executable))
else:
    BASE_DIR = Path(__file__).resolve().parent.parent

dotenv_path = BASE_DIR / '.env'

CONFIG = load_config_from_env(dotenv_path)

# eNvironment variables and settings
PRODUCTION = CONFIG.get('PRODUCTION', False)

# Debug mode
DEBUG = CONFIG.get('DEBUG', True)

# API connection parameters
SERVER_BASE_ADDRESS = str(CONFIG.get('SERVER_BASE_ADDRESS', None))
API_KEY = str(CONFIG.get('API_KEY', ' '))
API_SECRET = str(CONFIG.get('API_SECRET', ' '))
CLIENT_ID = str(CONFIG.get('CLIENT_ID', ' '))

# Logging configuration
LOG_DIR = BASE_DIR / 'logs'
LOG_ROOT_LEVEL = 'DEBUG'
LOG_CONSOLE_LEVEL = 'INFO'
LOG_INFO_FILE_ENABLED = True
LOG_INFO_FILENAME = 'app_info.log'
LOG_INFO_FILE_LEVEL = 'INFO'
LOG_ERROR_FILE_ENABLED = True
LOG_ERROR_FILENAME = 'app_error.log'
LOG_ERROR_FILE_LEVEL = 'ERROR'
LOG_MAX_BYTES = 10 * 1024 * 1024  # 10 MB
LOG_BACKUP_COUNT = 5

# Sage X3 database table settings
DEFAULT_LEGACY_DATE = date(1753, 1, 1)
DEFAULT_LEGACY_DATETIME = datetime(1753, 1, 1)

# Excel spreadsheet settings
START_CELL = 'A'  # The starting cell of the data range
END_CELL = 'AD'  # The ending cell of the data range
START_FEEDBACK_CELL = 'B'  # The starting cell of the feedback range
END_FEEDBACK_CELL = 'D'  # The ending cell of the feedback range

# Columns used to fill down values and group the data
PRIMARY_GROUP_COLUMN = 'Group By'
SECONDARY_GROUP_COLUMNS = ['Site', 'Entry Type', 'AccountingDate', 'Curr']
GROUPING_COLUMNS = [PRIMARY_GROUP_COLUMN] + SECONDARY_GROUP_COLUMNS

# Exact names of all expected columns in the spreadsheet
EXPECTED_COLUMNS = [
    'Group By',
    'Document',
    'Status',
    'Warning',
    'Site',
    'Entry Type',
    'AccountingDate',
    'VAT date',
    'Reversing Y/N (1=No 2=Yes)',
    'Reversing Date',
    'Header Description',
    'Source',
    'Curr',
    'Reference',
    'Nominal Code',
    'Line Description',
    'Collective',
    'BP',
    'Tax',
    'FIX',
    'BRK',
    'DEP',
    'LOC',
    'TYP',
    'PDT',
    'ANA',
    'Quantity',
    'Debit',
    'Credit',
    'Free Reference',
]

# API settings
DIMENSIONS_MAPPING = {
    'fixture': 'FIX',
    'broker': 'BRK',
    'department': 'DEP',
    'location': 'LOC',
    'type': 'TYP',
    'product': 'PDT',
    'analysis': 'ANA',
}

# Internationalization settings
LOCALE_DIR = os.path.join(BASE_DIR, 'locales')
DEFAULT_LANGUAGE = CONFIG.get('DEFAULT_LANGUAGE', 'en')
SUPPORTED_LANGUAGES = CONFIG.get('SUPPORTED_LANGUAGES', ['en', 'pt_PT'])
