import logging
import os
import sys
from datetime import date, datetime
from pathlib import Path

logger = logging.getLogger(__name__)

# FALLBACK_SERVER_BASE_ADDRESS = 'http://cfg-uks-x3-03:8241/graphql'
FALLBACK_SERVER_BASE_ADDRESS: str = 'http://localhost:3000/graphql'

# Build paths inside the project like this: BASE_DIR / 'subdir'.
IS_FROZEN: bool = getattr(sys, 'frozen', False)

if IS_FROZEN:
    BASE_DIR: Path = Path(os.path.dirname(sys.executable))
else:
    BASE_DIR: Path = Path(__file__).resolve().parent.parent

CONFIG_FILE_PATH: Path = BASE_DIR / 'config.ini'

# Excel API credentials
EXCEL_API_KEY: str = 'vw&_nbb08*1o--uxu9xs-pc3&ng!$i@ynh=^3o%s=3@sc^g$kt'

# Logging configuration
LOG_DIR: Path = BASE_DIR / 'logs'
LOG_ROOT_LEVEL: str = 'DEBUG'
LOG_CONSOLE_LEVEL: str = 'INFO'
LOG_INFO_FILE_ENABLED: bool = True
LOG_INFO_FILENAME: str = 'app_info.log'
LOG_INFO_FILE_LEVEL: str = 'INFO'
LOG_ERROR_FILE_ENABLED: bool = True
LOG_ERROR_FILENAME: str = 'app_error.log'
LOG_ERROR_FILE_LEVEL: str = 'ERROR'
LOG_MAX_BYTES: int = 10 * 1024 * 1024  # 10 MB
LOG_BACKUP_COUNT: int = 5

# Sage X3 database table settings
DEFAULT_LEGACY_DATE: date = date(1753, 1, 1)
DEFAULT_LEGACY_DATETIME: datetime = datetime(1753, 1, 1)

# Excel spreadsheet settings
START_CELL: str = 'A'  # The starting cell of the data range
END_CELL: str = 'AD'  # The ending cell of the data range

# Feedback range settings
START_FEEDBACK_CELL: str = 'B'  # The starting cell of the feedback range
END_FEEDBACK_CELL: str = 'D'  # The ending cell of the feedback range

FEEDBACK_HEADER_COLUMNS: dict[str, str] = {
    'Document': 'B',
    'Status': 'C',
    'Warning': 'D',
    'Site': 'E',
    'Entry Type': 'F',
    'AccountingDate': 'G',
    'VAT date': 'H',
    'Curr': 'L',
}

FEEDBACK_DETAIL_COLUMNS: dict[str, str] = {
    'businessPartner': 'Q',
    'tax': 'R',
}

FEEDBACK_DIMENSION_COLUMNS: dict[str, str] = {
    'fixture': 'S',
    'broker': 'T',
    'department': 'U',
    'location': 'V',
    'type': 'W',
    'product': 'X',
    'analysis': 'Y',
}

# Columns used to fill down values and group the data
PRIMARY_GROUP_COLUMN: str = 'Group By'
SECONDARY_GROUP_COLUMNS: list[str] = ['Site', 'Entry Type', 'AccountingDate', 'Header Description', 'Curr']
GROUPING_COLUMNS: list[str] = [PRIMARY_GROUP_COLUMN] + SECONDARY_GROUP_COLUMNS

# Exact names of all expected columns in the spreadsheet
EXPECTED_COLUMNS: list[str] = [
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
    'Debit',
    'Credit',
    'Quantity',
    'Free Reference',
    '_isLocked',
]

COLUMNS_TO_UPPERCASE: list[str] = [
    'Site',
    'Entry Type',
    'Curr',
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
]

DATE_COLUMNS: list[str] = [
    'AccountingDate',
    'VAT date',
    'Reversing Date',
]

# API settings
DIMENSIONS_MAPPING: dict[str, str] = {
    'fixture': 'FIX',
    'broker': 'BRK',
    'department': 'DEP',
    'location': 'LOC',
    'type': 'TYP',
    'product': 'PDT',
    'analysis': 'ANA',
}

# Internationalization settings
LOCALE_DIR: str = os.path.join(BASE_DIR, 'locales')
