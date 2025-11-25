import logging
from pathlib import Path
from typing import Any

from core.config.settings import CONFIG_FILE_PATH, FALLBACK_SERVER_BASE_ADDRESS
from core.utils.utils import load_config_from_ini

logger = logging.getLogger(__name__)


class Config:
    """
    Gere as configurações dinâmicas da aplicação, lidas a partir de um ficheiro .ini.
    Esta classe é projetada para ser instanciada.
    """

    def __init__(self, config_filepath: Path = CONFIG_FILE_PATH):
        """
        Inicializa a configuração.

        Args:
            config_filepath (Path): O caminho para o ficheiro config.ini a ser lido.
                                    Por defeito, usa o caminho padrão da aplicação.
        """
        self._config_path: Path = config_filepath

        # Atributos que serão preenchidos a partir do .ini
        self.SERVER_BASE_ADDRESS: str = FALLBACK_SERVER_BASE_ADDRESS
        self.API_KEY: str = ''
        self.API_SECRET: str = ''
        self.CLIENT_ID: str = ''
        self.PRODUCTION: bool = False
        self.DEBUG: bool = True
        self.DEFAULT_LANGUAGE: str = 'en'
        self.SUPPORTED_LANGUAGES: list = ['en', 'pt_PT']

        # Carrega as configurações na inicialização
        self.reload()

    def reload(self) -> None:
        """
        Lê (ou relê) o ficheiro de configuração do disco e atualiza
        os atributos da instância.
        """
        config_data: dict[str, Any] = {}
        try:
            if self._config_path.exists():
                logger.info(f'Loading/Reloading configuration from: {self._config_path}')
                config_data = load_config_from_ini(self._config_path)
            else:
                logger.warning(f"Config file not found at '{self._config_path}'. Using default/fallback values.")
        except Exception as e:
            logger.exception(f'Failed to load configuration file: {e}')
            # Em caso de falha de leitura, os valores padrão serão mantidos.

        # Environment variables and settings
        self.PRODUCTION = config_data.get('PRODUCTION', False)

        # Debug mode
        self.DEBUG = config_data.get('DEBUG', True)

        # API connection parameters
        self.SERVER_BASE_ADDRESS = str(config_data.get('SERVER_BASE_ADDRESS', self.SERVER_BASE_ADDRESS))
        self.API_KEY = str(config_data.get('API_KEY', ' '))
        self.API_SECRET = str(config_data.get('API_SECRET', ' '))
        self.CLIENT_ID = str(config_data.get('CLIENT_ID', ' '))

        # Internationalization settings
        self.DEFAULT_LANGUAGE = config_data.get('DEFAULT_LANGUAGE', 'en')
        self.SUPPORTED_LANGUAGES = config_data.get('SUPPORTED_LANGUAGES', ['en', 'pt_PT'])
