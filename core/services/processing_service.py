import logging
from pathlib import Path

from core.config.config import Config
from core.config.i18n import _
from core.handler.excel_handler import ExcelHandler
from core.services.api_service import ApiService
from core.services.validation_service import ValidationService
from core.utils.utils import create_config_file

logger = logging.getLogger(__name__)


class ProcessingService:
    """
    Orquestra os fluxos de trabalho principais da aplicação,
    coordenando os diferentes serviços.
    """

    def __init__(self, config: Config) -> None:
        self.config = config
        self.api_service = ApiService(config=self.config)
        self.excel_handler: ExcelHandler | None = None

    def _get_excel_handler(self, handler_type: str = 'sheet', sheet_index: int = 2) -> ExcelHandler:
        """
        Cria uma instância do ExcelHandler sob demanda.
        args:
            handler_type (str): Tipo de handler a criar ('app' ou 'sheet').
            sheet_index (int): Índice da folha para o handler do tipo 'sheet'.
        Returns:
            ExcelHandler: Instância criada do ExcelHandler.
        """
        if handler_type == 'app':
            return ExcelHandler.for_app_only()
        return ExcelHandler.for_sheet(sheet_index=sheet_index)

    def run_auth_process(self, username: str, password: str, config_folder: Path | None = None) -> None:
        """
        Autentica o utilizador, obtém credenciais e cria o ficheiro de configuração.
        args:
            username (str): Nome de utilizador para autenticação.
            password (str): Palavra-passe para autenticação.
            config (Config): Instância de configuração.
            config_folder (Path | None): Pasta onde o ficheiro de configuração será salvo. Se None,
            usa a pasta base da API.
        Returns:
            None
        """
        logger.info(_('Initiating credential generation process...'))

        self.excel_handler = self._get_excel_handler(handler_type='app')

        try:
            credentials = self.api_service.get_api_credentials(username, password)

            if not credentials.get('success'):
                error_msg = credentials.get('error', _('Authentication failed.'))
                logger.error(error_msg)
                self.excel_handler.alert_user(error_msg, _('Authentication Error'))
                return

            if config_folder:
                base_dir = config_folder
            else:
                base_dir = self.api_service.base_dir

            create_config_file(config_path=base_dir, api_url=self.api_service.api_url, credentials=credentials)

            success_msg = _(
                'Credentials generated successfully! Please restart the Excel file for the changes to take effect.'
            )
            self.excel_handler.alert_user(success_msg, _('Success'))

        except Exception as e:
            # Se qualquer passo falhar, regista o erro e tenta alertar o utilizador
            error_msg = _('An error occurred during credential generation: {error}').format(error=e)
            logger.exception(error_msg)
            if self.excel_handler:
                self.excel_handler.alert_user(error_msg, _('Error'))
            # Propaga a exceção para ser apanhada pelo 'main' e sair com código de erro
            raise

    def run_create_process(self, sheet_index: int = 2, excel_handler_override: ExcelHandler | None = None) -> None:
        """
        Executa o fluxo completo de criação de lançamentos.
        Recebe um ExcelHandler já inicializado.
        args:
            sheet_index (int): Índice da folha a processar.
            excel_handler_override (ExcelHandler | None): Instância do ExcelHandler para sobrescrever a criação padrão.
        Returns:
            None
        """

        logger.info(_('Starting creation process'))

        self.excel_handler = excel_handler_override or self._get_excel_handler(
            handler_type='sheet', sheet_index=sheet_index
        )

        max_lines = self.api_service.get_max_journal_lines()

        # Ler os dados
        raw_df = self.excel_handler.read_data_to_create()
        if raw_df.empty:
            logger.warning(_('The data table is empty. Process interrupted.'))
            self.excel_handler.alert_user(_('The data table is empty.'), _('Warning'))
            return

        # Processar os dados
        validator = ValidationService(raw_df)
        data_groups = validator.group_data(max_lines=max_lines)

        # Enviar para a API
        all_results = []

        # Feedback visual (barra de status do Excel)
        if self.excel_handler.app:
            self.excel_handler.app.screen_updating = False  # Opcional: Melhora performance

        for group_df in data_groups:
            api_response = self.api_service.create_journal_entry(group_df)
            result = {'indices': group_df.index, 'response': api_response}
            all_results.append(result)

        if self.excel_handler.app:
            self.excel_handler.app.screen_updating = True  # Reativa a atualização após o processamento

        # Escrever os resultados
        self.excel_handler.write_results_to_sheet(all_results)

        success_message = _('Process completed! {num_groups} groups sent.').format(num_groups=len(data_groups))
        logger.info(success_message)
        self.excel_handler.alert_user(success_message, _('Success'))

    def run_status_check_process(
        self, sheet_index: int = 2, excel_handler_override: ExcelHandler | None = None
    ) -> None:
        """Executa o fluxo de verificação de status."""
        logger.info(_('Starting status check process'))

        self.excel_handler = excel_handler_override or self._get_excel_handler(
            handler_type='sheet', sheet_index=sheet_index
        )

        # Ler os dados
        raw_df = self.excel_handler.read_data_to_update()
        if raw_df.empty:
            logger.warning(_('The data table is empty. Process interrupted.'))
            self.excel_handler.alert_user(_('The data table is empty.'), _('Warning'))
            return

        # Enviar para a API
        success_count = 0
        total_count = len(raw_df)

        logger.info(_('Found {total_count} documents to check.').format(total_count=total_count))

        for _i, row in raw_df.iterrows():
            doc_number = str(row['Document'])
            row_idx = int(row['original_row_index'])

            # Chama a API
            status_result = self.api_service.get_journal_status(doc_number)

            # Verifica se houve erro de comunicação
            if status_result.get('success', True) is False:
                error_msg = status_result.get('error', 'Unknown Error')
                self.excel_handler.update_row_status(row_idx, new_status=None, message=error_msg)
            else:
                # Tenta obter o status específico deste documento
                # A API retorna algo como { "DOC123": "Posted" }
                current_status = status_result.get(doc_number)

                if current_status:
                    self.excel_handler.update_row_status(row_idx, new_status=current_status, message='')
                    success_count += 1
                else:
                    self.excel_handler.update_row_status(
                        row_idx, new_status='Not Found', message='Document ID not found'
                    )

        msg = _('Status update complete. {success}/{total} updated.').format(success=success_count, total=total_count)
        logger.info(msg)
        self.excel_handler.alert_user(msg, _('Process Complete'))

    def alert_user_critical_error(self, message: str):
        """
        Alerta o utilizador sobre um erro crítico.
        args:
            message (str): Mensagem de erro a ser exibida.
        """
        if not self.excel_handler:
            self.excel_handler = self._get_excel_handler('app')
        self.excel_handler.alert_user(message, _('Critical Error'))
