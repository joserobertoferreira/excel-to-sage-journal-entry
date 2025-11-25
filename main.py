import logging
import sys
from pathlib import Path
from typing import NoReturn

from core.config.config import Config
from core.config.i18n import _
from core.config.logging import setup_logging
from core.config.settings import BASE_DIR
from core.handler.excel_handler import ExcelHandler
from core.services.api_service import ApiService
from core.services.processing_service import ProcessingService
from core.utils.utils import create_config_file

setup_logging()
logger = logging.getLogger(__name__)


def run_auth_process(username: str, password: str, config: Config, config_folder: Path | None = None) -> None:
    """
    Autentica o utilizador, obtém credenciais e cria o ficheiro de configuração.
    args:
        username (str): Nome de utilizador para autenticação.
        password (str): Palavra-passe para autenticação.
        config (Config): Instância de configuração.
        config_folder (Path | None): Pasta onde o ficheiro de configuração será salvo. Se None, usa a pasta base da API.
    Returns:
        None
    """
    logger.info(_('Initiating credential generation process...'))

    excel_app_handler = None
    try:
        # Usa o construtor 'for_app_only' para interagir com o Excel sem uma folha específica
        excel_app_handler = ExcelHandler.for_app_only()

        api_service = ApiService(config=config)
        credentials = api_service.get_api_credentials(username, password)

        if not credentials.get('success'):
            error_msg = credentials.get('error', _('Authentication failed.'))
            logger.error(error_msg)
            excel_app_handler.alert_user(error_msg, _('Authentication Error'))
            return

        if config_folder:
            base_dir = config_folder
        else:
            base_dir = api_service.base_dir

        create_config_file(config_path=base_dir, api_url=api_service.api_url, credentials=credentials)

        success_msg = _(
            'Credentials generated successfully! Please restart the Excel file for the changes to take effect.'
        )
        excel_app_handler.alert_user(success_msg, _('Success'))

    except Exception as e:
        # Se qualquer passo falhar, regista o erro e tenta alertar o utilizador
        error_msg = _('An error occurred during credential generation: {error}').format(error=e)
        logger.exception(error_msg)
        if excel_app_handler:
            excel_app_handler.alert_user(error_msg, _('Error'))
        # Propaga a exceção para ser apanhada pelo 'main' e sair com código de erro
        raise


def run_create_process(excel_handler: ExcelHandler, config: Config) -> None:
    """
    Executa o fluxo completo de criação de lançamentos.
    Recebe um ExcelHandler já inicializado.
    args:
        excel_handler (ExcelHandler): Instância do ExcelHandler para manipular o Excel.
        config (Config): Instância de configuração.
    Returns:
        None
    """

    logger.info(_('Starting creation process'))

    # Ler os dados
    raw_df = excel_handler.read_data_to_dataframe()
    if raw_df.empty:
        logger.warning(_('The data table is empty. Process interrupted.'))
        excel_handler.alert_user(_('The data table is empty.'), _('Warning'))
        return

    # Processar os dados
    processor = ProcessingService(raw_df)
    data_groups = processor.group_data()

    # Enviar para a API
    api_service = ApiService(config=config)
    all_results = []
    for group_df in data_groups:
        api_response = api_service.create_journal_entry(group_df)
        result = {'indices': group_df.index, 'response': api_response}
        all_results.append(result)

    # Escrever os resultados
    excel_handler.write_results_to_sheet(all_results)

    success_message = _('Creation process completed! {num_groups} groups sent.').format(num_groups=len(data_groups))
    logger.info(success_message)
    excel_handler.alert_user(success_message, _('Success'))


def run_status_check_process(excel_handler: ExcelHandler) -> None:
    """Executa o fluxo de verificação de status."""
    logger.info(_('Starting status check process'))

    excel_handler.alert_user(_('Functionality not yet implemented.'), _('Information'))
    pass


def main() -> NoReturn:
    """
    Ponto de entrada principal. Orquestra a execução, tratamento de erros
    e interação com o Excel.
    """
    logger.info(_(f'Arguments received: {sys.argv}'))

    config = Config()

    excel_handler = None
    try:
        # Validação de argumentos
        if len(sys.argv) < 2:
            raise IndexError(_("No command provided. Expected 'auth', 'create', or 'status'."))

        command = sys.argv[1]
        logger.info(f"Command: '{command}'")

        # Roteamento de comandos
        if command == 'auth':
            if len(sys.argv) < 3:
                raise IndexError(_('Authentication command requires username and password.'))
            arg_username = sys.argv[2]
            arg_password = sys.argv[3]

            username = arg_username.strip('\'"')
            password = arg_password.strip('\'"')

            run_auth_process(username, password, config=config, config_folder=BASE_DIR)
        elif command in {'create', 'status'}:
            if len(sys.argv) < 2:
                raise IndexError(_("Command '{command}' requires a sheet index.").format(command=command))
            # sheet_index = int(sys.argv[2])

            if not config.API_KEY or not config.CLIENT_ID:
                raise PermissionError(_('API credentials not found. Please run "Generate Credentials" first.'))

            # Usa o construtor 'for_sheet' para se conectar à folha correta
            excel_handler = ExcelHandler.for_sheet(sheet_index=2)

            if command == 'create':
                run_create_process(excel_handler, config=config)
            elif command == 'status':
                run_status_check_process(excel_handler)
            else:
                logger.error(f'Comando desconhecido recebido: {command}')
                if excel_handler:
                    excel_handler.alert_user(f"Unknown command: '{command}'", _('Error'))
        else:
            logger.error(_('Unknown command received: {command}').format(command=command))
            # Tenta alertar o utilizador, criando um handler de app temporário
            ExcelHandler.for_app_only().alert_user(
                _("Unknown command: '{command}'").format(command=command), _('Error')
            )

        logger.info(_('Main process completed successfully.'))
        sys.exit(0)  # Termina com código de sucesso
    except Exception as e:
        # Centralized and robust error handling
        logger.exception(_('!!! A CRITICAL ERROR OCCURRED IN THE MAIN FLOW !!!'))

        # Tenta notificar o utilizador no Excel
        # Se o excel_handler falhou na inicialização, esta variável será None
        error_message = _(
            'An unexpected error occurred. Please check the "logs" folder for details.\n\nError: {error}'
        ).format(error=e)

        handler_to_alert = excel_handler if excel_handler else ExcelHandler.for_app_only()
        handler_to_alert.alert_user(error_message, _('Critical Error'))

        sys.exit(1)  # Garante que o VBA saiba que houve uma falha


# Bloco de teste para desenvolvimento local (não é executado pelo .exe)
def run_tests_locally() -> None:
    test_command = sys.argv[1] if len(sys.argv) > 1 else 'create'

    # Valida se o comando de teste é um dos conhecidos.
    if test_command not in {'create', 'status', 'auth'}:
        print(f"Erro: Comando de teste desconhecido '{test_command}'. Use 'create', 'status' ou 'auth'.")
        sys.exit(1)

    TEST_BASE_DIR = Path(__file__).resolve().parent
    test_sheet_index = 2
    test_file_path = str(TEST_BASE_DIR / 'excel' / 'teste_pm.xlsm')
    test_config_path = TEST_BASE_DIR / 'config.ini'

    test_config = Config(config_filepath=test_config_path)

    logger.info(f'MODO DE TESTE LOCAL (Ficheiro: {test_file_path})')

    # Inicializa a variável fora do try para que exista no finally
    excel_app_for_testing = None
    excel_ctrl = None
    try:
        # for_testing agora retorna uma tupla (controlador, app)
        # Desempacotamos os dois valores em variáveis separadas
        excel_ctrl, excel_app_for_testing = ExcelHandler.for_testing(
            filepath=test_file_path, sheet_index=test_sheet_index
        )

        if test_command == 'auth':
            if len(sys.argv) < 4:
                raise IndexError(_('Authentication command requires username and password.'))
            username = sys.argv[2]
            password = sys.argv[3]
            run_auth_process(username, password, config=test_config, config_folder=TEST_BASE_DIR / 'excel')
        elif test_command == 'create':
            run_create_process(excel_ctrl, config=test_config)

            if excel_ctrl and excel_ctrl.wb:
                # Salva o resultado do teste de criação
                excel_ctrl.wb.save()
                logger.info(f"Teste de criação concluído! O ficheiro '{test_file_path}' foi atualizado.")
            else:
                logger.error('ExcelHandler ou workbook não inicializado corretamente durante o teste de criação.')

        elif test_command == 'status':
            run_status_check_process(excel_ctrl)
            logger.info('Teste de verificação de status concluído!')

    except Exception:
        logger.exception('Ocorreu um erro durante o teste local.')

    finally:
        # Agora fechamos a app que recebemos de for_testing
        if excel_app_for_testing:
            excel_app_for_testing.quit()
        logger.info('\nFIM DO MODO DE TESTE')


if __name__ == '__main__':
    is_frozen = getattr(sys, 'frozen', False)

    try:
        if is_frozen:
            # Se for o .exe (produção), chama a função main.
            main()
        else:
            # Se for 'python main.py' (desenvolvimento/teste), chama a função de teste.
            run_tests_locally()
    except Exception:
        logger.exception('A critical unhandled error occurred at the entry point!!!')
        sys.exit(1)
