import logging
import sys
from pathlib import Path
from typing import NoReturn

from core.config.config import Config
from core.config.i18n import _
from core.config.logging import setup_logging
from core.config.settings import BASE_DIR
from core.handler.excel_handler import ExcelHandler
from core.services.processing_service import ProcessingService

setup_logging()
logger = logging.getLogger(__name__)


def main() -> NoReturn:
    """
    Ponto de entrada principal. Orquestra a execução, tratamento de erros
    e interação com o Excel.
    """
    config = Config()
    processing_service = ProcessingService(config=config)

    try:
        # Validação de argumentos
        if len(sys.argv) < 2:
            raise IndexError(_("No command provided. Expected 'auth', 'create', or 'status'."))

        command = sys.argv[1]

        # Roteamento de comandos
        if command == 'auth':
            if len(sys.argv) < 3:
                raise IndexError(_('Authentication command requires username and password.'))

            arg_username = sys.argv[2]
            arg_password = sys.argv[3]

            username = arg_username.strip('\'"')
            password = arg_password.strip('\'"')

            logger.info(
                _('Arguments received: {argument_0}, {argument_1}, {argument_2}').format(
                    argument_0=command, argument_1=username, argument_2='*' * len(password)
                )
            )

            processing_service.run_auth_process(username, password, config_folder=BASE_DIR)

        elif command in {'create', 'status'}:
            if len(sys.argv) < 2:
                raise IndexError(_("Command '{command}' requires a sheet index.").format(command=command))
            # sheet_index = int(sys.argv[2])

            if not config.API_KEY or not config.CLIENT_ID:
                raise PermissionError(_('API credentials not found. Please run "Generate Credentials" first.'))

            if command == 'create':
                processing_service.run_create_process()
            elif command == 'status':
                processing_service.run_status_check_process()
        else:
            error_message = _('Unknown command received: {command}').format(command=command)
            logger.error(error_message)
            # # Tenta alertar o utilizador, criando um handler de app temporário
            processing_service.alert_user_critical_error(error_message)

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

        processing_service.alert_user_critical_error(error_message)

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
    test_file_path = str(TEST_BASE_DIR / 'excel' / 'teste_sage.xlsm')
    test_config_path = TEST_BASE_DIR / 'config.ini'

    test_config = Config(config_filepath=test_config_path)

    logger.info(f'MODO DE TESTE LOCAL (Ficheiro: {test_file_path})')

    # Inicializa a variável fora do try para que exista no finally
    excel_app_for_testing = None
    excel_ctrl = None
    try:
        processing_service = ProcessingService(config=test_config)

        if test_command == 'auth':
            if len(sys.argv) < 4:
                raise IndexError(_('Authentication command requires username and password.'))
            username = sys.argv[2]
            password = sys.argv[3]
            processing_service.run_auth_process(username, password, config_folder=TEST_BASE_DIR / 'excel')

        elif test_command in {'create', 'status'}:
            excel_ctrl, excel_app_for_testing = ExcelHandler.for_testing(
                filepath=test_file_path, sheet_index=test_sheet_index
            )

            if test_command == 'create':
                processing_service.run_create_process(excel_handler_override=excel_ctrl)
            elif test_command == 'status':
                processing_service.run_status_check_process(excel_handler_override=excel_ctrl)
                logger.info('Teste de verificação de status concluído!')

            if excel_ctrl and excel_ctrl.wb:
                excel_ctrl.wb.save()
                logger.info(f"Teste concluído! O ficheiro '{test_file_path}' foi atualizado.")

        else:
            logger.error('ExcelHandler ou workbook não inicializado corretamente durante o teste de criação.')

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
