import logging
import sys

from core.config.i18n import _
from core.config.logging import setup_logging
from core.handler.excel_handler import ExcelHandler
from core.services.api_service import ApiService
from core.services.processing_service import ProcessingService

setup_logging()
logger = logging.getLogger(__name__)


def run_create_process(excel_handler: ExcelHandler) -> None:
    """
    Executa o fluxo completo de criação de lançamentos.
    Recebe um ExcelHandler já inicializado.
    args:
        excel_handler (ExcelHandler): Instância do ExcelHandler para manipular o Excel.
    Returns:
        None
    """
    logger.info('Iniciando processo de criação')

    # Ler os dados
    raw_df = excel_handler.read_data_to_dataframe()
    if raw_df.empty:
        logger.warning('A tabela de dados está vazia. Processo interrompido.')
        excel_handler.alert_user(_('The data table is empty.'), _('Warning'))
        return

    # Processar os dados
    processor = ProcessingService(raw_df)
    data_groups = processor.group_data()

    # Enviar para a API
    api_service = ApiService()
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
    logger.info('Iniciando processo de verificação de status')

    excel_handler.alert_user(_('Functionality not yet implemented.'), _('Information'))
    pass


def main() -> None:
    """
    Ponto de entrada principal. Orquestra a execução, tratamento de erros
    e interação com o Excel.
    """
    logger.info(f'Argumentos recebidos: {sys.argv}')

    excel_handler = None
    try:
        # Validação de argumentos
        if len(sys.argv) < 3:  # noqa: PLR2004
            raise IndexError('Argumentos insuficientes. Esperado: <comando> <indice_planilha>')

        command = sys.argv[1]
        sheet_index = int(sys.argv[2])
        logger.info(f"Comando: '{command}', Planilha: {sheet_index}")

        # Inicializa o controlador do Excel uma única vez
        excel_handler = ExcelHandler(sheet_index=sheet_index)

        # Roteamento de comandos
        if command == 'create':
            run_create_process(excel_handler)
        elif command == 'status':
            run_status_check_process(excel_handler)
        else:
            logger.error(f'Comando desconhecido recebido: {command}')
            if excel_handler:
                excel_handler.alert_user(f"Unknown command: '{command}'", _('Error'))

        logger.info('Processo principal concluído com sucesso.')

    except Exception as e:
        # Tratamento de erro centralizado e robusto
        logger.exception('!!! OCORREU UM ERRO CRÍTICO NO FLUXO PRINCIPAL !!!')

        # Tenta notificar o utilizador no Excel
        # Se o excel_handler falhou na inicialização, esta variável será None
        if excel_handler:
            error_message = _(
                "An unexpected error occurred. Please check the 'logs' folder for details.\n\nError: {error}"
            ).format(error=e)
            excel_handler.alert_user(error_message, _('Critical Error'))
        else:
            # Fallback se a conexão com o Excel nem sequer foi estabelecida
            # (Este print só será visível no console de depuração do VBA)
            error_message = _(
                'An unexpected error occurred before connecting to Excel. '
                + "Please check the 'logs' folder for details.\n\nError: {error}"
            ).format(error=e)
            logger.error(error_message)

        sys.exit(1)  # Garante que o VBA saiba que houve uma falha


# Bloco de teste para desenvolvimento local (não é executado pelo .exe)
if __name__ == '__main__' and not getattr(sys, 'frozen', False):
    test_command = sys.argv[1] if len(sys.argv) > 1 else 'create'

    # Valida se o comando de teste é um dos conhecidos.
    if test_command not in {'create', 'status'}:
        print(f"Erro: Comando de teste desconhecido '{test_command}'. Use 'create' ou 'status'.")
        sys.exit(1)

    test_sheet_index = 1
    test_file_path = 'dados_de_teste.xlsx'

    logger.info(f'MODO DE TESTE LOCAL (Ficheiro: {test_file_path})')

    # Inicializa a variável fora do try para que exista no finally
    excel_app_for_testing = None
    try:
        # for_testing agora retorna uma tupla (controlador, app)
        # Desempacotamos os dois valores em variáveis separadas
        excel_ctrl, excel_app_for_testing = ExcelHandler.for_testing(
            filepath=test_file_path, sheet_index=test_sheet_index
        )

        if test_command == 'create':
            run_create_process(excel_ctrl)
            # Salva o resultado do teste de criação
            excel_ctrl.wb.save()
            logger.info(f"Teste de criação concluído! O ficheiro '{test_file_path}' foi atualizado.")

        elif test_command == 'status':
            run_status_check_process(excel_ctrl)
            logger.info('Teste de verificação de status concluído!')

        # Salva o resultado do teste
        excel_ctrl.wb.save()
        logger.info(f"Teste concluído! O ficheiro '{test_file_path}' foi atualizado.")

    except Exception:
        logger.exception('Ocorreu um erro durante o teste local.')

    finally:
        # Agora fechamos a app que recebemos de for_testing
        if excel_app_for_testing:
            excel_app_for_testing.quit()
        logger.info('\nFIM DO MODO DE TESTE')
