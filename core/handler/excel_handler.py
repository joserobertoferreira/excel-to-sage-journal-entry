import logging
from typing import Any, Optional

import pandas as pd
import xlwings as xw
from xlwings import App, Book, Sheet

from core.config.i18n import _
from core.config.settings import END_CELL, END_FEEDBACK_CELL, EXPECTED_COLUMNS, START_CELL, START_FEEDBACK_CELL

logger = logging.getLogger(__name__)


class ExcelHandler:
    """
    Controla toda a interação com a pasta de trabalho do Excel.
    Responsável por ler os dados da tabela e escrever os resultados.
    """

    def __init__(self):
        """
        Construtor base. Inicializa todos os atributos como None.
        Use os construtores de classe (@classmethod) para criar instâncias funcionais.
        """
        self.app: Optional[App] = None
        self.wb: Optional[Book] = None
        self.sheet: Optional[Sheet] = None

    @classmethod
    def for_sheet(cls, sheet_index: int) -> 'ExcelHandler':
        """
        Construtor principal: Conecta-se à instância ativa do Excel e a uma folha específica.
        Args:
            sheet_index (int): O índice da folha a ser selecionada (1-based).
        Returns:
            ExcelHandler: Instância conectada à folha especificada.
        """
        instance = cls()  # Cria uma instância com todos os atributos como None

        try:
            logger.info(_('Connecting to active Excel instance and sheet...'))

            # xw.apps.active aponta para a última instância do Excel que foi usada.
            instance.app = xw.apps.active
            if not instance.app:
                raise ConnectionError(_('No active Excel instance found.'))

            # wb.active é a pasta de trabalho que está em primeiro plano.
            instance.wb = instance.app.books.active
            if not instance.wb:
                raise ConnectionError(_('No active workbook found in the Excel instance.'))

            logger.info(_('Successfully connected to workbook: {workbook_name}').format(workbook_name=instance.wb.name))

            # Seleciona a folha pelo seu índice (VBA é 1-based, xlwings é 0-based).
            instance.sheet = instance.wb.sheets[sheet_index - 1]

            if instance.sheet:
                logger.info(
                    _('Selected sheet: {sheet_name} (index {sheet_index})').format(
                        sheet_name=instance.sheet.name, sheet_index=sheet_index
                    )
                )
            else:
                raise ValueError(_('Sheet with index {sheet_index} does not exist.').format(sheet_index=sheet_index))
            return instance
        except Exception as e:
            logger.error(_('Failed to initialize ExcelHandler.'))
            # Propagate the error so the main function can catch it.
            raise e

    @classmethod
    def for_app_only(cls) -> 'ExcelHandler':
        """
        Um construtor alternativo que se conecta apenas à aplicação Excel,
        sem selecionar uma folha específica. Útil para alertas globais.
        """
        instance = cls()
        try:
            logger.info(_('Connecting to active Excel application instance...'))
            instance.app = xw.apps.active
            if not instance.app:
                raise ConnectionError(_('No active Excel instance found.'))

            # Tenta obter a pasta de trabalho ativa, mas não falha se não houver
            instance.wb = instance.app.books.active
            return instance
        except Exception:
            logger.warning(_('Failed to fully connect to Excel application.'), exc_info=True)
            return instance  # Retorna a instância parcialmente vazia

    @classmethod
    def for_testing(cls, filepath: str, sheet_index: int = 1) -> tuple['ExcelHandler', App]:
        """
        Abre um ficheiro específico para testes e retorna tanto o controlador
        quanto a instância da aplicação Excel para que possa ser fechada.
        """
        app = None
        try:
            # Cria a instância da classe normalmente
            instance = cls()

            # Cria uma nova instância invisível do Excel
            app = xw.App(visible=False)
            instance.app = app

            # Abre a pasta de trabalho e configura a instância
            instance.wb = app.books.open(filepath)
            instance.sheet = instance.wb.sheets[sheet_index - 1]

            # Retorna uma tupla: (o controlador, a app a ser fechada)
            return instance, app

        except Exception as e:
            # Se algo falhar, garante que a app do Excel seja fechada
            if app:
                app.quit()
            logger.error(_('Erro ao abrir o ficheiro de teste {filepath}: {error}').format(filepath=filepath, error=e))
            raise e

    def alert_user(self, message: str, title: str = 'Information') -> None:
        """
        Exibe uma caixa de diálogo (alerta) para o utilizador no Excel.

        Args:
            message (str): A mensagem a ser exibida.
            title (str, optional): O título da caixa de diálogo. Defaults to "Information".
        """
        if self.app:
            try:
                self.app.alert(message, title)
                logger.info(_('Alert displayed: {title} - {message}').format(title=title, message=message))
            except Exception:
                # Se não conseguir mostrar o alerta (ex: o Excel foi fechado),
                # pelo menos regista-o no log sem quebrar a aplicação.
                logger.error(_('Failed to display alert: {title} - {message}').format(title=title, message=message))
        else:
            logger.warning(_('No Excel app instance available to display the alert..'))

    def find_last_row(self, key_column: str = 'O', start_row: int = 3) -> int:
        """
        Encontra a última linha de dados com base na coluna informada.
        A tabela termina na primeira linha em que a coluna está vazia.
        """
        if not self.sheet:
            error_msg = _('Cannot find last row because no sheet is selected.')
            logger.error(error_msg)
            raise AttributeError(error_msg)

        logger.debug(
            _('Searching for the last row using column {key_column} starting from row {start_row}.').format(
                key_column=key_column, start_row=start_row
            )
        )
        start_cell = self.sheet.range(f'{key_column}{start_row}')

        # .end('down') vai para a última célula preenchida ANTES de uma célula vazia.
        # Se O3 estiver preenchida e O4 vazia, ele retornará o range de O3.
        # Se O3 a O10 estiverem preenchidas e O11 vazia, ele retornará O10.
        # Se O3 estiver vazia, ele irá para baixo até encontrar uma célula preenchida e parar acima dela,
        # o que pode dar um resultado inesperado. Por isso, primeiro verificamos se O3 está vazia.
        if start_cell.value is None:
            logger.debug(
                _('The starting cell {start_cell_address} is empty. No data found.').format(
                    start_cell_address=f'{key_column}{start_row}'
                )
            )
            return start_row - 1

        # .end('down') é mais fiável quando não há linhas em branco no meio.
        # Se puder haver linhas em branco, uma abordagem de loop seria mais segura,
        # mas para dados contíguos, isto é perfeito.
        last_cell = start_cell.end('down')
        logger.debug(
            _('The last cell with data found was {last_cell_address}.').format(last_cell_address=last_cell.address)
        )
        return last_cell.row

    def read_data_to_dataframe(self) -> pd.DataFrame:
        """
        Lê os dados da tabela para um DataFrame, determinando o tamanho dinamicamente.
        Returns:
            pd.DataFrame: O DataFrame contendo os dados lidos.
        """
        if not self.sheet:
            error_msg = _('Cannot read data because no sheet is selected.')
            logger.error(error_msg)
            raise AttributeError(error_msg)

        logger.info(_('Determining the table size based on column "O"...'))

        start_cell = 3
        last_row = self.find_last_row(key_column='O', start_row=start_cell)

        # Se last_row for 2, significa que não há dados abaixo do cabeçalho.
        if last_row < start_cell:
            logger.info(
                _('No data rows found (column "O" is empty starting from row {start_cell}).').format(
                    start_cell=start_cell
                )
            )
            return pd.DataFrame()  # Retorna um DataFrame vazio

        logger.info(_('Last data row found at row: {last_row}').format(last_row=last_row))

        # Constrói o range exato para leitura, do cabeçalho até a última linha de dados.
        data_range_str = f'{START_CELL}{start_cell - 1}:{END_CELL}{last_row}'

        logger.info(_('Reading data from range: {data_range_str}...').format(data_range_str=data_range_str))

        # .options(pd.DataFrame, header=1, index=False, expand='table') é a forma mais robusta
        # de ler diretamente para um DataFrame.
        # header=1: Diz que a primeira linha do range (A2) é o cabeçalho.
        # index=False: Não tenta usar a primeira coluna como índice do DataFrame.
        data_df = self.sheet.range(data_range_str).options(pd.DataFrame, index=False).value

        if data_df is None or data_df.empty:
            logger.info(_('The data table is empty after reading.'))
            return pd.DataFrame()

        # Limpeza de espaços em branco em colunas de texto
        for col in data_df.select_dtypes(include=['object']).columns:
            if data_df[col].dtype == 'object':  # Garante que é uma coluna de strings
                data_df[col] = data_df[col].str.strip()

        # Validação final para garantir que as colunas lidas correspondem ao esperado
        # Isso ajuda a pegar erros de layout na folha
        if list(data_df.columns) != EXPECTED_COLUMNS:
            logger.warning(_('The headers in the sheet do not match exactly with EXPECTED_COLUMNS in config.py.'))
            raise ValueError(_('Header inconsistency between Excel and the configuration.'))

        # Opcional: Renomear as colunas do DataFrame para garantir consistência
        data_df.columns = EXPECTED_COLUMNS

        logger.info(_('{count} rows of data were read and processed.').format(count=len(data_df)))
        return data_df

    def write_results_to_sheet(self, results_list: list[dict[str, Any]]) -> None:
        """
        Escreve os resultados do processamento da API de volta na folha.
        Atualiza as colunas 'Document', 'Status' e 'Warning'.
        """
        if not self.sheet:
            error_msg = _('Cannot write results because no sheet is selected.')
            logger.error(error_msg)
            raise AttributeError(error_msg)

        if not results_list:
            logger.info(_('No results to write to the sheet.'))
            return

        logger.info(_('Writing {count} results to the sheet...').format(count=len(results_list)))

        # Limpa o conteúdo antigo das colunas de resultado para evitar confusão
        # (Usa o mesmo last_row da leitura para ser consistente)
        start_cell = 3
        last_row = self.find_last_row(key_column='O', start_row=start_cell)

        if last_row < start_cell:
            logger.warning(_('No data rows to update.'))
            return

        # Limpa a área de feedback de uma só vez
        feedback_range = f'{START_FEEDBACK_CELL}{start_cell}:{END_FEEDBACK_CELL}{last_row}'
        logger.info(_('Clearing feedback area: {feedback_range}').format(feedback_range=feedback_range))
        self.sheet.range(feedback_range).clear_contents()

        # Cria uma "matriz" (lista de listas) vazia com o tamanho exato necessário.
        # Terá (last_row - start_row + 1) linhas e 3 colunas (Doc, Status, Warning).
        num_rows = last_row - start_cell + 1
        output_data = [[''] * 3 for _ in range(num_rows)]

        # Preenche a matriz com os resultados.
        for result in results_list:
            response = result['response']

            if response['success']:
                values = [response.get('document'), response.get('status'), '']
            else:
                values = [_('ERROR'), _('FAILURE'), response.get('error')]

            for df_index in result['indices']:
                # O índice da matriz é o índice do DataFrame.
                matrix_index = df_index
                if matrix_index < len(output_data):
                    output_data[matrix_index] = values

        # Escreve os dados no Excel. Esta abordagem é mais lenta que em bloco,
        # mas mais segura para não sobrescrever formatações.
        # 3. Escreve a matriz inteira no Excel de uma SÓ VEZ.
        # O range começa em B3 (START_FEEDBACK_CELL).
        if not self.wb:
            logger.error(_('Cannot save because no workbook is active.'))
            self.alert_user(_('Could not save the file automatically because no workbook is active.'), _('Save Error'))
            return

        logger.info(_('Writing all results in bulk to the sheet...'))
        self.sheet.range(f'{START_FEEDBACK_CELL}{start_cell}').value = output_data
        logger.info(_('Bulk write completed.'))

        # Salvamento automático do ficheiro para garantir que os resultados não se percam
        try:
            logger.info(_('Saving the workbook: {wb_name}...').format(wb_name=self.wb.name))
            self.wb.save()
            logger.info(_('Workbook saved successfully.'))
        except Exception:
            logger.error(_('Failed to save the workbook {wb_name}.').format(wb_name=self.wb.name))
            self.alert_user(
                _(
                    'Results were written to the sheet, but an error occurred while saving the file. '
                    'Please save the file manually.'
                ),
                _('Error Saving File'),
            )
