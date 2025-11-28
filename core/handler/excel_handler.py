import logging
from typing import Any, Optional

import pandas as pd
import xlwings as xw
from xlwings import App, Book, Sheet

from core.config.i18n import _
from core.config.settings import (
    END_CELL,
    END_FEEDBACK_CELL,
    EXPECTED_COLUMNS,
    START_CELL,
    START_FEEDBACK_CELL,
    STATUS_FEEDBACK_CELL,
)

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

    def find_last_row(self, key_column: str = 'N', start_row: int = 3) -> int:
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

        # Pega a coluna inteira a partir da linha inicial
        coluna_range = self.sheet.range(f'{key_column}{start_row}').expand('down')

        # Se houver apenas uma célula, o expand pode retornar um único valor em vez de uma lista
        if not isinstance(coluna_range.value, (list, tuple)):
            # Se a célula inicial tiver um valor, a última linha é a linha inicial
            if coluna_range.value is not None:
                logger.debug(
                    f'Apenas uma célula encontrada em {coluna_range.address} com valor. Última linha é {start_row}.'
                )
                return start_row
            # Se a célula inicial estiver vazia, não há dados
            else:
                logger.debug(f'A célula inicial {coluna_range.address} está vazia. Não há dados.')
                return start_row - 1

        # Itera sobre os valores da coluna
        last_row_found = start_row - 1
        for i, value in enumerate(coluna_range.value):
            if value is None:
                # Encontramos a primeira célula vazia, a linha anterior era a última
                last_row_found = start_row + i - 1
                break
        else:
            # Se o loop terminar sem encontrar uma célula vazia (sem 'break'),
            # a última linha é a última do range expandido.
            last_row_found = start_row + len(coluna_range.value) - 1

        logger.debug(_('Last data row found at row: {last_row}').format(last_row=last_row_found))
        return last_row_found

    def read_data_to_create(self) -> pd.DataFrame:
        """
        Lê um bloco de dados predefinido e filtra apenas as linhas elegíveis
        para processamento com base nas regras de negócio.

        Regras para uma linha ser elegível:
        1. A coluna 'Nominal Code' (N) deve ter um valor.
        2. A coluna '_isLocked' (AD) deve estar vazia ou conter o valor 0.
        Returns:
            pd.DataFrame: O DataFrame contendo os dados lidos.
        """
        if not self.sheet:
            error_msg = _('Cannot read data because no sheet is selected.')
            logger.error(error_msg)
            raise AttributeError(error_msg)

        logger.info(_('Reading a predefined data block to filter for processable rows...'))

        start_row = 3
        last_row = self.find_last_row()

        # Se last_row for 2, significa que não há dados abaixo do cabeçalho.
        if last_row < start_row:
            logger.info(
                _('No data rows found (column "Nominal Code" is empty starting from row {start_row}).').format(
                    start_row=start_row
                )
            )
            return pd.DataFrame()  # Retorna um DataFrame vazio

        logger.info(_('Last data row found at row: {last_row}').format(last_row=last_row))

        # Constrói o range exato para leitura, do cabeçalho até a última linha de dados.
        data_range_str = f'{START_CELL}{start_row - 1}:{END_CELL}{last_row}'

        logger.info(_('Reading data from range: {range}...').format(range=data_range_str))

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

        # Condição 1: 'Nominal Code' não pode ser nulo/vazio.
        condition_nominal_code = data_df['Nominal Code'].notna()

        # Condição 2: '_isLocked' deve ser nulo/vazio ou igual a 0.
        # .fillna(0) é seguro aqui. Usamos .astype(float) para evitar erros de tipo se houver strings.
        condition_is_locked = pd.to_numeric(data_df['_isLocked'], errors='coerce').fillna(0) == 0

        # Combina as duas condições
        processable_rows_df = data_df[condition_nominal_code & condition_is_locked].copy()
        # Usar .copy() para evitar SettingWithCopyWarning ao fazer a limpeza de espaços.

        if processable_rows_df.empty:
            logger.info(_('No processable rows found after filtering.'))
            return pd.DataFrame()

        # Limpeza de espaços em branco
        for col in processable_rows_df.select_dtypes(include=['object']).columns:
            if processable_rows_df[col].dtype == 'object':
                processable_rows_df[col] = processable_rows_df[col].str.strip()

        logger.info(
            _('{count} processable rows found within the dynamic range of {total} rows.').format(
                count=len(processable_rows_df), total=len(data_df)
            )
        )
        return processable_rows_df

    def read_data_to_update(self) -> pd.DataFrame:
        """
        Lê a tabela e filtra apenas as linhas que possuem um Número de Documento,
        mas cujo Status ainda não é 'Final'.

        Returns:
            pd.DataFrame: DataFrame com as colunas 'Document', 'Status' e 'original_row_index'.
        """
        if not self.sheet:
            raise AttributeError(_('No sheet selected.'))

        logger.info(_('Scanning for documents requiring status update...'))

        start_row = 3
        last_row = self.find_last_row(key_column=START_FEEDBACK_CELL)  # Procura pela coluna do Documento (B)

        if last_row < start_row:
            return pd.DataFrame()

        # Lê as colunas B (Document) e C (Status).
        data_range_str = f'{START_FEEDBACK_CELL}{start_row - 1}:{STATUS_FEEDBACK_CELL}{last_row}'

        # Lê para DataFrame (Header=1 assume que a linha 2 é cabeçalho)
        df = self.sheet.range(data_range_str).options(pd.DataFrame, index=False).value

        if df is None or df.empty:
            logger.info(_('No data found in the document/status range.'))
            return pd.DataFrame()

        # Adiciona o índice original do Excel para sabermos onde escrever de volta
        # (start_row + índice do dataframe)
        if len(df.columns) >= 2:
            df.columns = ['Document', 'Status']

        df['original_row_index'] = df.index + start_row

        # Deve ter Documento preenchido (não nulo e não vazio)
        has_doc = df['Document'].notna() & df['Document'].astype(str).str.strip()

        # O Status deve ser DIFERENTE de 'Final'
        # Também incluímos status vazios como candidatos a atualização.
        not_final = df['Status'].astype(str).str.strip().str.lower() != 'final'

        candidates = df[has_doc & not_final].copy()

        logger.info(_('Found {count} documents to update.').format(count=len(candidates)))

        return candidates

    def update_row_status(self, row_index: int, new_status: str | None = None, message: str | None = None) -> None:
        """
        Atualiza o Status (Col C) e opcionalmente uma mensagem (Col D ou Warning) de uma linha específica.

        Args:
            row_index (int): O índice da linha no Excel a ser atualizada (1-based)
            new_status (str | None): O novo status a ser escrito.
            message (str | None, optional): Mensagem de erro ou informação a ser escrita. Defaults to None.
        """
        if not self.sheet:
            return

        try:
            # Escreve o Status na Coluna C (3ª coluna)
            if new_status is not None:
                self.sheet.range(f'{STATUS_FEEDBACK_CELL}{row_index}').value = new_status

            # Se houver mensagem de erro/info, escreve na Coluna D
            if message is not None:
                self.sheet.range(f'{END_FEEDBACK_CELL}{row_index}').value = message

        except Exception as e:
            logger.error(_('Failed to write status to row {row_index}: {e}').format(row_index=row_index, e=e))

    def write_results_to_sheet(self, results_list: list[dict[str, Any]]) -> None:
        """
        Escreve os resultados do processamento da API de volta na folha.
        Atualiza as colunas 'Document', 'Status' e 'Warning'.
        """
        if not self.wb:
            logger.error(_('Cannot save because no workbook is active.'))
            self.alert_user(_('Could not save the file automatically because no workbook is active.'), _('Save Error'))
            return

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
        last_row = self.find_last_row()

        if last_row < start_cell:
            logger.warning(_('No data rows to update.'))
            return

        # Preenche as colunas de resultado com os dados.
        logger.info(_('Writing all results in bulk to the sheet...'))
        for result in results_list:
            response = result['response']

            if response['success']:
                values = [response.get('document'), response.get('status'), '']
                lock_value = 1  # Marca como bloqueado
            else:
                values = [_('ERROR'), _('FAILURE'), response.get('error')]
                lock_value = 0  # Mantém desbloqueado

            for df_index in result['indices']:
                excel_row = df_index + 3

                self.sheet.range(f'{START_FEEDBACK_CELL}{excel_row}').value = values
                self.sheet.range(f'{END_CELL}{excel_row}').value = lock_value

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
