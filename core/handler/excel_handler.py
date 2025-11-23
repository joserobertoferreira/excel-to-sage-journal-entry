import logging
from typing import Any

import pandas as pd
import xlwings as xw
from xlwings import App

from core.config.i18n import _
from core.config.settings import END_CELL, END_FEEDBACK_CELL, EXPECTED_COLUMNS, START_CELL, START_FEEDBACK_CELL

logger = logging.getLogger(__name__)


class ExcelHandler:
    """
    Controla toda a interação com a pasta de trabalho do Excel.
    Responsável por ler os dados da tabela e escrever os resultados.
    """

    def __init__(self, sheet_index: int):
        """
        Conecta-se à instância ativa do Excel e à folha especificada.
        Esta abordagem é robusta para executáveis chamados via VBA.
        """
        try:
            logger.info('A tentar conectar-se à instância ativa do Excel...')

            # xw.apps.active aponta para a última instância do Excel que foi usada.
            app = xw.apps.active
            if not app:
                raise ConnectionError('Nenhuma instância ativa do Excel encontrada.')

            # wb.active é a pasta de trabalho que está em primeiro plano.
            self.wb = app.books.active
            if not self.wb:
                raise ConnectionError('Nenhuma pasta de trabalho ativa encontrada na instância do Excel.')

            logger.info(f'Conectado com sucesso à pasta de trabalho: {self.wb.name}')

            # Seleciona a folha pelo seu índice (VBA é 1-based, xlwings é 0-based).
            self.sheet = self.wb.sheets[sheet_index - 1]

            logger.info(f"Selecionada a folha: '{self.sheet.name}' (índice {sheet_index})")

        except Exception as e:
            logger.error('Falha ao inicializar o ExcelController.')
            # Propaga o erro para que a função main possa capturá-lo.
            raise e

    @classmethod
    def for_testing(cls, filepath: str, sheet_index: int = 1) -> tuple['ExcelHandler', App]:
        """
        Abre um ficheiro específico para testes e retorna tanto o controlador
        quanto a instância da aplicação Excel para que possa ser fechada.
        """
        instance = None
        app = None
        try:
            # Cria a instância da classe normalmente
            instance = cls.__new__(cls)

            # Cria uma nova instância invisível do Excel
            app = xw.App(visible=False)

            # Abre a pasta de trabalho e configura a instância
            instance.wb = app.books.open(filepath)
            instance.sheet = instance.wb.sheets[sheet_index - 1]

            # Retorna uma tupla: (o controlador, a app a ser fechada)
            return instance, app

        except Exception as e:
            # Se algo falhar, garante que a app do Excel seja fechada
            if app:
                app.quit()
            logger.error(f"Erro ao abrir o ficheiro de teste '{filepath}': {e}")
            raise e

    def alert_user(self, message: str, title: str = 'Informação') -> None:
        """
        Exibe uma caixa de diálogo (alerta) para o utilizador no Excel.

        Args:
            message (str): A mensagem a ser exibida.
            title (str, optional): O título da caixa de diálogo. Defaults to "Informação".
        """
        try:
            if self.wb and self.wb.app:
                self.wb.app.alert(message, title)
                logger.info(f"Alerta exibido ao utilizador: '{title}' - '{message}'")
        except Exception:
            # Se não conseguir mostrar o alerta (ex: o Excel foi fechado),
            # pelo menos regista-o no log sem quebrar a aplicação.
            logger.error(f"Falha ao tentar exibir o alerta ao utilizador: '{title}' - '{message}'")

    def find_last_row(self, key_column: str = 'O', start_row: int = 3) -> int:
        """
        Encontra a última linha de dados com base na coluna informada.
        A tabela termina na primeira linha em que a coluna está vazia.
        """
        logger.debug(f"Procurar a última linha usando a coluna '{key_column}' a partir da linha {start_row}.")
        start_cell = self.sheet.range(f'{key_column}{start_row}')

        # .end('down') vai para a última célula preenchida ANTES de uma célula vazia.
        # Se O3 estiver preenchida e O4 vazia, ele retornará o range de O3.
        # Se O3 a O10 estiverem preenchidas e O11 vazia, ele retornará O10.
        # Se O3 estiver vazia, ele irá para baixo até encontrar uma célula preenchida e parar acima dela,
        # o que pode dar um resultado inesperado. Por isso, primeiro verificamos se O3 está vazia.
        if start_cell.value is None:
            logger.debug(f"A célula inicial '{key_column}{start_row}' está vazia. Não há dados.")
            return start_row - 1

        # .end('down') é mais fiável quando não há linhas em branco no meio.
        # Se puder haver linhas em branco, uma abordagem de loop seria mais segura,
        # mas para dados contíguos, isto é perfeito.
        last_cell = start_cell.end('down')
        logger.debug(f'A última célula com dados encontrada foi {last_cell.address}.')
        return last_cell.row

    def read_data_to_dataframe(self) -> pd.DataFrame:
        """Lê os dados da tabela para um DataFrame, determinando o tamanho dinamicamente."""
        logger.info("A determinar o tamanho da tabela com base na coluna 'O'...")

        start_cell = 3
        last_row = self.find_last_row(key_column='O', start_row=start_cell)

        # Se last_row for 2, significa que não há dados abaixo do cabeçalho.
        if last_row < start_cell:
            logger.info("Nenhuma linha de dados encontrada (a coluna 'O' está vazia a partir da linha 3).")
            return pd.DataFrame()  # Retorna um DataFrame vazio

        logger.info(f'Última linha de dados encontrada na linha: {last_row}')

        # Constrói o range exato para leitura, do cabeçalho até a última linha de dados.
        data_range_str = f'{START_CELL}{start_cell - 1}:{END_CELL}{last_row}'

        logger.info(f'Lendo dados do range: {data_range_str}...')

        # .options(pd.DataFrame, header=1, index=False, expand='table') é a forma mais robusta
        # de ler diretamente para um DataFrame.
        # header=1: Diz que a primeira linha do range (A2) é o cabeçalho.
        # index=False: Não tenta usar a primeira coluna como índice do DataFrame.
        data_df = self.sheet.range(data_range_str).options(pd.DataFrame, index=False).value

        # Limpeza de espaços em branco em colunas de texto
        for col in data_df.select_dtypes(include=['object']).columns:
            if data_df[col].dtype == 'object':  # Garante que é uma coluna de strings
                data_df[col] = data_df[col].str.strip()

        # Validação final para garantir que as colunas lidas correspondem ao esperado
        # Isso ajuda a pegar erros de layout na folha
        if list(data_df.columns) != EXPECTED_COLUMNS:
            logger.warning('Os cabeçalhos na folha não correspondem exatamente a EXPECTED_COLUMNS no config.py.')
            raise ValueError('Inconsistência de cabeçalho entre o Excel e a configuração.')

        # Opcional: Renomear as colunas do DataFrame para garantir consistência
        data_df.columns = EXPECTED_COLUMNS

        logger.info(f'Foram lidas e processadas {len(data_df)} linhas de dados.')
        return data_df

    def write_results_to_sheet(self, results_list: list[dict[str, Any]]) -> None:
        """
        Escreve os resultados do processamento da API de volta na folha.
        Atualiza as colunas 'Document', 'Status' e 'Warning'.
        """
        if not results_list:
            logger.info('Nenhum resultado para escrever na folha.')
            return

        logger.info(f'Escrever {len(results_list)} resultados na folha...')

        # Limpa o conteúdo antigo das colunas de resultado para evitar confusão
        # (Usa o mesmo last_row da leitura para ser consistente)
        start_cell = 3
        last_row = self.find_last_row(key_column='O', start_row=start_cell)

        if last_row < start_cell:
            logger.warning('Não há linhas de dados para atualizar.')
            return

        # Limpa a área de feedback de uma só vez
        feedback_range = f'{START_FEEDBACK_CELL}{start_cell}:{END_FEEDBACK_CELL}{last_row}'
        logger.info(f'A limpar a área de feedback: {feedback_range}')
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
        logger.info('A escrever todos os resultados em bloco na folha...')
        self.sheet.range(f'{START_FEEDBACK_CELL}{start_cell}').value = output_data
        logger.info('Escrita em bloco concluída.')

        # Salvamento automático do ficheiro para garantir que os resultados não se percam
        try:
            logger.info(f'Salvar a pasta de trabalho: {self.wb.name}...')
            self.wb.save()
            logger.info('Pasta de trabalho salva com sucesso.')
        except Exception:
            logger.error(f'Falha ao salvar a pasta de trabalho {self.wb.name}.')
            self.alert_user(
                _(
                    'Results were written to the sheet, but an error occurred while saving the file. '
                    'Please save the file manually.'
                ),
                _('Error Saving File'),
            )
