import logging

import numpy as np
import pandas as pd

from core.config.i18n import _
from core.config.settings import (
    COLUMNS_TO_UPPERCASE,
    DATE_COLUMNS,
    GROUPING_COLUMNS,
    PRIMARY_GROUP_COLUMN,
    SECONDARY_GROUP_COLUMNS,
)

logger = logging.getLogger(__name__)


class ValidationService:
    """
    Encapsula toda a lógica de negócio para validar e transformar os dados lidos do Excel.
    """

    def __init__(self, df: pd.DataFrame):
        if df.empty:
            raise ValueError(_('The initial DataFrame cannot be empty for processing.'))

        self.df = df.copy()  # Trabalha com uma cópia para não alterar o original

    def _validate_data_structure(self, max_lines: int) -> None:
        """
        Executa validações na estrutura geral do DataFrame.
        1. Verifica se o número total de linhas excede o limite do ERP.
        2. Verifica se existem linhas em branco ('buracos') no meio dos dados.
        """
        logger.info(_('Validating data structure (row limit and contiguity)...'))

        # 1. Validação do Limite de Linhas ---
        total_rows = len(self.df)
        if total_rows > max_lines:
            error_msg = _(
                'Data validation failed: The total number of rows ({total}) exceeds the maximum allowed limit of {max}.'
            ).format(total=total_rows, max=max_lines)
            logger.error(error_msg)
            raise ValueError(error_msg)

        logger.info(
            _('Row limit validation passed: {total_rows}/{max_lines} rows.').format(
                total_rows=total_rows, max_lines=max_lines
            )
        )

        # 2. Validação de Contiguidade (Sem "Buracos") ---
        # A coluna 'Nominal Code' é a nossa chave para definir uma linha como "não vazia".
        key_column = 'Nominal Code'

        # Cria uma série booleana: True se a linha tem um 'Nominal Code'.
        is_valid_row = self.df[key_column].notna() & self.df[key_column]

        # Se não houver nenhuma linha válida, não há nada a validar.
        if not is_valid_row.any():
            logger.info(_('No valid data rows found; skipping contiguity validation.'))
            return

        # Encontra o índice da primeira e da última linha com dados.
        first_valid_index = is_valid_row.idxmax()
        last_valid_index = is_valid_row[::-1].idxmax()

        # Seleciona o "miolo" do DataFrame, entre a primeira e a última linha válida.
        # .loc[a:b] inclui tanto 'a' quanto 'b'.
        core_data_slice = self.df.loc[first_valid_index:last_valid_index]

        # A validação chave: Há alguma linha DENTRO deste miolo que seja inválida?
        if core_data_slice[key_column].isna().any() or (core_data_slice[key_column] == '').any():  # noqa: PLC1901
            # Encontra a primeira linha em branco para dar um erro mais útil
            first_blank_row_index = core_data_slice[
                core_data_slice[key_column].isna() | (core_data_slice[key_column] == '')
            ].index[0]
            # Converte o índice do DataFrame para o número da linha do Excel
            excel_row_num = first_blank_row_index + 3

            error_msg = _(
                'Data validation failed: Blank rows were found in the middle of the data. '
                "The 'Nominal Code' column cannot be empty between rows with data. "
                'Please check Excel row {row_num}.'
            ).format(row_num=excel_row_num)
            logger.error(error_msg)
            raise ValueError(error_msg)

        logger.info(_('Contiguity validation (no gaps) passed.'))

    def _validate_initial_data(self):
        """
        Verifica se a primeira linha de dados tem pelo menos um valor
        nas colunas de agrupamento.
        """
        logger.info(_('Validating initial data...'))

        # Pega a primeira linha do DataFrame (índice 0) e seleciona apenas as colunas de agrupamento
        first_row = self.df.iloc[0]
        first_row_grouping_data = first_row[GROUPING_COLUMNS]

        # .isnull() retorna True para NaN/None. .all() verifica se TODOS os valores são True.
        # Também verificamos se todas são strings vazias, pois isso também é inválido.
        is_all_null = first_row_grouping_data.isnull().all()
        is_all_empty_string = (first_row_grouping_data == '').all()  # noqa: PLC1901

        if is_all_null or is_all_empty_string:
            # Se todas as células de agrupamento na primeira linha estiverem vazias, levanta um erro.
            error_message = _(
                'Validation failed: The first data row (row 3 in Excel) '
                'cannot have all grouping columns empty. '
                'Please fill at least one of the following columns: {columns}'
            ).format(columns=', '.join(GROUPING_COLUMNS))
            # Levantar um ValueError é apropriado aqui. Ele será capturado pelo bloco try/except no main.py.
            logger.error(error_message)
            raise ValueError(error_message)

    def _preprocess_data(self) -> None:
        """Prepara o DataFrame preenchendo os valores de agrupamento para baixo."""
        logger.info(_('Preprocessing data (filling grouping columns downwards)...'))

        for col in COLUMNS_TO_UPPERCASE:
            # Verifica se a coluna existe no DataFrame antes de tentar a conversão
            if col in self.df.columns:
                # .str.upper() converte toda a coluna de uma só vez.
                # pd.notna(...) garante que não tentamos aplicar .str a valores nulos (NaN),
                # o que evita erros. A atribuição é feita apenas para as linhas não nulas.
                non_null_mask = self.df[col].notna()
                self.df.loc[non_null_mask, col] = self.df.loc[non_null_mask, col].astype(str).str.upper()

        reversing_col_name = 'Reversing Y/N (1=No 2=Yes)'
        if reversing_col_name in self.df.columns:
            # Converte para numérico (lidando com erros) e preenche com 1
            self.df[reversing_col_name] = pd.to_numeric(self.df[reversing_col_name], errors='coerce').fillna(1)

        logger.info(_('Validate and format date columns...'))

        for col in DATE_COLUMNS:
            if col in self.df.columns:
                # Converte apenas as linhas que não estão vazias
                non_null_mask = self.df[col].notna()
                if not non_null_mask.any():
                    continue  # Pula para a próxima coluna se esta estiver toda vazia

                # Tenta converter para data. 'coerce' transforma inválidos em NaT.
                # 'dayfirst=True' ajuda o pandas a priorizar o formato DD/MM (comum na Europa)
                converted_dates = pd.to_datetime(self.df.loc[non_null_mask, col], errors='coerce', dayfirst=True)

                # Verifica se a conversão criou algum NaT (data inválida)
                if converted_dates.isnull().any():
                    # Encontra o primeiro erro para dar um feedback útil
                    first_error_index = converted_dates[converted_dates.isnull()].index[0]
                    excel_row_num = first_error_index + 3
                    original_value = self.df.loc[first_error_index, col]

                    error_msg = _(
                        "Data validation failed: Invalid date format found in column '{col_name}' "
                        "at Excel row {row_num}. Value: '{value}'"
                    ).format(col_name=col, row_num=excel_row_num, value=original_value)
                    logger.error(error_msg)
                    raise ValueError(error_msg)

                # Se todas as datas são válidas, formata para YYYY-MM-DD para consistência
                self.df.loc[non_null_mask, col] = converted_dates.dt.strftime('%Y-%m-%d')

        logger.info(_('Date validation completed.'))

        # Converte strings vazias para NaN para que ffill funcione
        self.df[GROUPING_COLUMNS] = self.df[GROUPING_COLUMNS].replace('', np.nan)
        # Preenche os valores para baixo
        self.df[GROUPING_COLUMNS] = self.df[GROUPING_COLUMNS].ffill()
        # Converte quaisquer NaNs restantes (ex: colunas totalmente vazias) para strings vazias
        self.df[GROUPING_COLUMNS] = self.df[GROUPING_COLUMNS].fillna('')
        logger.info(_('Preprocessing completed.'))

    def _validate_group_headers(self, data_groups: list[pd.DataFrame]) -> None:
        """
        Verifica se cada grupo de dados tem uma 'Header Description' válida.
        """
        logger.info(_('Validating the headers of each data group...'))

        for i, group_df in enumerate(data_groups):
            # A primeira linha do grupo contém o valor de cabeçalho para todo o grupo (devido ao ffill)
            header_row = group_df.iloc[0]
            description = header_row.get('Header Description')

            if pd.isna(description) or not str(description).strip():
                # Tenta obter o ID do grupo para uma mensagem de erro mais útil
                group_id = header_row.get(PRIMARY_GROUP_COLUMN, f'Grupo #{i + 1}')

                error_msg = _(
                    "Data validation failed for group '{group_id}': "
                    "The 'Header Description' column cannot be empty for a data group."
                ).format(group_id=group_id)
                logger.error(error_msg)
                raise ValueError(error_msg)

        logger.info(_('Group headers validation successful.'))

    def _validate_group_consistency(self):
        """
        Valida a consistência dos grupos definidos pelo usuário.
        Para cada grupo em 'Group By', verifica se as colunas secundárias têm apenas um valor único.
        """
        logger.info(_('Validating group consistency...'))

        # Agrupa pelo 'Group By' já preenchido
        groups = self.df.groupby(PRIMARY_GROUP_COLUMN)

        # Itera sobre as colunas que precisam ser consistentes
        for col in SECONDARY_GROUP_COLUMNS:
            # nunique() conta o número de valores únicos por grupo
            unique_counts = groups[col].nunique()

            # Se qualquer grupo tiver mais de 1 valor único, há uma inconsistência
            if (unique_counts > 1).any():
                # Encontra o primeiro grupo e valor inconsistente para a mensagem de erro
                inconsistent_groups = unique_counts[unique_counts > 1]
                group_name = inconsistent_groups.index[0]

                # Filtra as linhas para obter um DataFrame temporário
                inconsistent_df = self.df[self.df[PRIMARY_GROUP_COLUMN] == group_name]

                # Extrai a coluna (Series) desse DataFrame e obtém os valores únicos
                inconsistent_values = inconsistent_df[col].unique()

                error_message = _(
                    "Data consistency error in group '{group_name}'.\n"
                    "The column '{col}' has multiple values ({values}) "
                    "within the same group defined in 'Group By'. "
                    'Please correct the data or create a new group.'
                ).format(group_name=group_name, col=col, values=list(inconsistent_values))
                logger.error(error_message)
                raise ValueError(error_message)

        logger.info(_('Group consistency validated successfully.'))

    def _generate_automatic_groups(self) -> None:
        """Gera IDs de grupo sequenciais quando a coluna 'Group By' está vazia."""
        logger.info(_('Mode: Automatic group generation.'))

        # Cria um ID de grupo sequencial baseado na mudança de valores nas colunas secundárias
        group_starts = (self.df[SECONDARY_GROUP_COLUMNS] != self.df[SECONDARY_GROUP_COLUMNS].shift()).any(axis=1)
        group_ids = group_starts.cumsum()
        self.df[PRIMARY_GROUP_COLUMN] = group_ids
        logger.info(_('Automatic groups generated successfully.'))

    def group_data(self, max_lines: int) -> list[pd.DataFrame]:
        """
        Agrupa o DataFrame pré-processado em uma lista de DataFrames menores,
        um para cada conjunto de dados a ser enviado para a API.
        Args:
            max_lines (int): O número máximo de linhas permitidas no DataFrame.
        Returns:
            list[pd.DataFrame]: Lista de DataFrames agrupados.
        """
        # Valida a estrutura geral dos dados
        self._validate_data_structure(max_lines=max_lines)

        # Valida os dados iniciais
        self._validate_initial_data()

        # Substitui vazios por NaN e preenche TUDO para baixo.
        self._preprocess_data()

        # Verificar se a coluna 'Group By' foi preenchida pelo usuário.
        # .str.len() > 0 é uma forma segura de verificar se há strings não vazias
        # E .any() verifica se pelo menos uma linha satisfaz a condição
        user_defined_groups = self.df[PRIMARY_GROUP_COLUMN].astype(str).str.len().gt(0).any()

        if user_defined_groups:
            # Grupos definidos pelo utilizador
            logger.info(_('Mode: User-defined groups.'))

            # Executa a validação de consistência
            self._validate_group_consistency()

        else:
            # Grupos automáticos
            logger.info(_('Mode: Automatic groups.'))

            self._generate_automatic_groups()

        # Agrupamento final
        logger.info(_('Grouping by column: {column}...').format(column=PRIMARY_GROUP_COLUMN))

        grouped_data = self.df.groupby(PRIMARY_GROUP_COLUMN, sort=False)
        data_sets = [group_df for _, group_df in grouped_data]

        logger.info(_('Data divided into {count} sequential sets.').format(count=len(data_sets)))

        # Valida os cabeçalhos de cada grupo
        if data_sets:
            self._validate_group_headers(data_sets)

        return data_sets
