import logging

import numpy as np
import pandas as pd

from core.config.i18n import _
from core.config.settings import GROUPING_COLUMNS, PRIMARY_GROUP_COLUMN, SECONDARY_GROUP_COLUMNS

logger = logging.getLogger(__name__)


class ProcessingService:
    """
    Encapsula toda a lógica de negócio para validar e transformar os dados lidos do Excel.
    """

    def __init__(self, df: pd.DataFrame):
        if df.empty:
            raise ValueError(_('The initial DataFrame cannot be empty for processing.'))

        self.df = df.copy()  # Trabalha com uma cópia para não alterar o original

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
        # Converte strings vazias para NaN para que ffill funcione
        self.df[GROUPING_COLUMNS] = self.df[GROUPING_COLUMNS].replace('', np.nan)
        # Preenche os valores para baixo
        self.df[GROUPING_COLUMNS] = self.df[GROUPING_COLUMNS].ffill()
        # Converte quaisquer NaNs restantes (ex: colunas totalmente vazias) para strings vazias
        self.df[GROUPING_COLUMNS] = self.df[GROUPING_COLUMNS].fillna('')
        logger.info(_('Preprocessing completed.'))

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

    def group_data(self) -> list[pd.DataFrame]:
        """
        Agrupa o DataFrame pré-processado em uma lista de DataFrames menores,
        um para cada conjunto de dados a ser enviado para a API.
        Returns:
            list[pd.DataFrame]: Lista de DataFrames agrupados.
        """
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

        return data_sets
