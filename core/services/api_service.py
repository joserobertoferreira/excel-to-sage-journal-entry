import json
import logging
from typing import Any

import pandas as pd
import requests

from core.auth.auth import generate_auth_headers
from core.config.settings import DIMENSIONS_MAPPING, PRIMARY_GROUP_COLUMN, SERVER_BASE_ADDRESS

logger = logging.getLogger(__name__)


class ApiService:
    """
    Serviço para interagir com a API externa.
    """

    def __init__(self):
        logger.info('Inicializar ApiService...')

        self.api_url = SERVER_BASE_ADDRESS
        self.dimensions = DIMENSIONS_MAPPING

    def _create_dimensions(self, row: pd.Series) -> dict:
        """
        Cria o dicionário de dimensões para uma única linha.
        Args:
            row (pd.Series): A linha do DataFrame representando um item de lançamento.
        Returns:
            dict: O dicionário de dimensões.
        """

        dimensions_dict = {}
        for api_key, excel_col in self.dimensions.items():
            value = row.get(excel_col)
            if pd.notna(value) and value:
                dimensions_dict[api_key] = value
        return dimensions_dict

    def _create_line_item(self, row: pd.Series, header_data: pd.Series) -> dict:
        """
        Constrói um único dicionário de 'line' a partir de uma linha do DataFrame.
        Args:
            row (pd.Series): A linha do DataFrame representando um item de lançamento.
            header_data (pd.Series): A linha do DataFrame representando os dados do cabeçalho.
        Returns:
            dict: O dicionário formatado para a linha do lançamento.
        """
        # Campos Principais e Obrigatórios
        account_code = str(int(row['Nominal Code'])) if pd.notna(row['Nominal Code']) else ''

        line_description = (
            row['Line Description'] if pd.notna(row['Line Description']) else header_data['Header Description']
        )

        line_dict = {
            'account': account_code,
            'lineDescription': line_description,
        }

        # Lógica de Valor (Quantity vs Debit/Credit)
        quantity = row.get('Quantity')

        if pd.notna(quantity) and quantity != 0:
            line_dict['quantity'] = float(quantity)
        else:
            debit = row.get('Debit')
            if pd.notna(debit) and debit != 0:
                line_dict['debit'] = round(float(debit), 2)

            credit = row.get('Credit')
            if pd.notna(credit) and credit != 0:
                line_dict['credit'] = round(float(credit), 2)

        # Campos Opcionais com Transformação
        optional_fields = {
            'businessPartner': 'BP',
            'freeReference': 'Free Reference',
            'taxCode': 'Tax',
        }

        for api_key, excel_col in optional_fields.items():
            value = row.get(excel_col)
            if pd.notna(value) and value:
                processed_value = str(value).strip()
                if api_key in {'businessPartner', 'taxCode'}:
                    processed_value = processed_value.upper()
                line_dict[api_key] = processed_value

        # Dimensões
        dimensions = self._create_dimensions(row)
        if dimensions:
            line_dict['dimensions'] = dimensions

        return line_dict

    def _build_create_journal_input(self, group_df: pd.DataFrame) -> dict:
        """
        Constrói o dicionário de variáveis para a mutação CreateJournalEntry.
        Args:
            group_df (pd.DataFrame): DataFrame contendo os dados do grupo.
        Returns:
            dict: O dicionário de variáveis para a mutação GraphQL.
        """
        header_data = group_df.iloc[0]
        lines_list = [self._create_line_item(row, header_data) for _, row in group_df.iterrows()]

        return {
            'input': {
                'site': header_data['Site'].upper(),
                'documentType': header_data['Entry Type'].upper(),
                'accountingDate': pd.to_datetime(header_data['AccountingDate']).strftime('%Y-%m-%d'),
                'descriptionByDefault': header_data['Header Description'],
                'sourceCurrency': header_data['Curr'].upper(),
                'reference': header_data['Reference'],
                'lines': lines_list,
            }
        }

    def _build_graphql_input(self, group_df: pd.DataFrame) -> dict:
        """
        Constrói o dicionário do input para a mutação GraphQL a partir de um grupo de dados (DataFrame).
        Args:
            group_df (pd.DataFrame): DataFrame contendo os dados do grupo.
        Returns:
            dict: O dicionário do input para a mutação GraphQL.
        """

        # Pega os dados da primeira linha, pois já foram validados e preenchidos
        header_data = group_df.iloc[0]

        # Constrói a lista de linhas para a mutação
        lines_list = [self._create_line_item(row, header_data) for _, row in group_df.iterrows()]

        # Constrói o objeto 'input' principal
        variables = {
            'input': {
                'site': header_data['Site'].upper(),
                'documentType': header_data['Entry Type'].upper(),
                'accountingDate': pd.to_datetime(header_data['AccountingDate']).strftime('%Y-%m-%d'),
                'descriptionByDefault': header_data['Header Description'],
                'sourceCurrency': header_data['Curr'].upper(),
                'reference': header_data['Reference'],
                'lines': lines_list,
            }
        }
        return variables

    def _execute_graphql(self, query: str, variables: dict[str, Any], operation_name: str) -> dict:
        """
        Executa uma query/mutação GraphQL genérica e trata a comunicação e os erros.
        """
        logger.info(f"A executar operação GraphQL: '{operation_name}'...")

        payload = {
            'operationName': operation_name,
            'query': query,
            'variables': variables,
        }

        auth_headers = generate_auth_headers()

        logger.debug('Payload GraphQL: %s', json.dumps(payload, indent=2))

        try:
            response = requests.post(self.api_url, headers=auth_headers, data=json.dumps(payload), timeout=60)
            response.raise_for_status()
            return response.json()

        except requests.exceptions.HTTPError as http_err:
            logger.error(f"Erro HTTP na operação '{operation_name}': {http_err}. Resposta: {http_err.response.text}")
            # Retorna uma estrutura de erro consistente do GraphQL
            return {'errors': [{'message': f'Erro {http_err.response.status_code}. Ver logs.'}]}

        except requests.exceptions.RequestException as req_err:
            logger.error(f"Erro de rede na operação '{operation_name}': {req_err}")
            return {'errors': [{'message': 'Erro de rede. Verifique a conexão com a API.'}]}

    def create_journal_entry(self, group_df: pd.DataFrame) -> dict:
        """
        Envia uma mutação GraphQL para criar um lançamento contabilístico.

        Args:
            group_df (pd.DataFrame): DataFrame contendo os dados a serem enviados.

        Returns:
            dict: A resposta da API em formato JSON.

        Raises:
            requests.exceptions.RequestException: Em caso de falha de rede ou erro HTTP.
        """
        group_id = group_df.iloc[0].get(PRIMARY_GROUP_COLUMN, 'desconhecido')
        logger.info(f"A criar lançamento para o grupo '{group_id}'...")

        # Construir a mutação GraphQL
        query = """
        mutation CreateJournalEntry($input: CreateJournalEntryInput!) {
          createJournalEntry(input: $input) {
            journalEntryNumber
            journalEntryStatus
          }
        }
        """
        variables = self._build_create_journal_input(group_df)

        # Executar a mutação com tratamento de erros
        response_data = self._execute_graphql(query, variables, 'CreateJournalEntry')

        # Processa a resposta específica desta mutação
        if 'errors' in response_data:
            error_messages = [e.get('message') for e in response_data['errors']]
            return {'success': False, 'error': '; '.join(error_messages)}

        result = response_data.get('data', {}).get('createJournalEntry', {})
        logger.info(f"Grupo '{group_id}' enviado com sucesso. Documento: {result.get('journalEntryNumber')}")
        return {
            'success': True,
            'document': result.get('journalEntryNumber'),
            'status': result.get('journalEntryStatus'),
        }

    def get_journal_statuses(self, document_numbers: list[str]) -> dict:
        """Busca o status de uma lista de documentos."""
        logger.info(f'A verificar o status para {len(document_numbers)} documentos...')

        # A sua query de status virá aqui. Exemplo hipotético:
        query = """
        query GetJournalStatuses($numbers: [String!]!) {
          journalEntries(filter: { numbers: $numbers }) {
            number
            status
          }
        }
        """
        variables = {'numbers': document_numbers}

        response_data = self._execute_graphql(query, variables, 'GetJournalStatuses')

        # A sua lógica para processar a resposta da query de status virá aqui.
        # Por exemplo, transformar a lista de resultados num dicionário para fácil acesso.
        if 'errors' in response_data:
            # ... tratamento de erro ...
            return {}

        results = response_data.get('data', {}).get('journalEntries', [])
        status_map = {entry['number']: entry['status'] for entry in results}
        return status_map
