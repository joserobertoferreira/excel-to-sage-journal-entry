import json
import logging
from typing import Any

import pandas as pd
import requests
from requests.exceptions import HTTPError, RequestException

from core.auth.auth import generate_auth_headers
from core.config.config import Config
from core.config.i18n import _
from core.config.settings import BASE_DIR, DIMENSIONS_MAPPING, PRIMARY_GROUP_COLUMN

logger = logging.getLogger(__name__)


class ApiService:
    """
    Serviço para interagir com a API externa.
    """

    def __init__(self, config: Config):
        logger.info(_('Start ApiService...'))

        self.config = config
        self.api_url = config.SERVER_BASE_ADDRESS
        self.base_dir = BASE_DIR
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
            if pd.notna(value):
                str_value = str(value).strip()

                if str_value:
                    dimensions_dict[api_key] = str_value

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
        try:
            nominal_code = row['Nominal Code']
            if pd.notna(nominal_code):
                account_code = str(int(nominal_code)).strip()
            else:
                account_code = ''
        except (ValueError, TypeError):
            # Fallback caso venha texto na coluna de conta
            account_code = str(row.get('Nominal Code', '')).strip()

        line_desc_val = row.get('Line Description')
        if pd.notna(line_desc_val) and str(line_desc_val).strip():
            line_description = str(line_desc_val).strip()
        else:
            line_description = header_data['Header Description']

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

    def _execute_graphql(
        self,
        query: str,
        variables: dict[str, Any],
        operation_name: str,
        authorization: bool = True,
        admin: bool = False,
    ) -> dict:
        """
        Executa uma query/mutação GraphQL genérica e trata a comunicação e os erros.
        Args:
            query (str): A query ou mutação GraphQL.
            variables (dict): As variáveis para a operação.
            operation_name (str): O nome da operação GraphQL.
            authorization (bool): Indica se deve incluir cabeçalhos de autorização.
            admin (bool): Indica se deve usar credenciais de administrador.
        Returns:
            dict: A resposta da API em formato JSON.
        Raises:
            HTTPError: Em caso de erro HTTP.
            RequestException: Em caso de falha de rede ou erro HTTP.
        """
        logger.info(_('Execute GraphQL API: "{operation_name}"...').format(operation_name=operation_name))

        payload = {
            'operationName': operation_name,
            'query': query,
            'variables': variables,
        }

        if authorization:
            auth_headers = generate_auth_headers(config=self.config, admin=admin)
        else:
            auth_headers = {'content-type': 'application/json', 'Accept': '*/*'}

        try:
            response = requests.post(self.api_url, headers=auth_headers, data=json.dumps(payload), timeout=60)
            response.raise_for_status()
            return response.json()

        except HTTPError as http_err:
            logger.error(
                _('HTTP error in operation "{operation_name}": {http_err}. Response: {response_text}').format(
                    operation_name=operation_name, http_err=http_err, response_text=http_err.response.text
                )
            )
            # Retorna uma estrutura de erro consistente do GraphQL
            error_message = _('HTTP error {http_err.response.status_code} occurred. Check logs.').format(
                http_err=http_err
            )
            return {'errors': [{'message': error_message}]}

        except RequestException as req_err:
            logger.error(
                _('Network error in operation "{operation_name}": {req_err}').format(
                    operation_name=operation_name, req_err=req_err
                )
            )
            error_message = _('Network error. Please check the API connection.')
            return {'errors': [{'message': error_message}]}

    def create_journal_entry(self, group_df: pd.DataFrame) -> dict:
        """
        Envia uma mutação GraphQL para criar um lançamento contabilístico.

        Args:
            group_df (pd.DataFrame): DataFrame contendo os dados a serem enviados.

        Returns:
            dict: A resposta da API em formato JSON.
        """
        group_id = group_df.iloc[0].get(PRIMARY_GROUP_COLUMN, _('unknown'))
        logger.info(_('Create journal entry for group "{group_id}"...').format(group_id=group_id))

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

        logger.info(
            _('Group "{group_id}" successfully sent. Document: {document_number}').format(
                group_id=group_id, document_number=result.get('journalEntryNumber')
            )
        )

        return {
            'success': True,
            'document': result.get('journalEntryNumber'),
            'status': result.get('journalEntryStatus'),
        }

    def get_journal_status(self, document_number: str) -> dict:
        """
        Busca o status do documento informado
        Args:
            document_number (str): O número do documento a ser verificado.
        Returns:
            dict: Um dicionário contendo o status do documento.
        """
        logger.info(_('Checking status for document "{document_number}"...').format(document_number=document_number))

        # A sua query de status virá aqui. Exemplo hipotético:
        query = """
        query GetJournalEntryStatus($input: JournalEntryInputUnique!) {
            getJournalEntryStatus(input: $input) {
                journalEntryNumber
                journalEntryStatus
                journalEntryType
            }
        }
        """

        variables = {'input': {'journalEntryNumber': document_number}}

        response_data = self._execute_graphql(query, variables, 'GetJournalEntryStatus')

        if 'errors' in response_data:
            error_messages = [e.get('message') for e in response_data['errors']]
            return {'success': False, 'error': '; '.join(error_messages)}

        results = response_data.get('data', {}).get('getJournalEntryStatus', {})
        status_map = {results.get('journalEntryNumber'): results.get('journalEntryStatus')}
        return status_map

    def _build_api_credential_input(self, username: str, password: str) -> dict:  # noqa: PLR6301
        """
        Constrói o dicionário de variáveis para a mutação CreateApiCredential
        Args:
            username (str): O nome de usuário para autenticação.
            password (str): A senha para autenticação.
        Returns:
            dict: O dicionário de variáveis para a mutação GraphQL.
        """

        return {
            'input': {
                'login': username.lower(),
                'password': password,
            }
        }

    def get_api_credentials(self, username: str, password: str) -> dict:
        """
        Chama o endpoint de autenticação para gerar as credenciais da API.
        Args:
            username (str): O nome de usuário para autenticação.
            password (str): A senha para autenticação.
        Returns:
            dict: Um dicionário contendo as credenciais da API ou mensagens de erro.
        Raises:
            HTTPError: Em caso de erro HTTP.
            RequestException: Em caso de falha de rede ou erro HTTP.
        """
        logger.info(_('Attempting to authenticate user: {username}').format(username=username))

        # Construir a mutação GraphQL
        query = """
        query GetApiCredential($input: GetApiCredentialInput!) {
            getApiCredential(input: $input) {
                appKey
                appSecret
                clientId
            }
        }
        """

        variables = self._build_api_credential_input(username.lower(), password)

        # Executar a query com tratamento de erros
        response_data = self._execute_graphql(query, variables, 'GetApiCredential', authorization=True, admin=True)

        # Processa a resposta específica desta mutação
        if 'errors' in response_data:
            error_messages = [e.get('message') for e in response_data['errors']]
            return {'success': False, 'error': '; '.join(error_messages)}

        result = response_data.get('data', {}).get('getApiCredential', {})

        required_keys = ['appKey', 'appSecret', 'clientId']

        if not all(key in result for key in required_keys):
            logger.error(_('Authentication response is missing required keys.'))
            return {'success': False, 'error': _('Invalid response from authentication server.')}

        app_key = result.get('appKey')

        if app_key is None or not app_key.strip():
            logger.info(_('Credentials will be generated upon first use.'))

            response_data = self.create_api_credentials(username, password)

            # Processa a resposta específica desta mutação
            if 'errors' in response_data:
                error_messages = [e.get('message') for e in response_data['errors']]
                return {'success': False, 'error': '; '.join(error_messages)}

            result = response_data.get('data', {}).get('createApiCredential', {})

        return {'success': True, **result}

    def create_api_credentials(self, username: str, password: str) -> dict:
        """
        Chama o endpoint de criação para gerar as credenciais da API.
        Args:
            username (str): O nome de usuário para autenticação.
            password (str): A senha para autenticação.
        Returns:
            dict: Um dicionário contendo as credenciais da API ou mensagens de erro.
        Raises:
            HTTPError: Em caso de erro HTTP.
            RequestException: Em caso de falha de rede ou erro HTTP.
        """
        logger.info(_('Create user credentials for user: {username}').format(username=username))

        # Construir a mutação GraphQL
        query = """
        mutation CreateApiCredential($input: CreateApiCredentialInput!) {
            createApiCredential(input: $input) {
                appKey
                appSecret
                clientId
                name
            }
        }
        """

        variables = self._build_api_credential_input(username.lower(), password)

        # Executar a mutação com tratamento de erros
        response_data = self._execute_graphql(
            query, variables, operation_name='CreateApiCredential', authorization=False
        )

        # Processa a resposta específica desta mutação
        if 'errors' in response_data:
            error_messages = [e.get('message') for e in response_data['errors']]
            return {'success': False, 'error': '; '.join(error_messages)}

        logger.info(_('Credentials for user {username} successfully created').format(username=username.lower()))

        return {'success': True, **response_data}
