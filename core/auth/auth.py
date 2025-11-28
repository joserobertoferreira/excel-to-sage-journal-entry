import hashlib
import hmac
import time

from core.config.config import Config
from core.config.settings import EXCEL_API_KEY


def generate_auth_headers(config: Config, admin: bool) -> dict:
    """
    Gera os cabeçalhos de autenticação HMAC dinâmicos para a API.

    Esta é uma função "pura": não tem efeitos colaterais e retorna um resultado
    previsível com base no tempo e nas configurações.

    Args:
        config (Config): A instância de configuração contendo as credenciais da API.
        admin (bool): Indica se as credenciais de administrador devem ser usadas.
    Returns:
        dict: Um dicionário contendo os cabeçalhos de autenticação necessários.
    """

    # Gerar timestamp
    timestamp = int(time.time())

    # Criar a mensagem a ser assinada
    if admin:
        auth_headers = {
            'content-type': 'application/json',
            'Accept': '*/*',
            'X-Admin-Key': EXCEL_API_KEY,
        }
    else:
        secret_key = config.API_SECRET
        app_key = config.API_KEY
        client_id = config.CLIENT_ID

        message_to_sign = f'{app_key}{client_id}{timestamp}'

        # Gerar a assinatura HMAC
        signature = hmac.new(
            key=secret_key.encode('utf-8'), msg=message_to_sign.encode('utf-8'), digestmod=hashlib.sha256
        ).hexdigest()

        # Montar e retornar o objeto de cabeçalhos de autenticação
        auth_headers = {
            'content-type': 'application/json',
            'Accept': '*/*',
            'X-App-Key': app_key,
            'X-Client-Id': client_id,
            'X-Timestamp': str(timestamp),
            'X-Signature': signature,
        }

    return auth_headers
