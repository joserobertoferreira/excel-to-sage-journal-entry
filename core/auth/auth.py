import hashlib
import hmac
import time

from core.config import settings


def generate_auth_headers() -> dict:
    """
    Gera os cabeçalhos de autenticação HMAC dinâmicos para a API.

    Esta é uma função "pura": não tem efeitos colaterais e retorna um resultado
    previsível com base no tempo e nas configurações.

    Returns:
        dict: Um dicionário contendo os cabeçalhos de autenticação necessários.
    """

    # Gerar timestamp
    timestamp = int(time.time())

    # Criar a mensagem a ser assinada
    message_to_sign = f'{settings.API_KEY}{settings.CLIENT_ID}{timestamp}'

    # Gerar a assinatura HMAC
    signature = hmac.new(
        key=settings.API_SECRET.encode('utf-8'), msg=message_to_sign.encode('utf-8'), digestmod=hashlib.sha256
    ).hexdigest()

    # Montar e retornar o objeto de cabeçalhos de autenticação
    auth_headers = {
        'content-type': 'application/json',
        'Accept': '*/*',
        'X-App-Key': settings.API_KEY,
        'X-Client-Id': settings.CLIENT_ID,
        'X-Timestamp': str(timestamp),
        'X-Signature': signature,
    }

    return auth_headers
