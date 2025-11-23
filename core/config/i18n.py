import gettext
import locale
import logging

# Importa as configurações do nosso ficheiro de settings
from .settings import DEFAULT_LANGUAGE, LOCALE_DIR, SUPPORTED_LANGUAGES

logger = logging.getLogger(__name__)


def get_best_language() -> str:
    """
    Determina a melhor língua a ser usada com base nas configurações do sistema
    e nas línguas suportadas pela aplicação.
    """
    try:
        # Tenta obter a língua do sistema (ex: 'pt_PT', 'en_US')
        system_lang, _ = locale.getdefaultlocale()
    except Exception:
        # Fallback se não conseguir obter a língua do sistema
        system_lang = DEFAULT_LANGUAGE

    if system_lang in SUPPORTED_LANGUAGES:
        logger.info(f"Língua do sistema '{system_lang}' encontrada e suportada.")
        return system_lang

    # Se a língua completa (ex: 'en_US') não estiver na lista,
    # tenta a versão curta (ex: 'en').
    short_lang = system_lang.split('_')[0] if system_lang else DEFAULT_LANGUAGE
    if short_lang in SUPPORTED_LANGUAGES:
        logger.info(f"Língua do sistema '{system_lang}' não é suportada, a usar fallback para '{short_lang}'.")
        return short_lang

    logger.warning(f"A língua do sistema '{system_lang}' não é suportada. A usar a língua padrão '{DEFAULT_LANGUAGE}'.")
    return DEFAULT_LANGUAGE


# Lógica para carregar as traduções

# Determina qual a língua a carregar
active_language = get_best_language()

# Instala a tradução para a língua selecionada
try:
    # 'base' é o nosso domínio de tradução
    # languages=[active_language] diz ao gettext qual a língua específica a procurar
    translation = gettext.translation('messages', localedir=str(LOCALE_DIR), languages=[active_language])
    translation.install()
    _ = translation.gettext
    logger.info(f"Tradução para '{active_language}' carregada com sucesso.")
except FileNotFoundError:
    # Se o ficheiro .mo para a língua ativa não for encontrado,
    # instala a função de fallback que não faz nada (retorna o texto original).
    logger.warning(f"Ficheiro de tradução para '{active_language}' não encontrado. Usar o texto original (inglês).")
    _ = gettext.gettext
