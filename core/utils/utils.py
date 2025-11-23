from pathlib import Path


def load_config_from_env(env_path: Path):
    """
    Lê um ficheiro .env e retorna um dicionário com as configurações.
    É uma implementação simples e robusta, sem dependências externas.
    """
    config = {}
    if not env_path.exists():
        # Levanta um erro se o ficheiro .env não for encontrado.
        # Isto é melhor do que falhar silenciosamente.
        raise FileNotFoundError(f"Ficheiro de configuração .env não encontrado em '{env_path}'")

    with open(env_path, 'r', encoding='utf-8') as f:
        for line in f:
            # Ignora comentários e linhas em branco
            read_line = line.strip()
            if not read_line or read_line.startswith('#'):
                continue

            # Divide a linha na primeira ocorrência de '='
            if '=' in read_line:
                key, value = read_line.split('=', 1)

                # Remove aspas (simples ou duplas) do valor
                if value.startswith('"') and value.endswith('"'):
                    value = value[1:-1]
                elif value.startswith("'") and value.endswith("'"):
                    value = value[1:-1]

                config[key.strip()] = value.strip()

    return config
