# Guia de Gestão de Traduções (i18n)

Este documento descreve o processo passo a passo para adicionar novas línguas e atualizar as traduções existentes na aplicação, utilizando a ferramenta `pybabel`.

## Pré-requisitos

Antes de começar, certifique-se de que o seu ambiente de desenvolvimento está configurado:

1.  **Ambiente Virtual Ativo:** Certifique-se de que o seu ambiente virtual (`.venv`) está ativo.
2.  **Dependências Instaladas:** Todas as dependências de desenvolvimento, incluindo `Babel`, devem estar instaladas. Se não estiverem, execute:
    ```bash
    uv pip install -e .[dev]
    ```
3.  **Ficheiros de Configuração:** O projeto deve conter os seguintes ficheiros na sua raiz:
    *   `pyproject.toml` (com `babel` nas dependências de desenvolvimento).
    *   `babel.cfg` (com a configuração de extração).
4.  **Pasta `locales`:** Uma pasta chamada `locales` deve existir na raiz do projeto. Se não existir, crie-a com o comando `mkdir locales`.

---

## Fluxo de Trabalho

Existem três processos principais: criar uma nova língua, atualizar as existentes e compilar as traduções.

### A. Como Adicionar uma Nova Língua

Execute este processo apenas **uma vez** para cada nova língua que pretende suportar (ex: Francês `fr`, Espanhol `es`).

**1. Extrair as Strings do Código-Fonte**

Este comando varre todo o código Python, encontra as strings marcadas para tradução (ex: `_("Hello")`) e cria um ficheiro modelo `messages.pot` na pasta `locales`.

```bash
uv run pybabel extract -F babel.cfg -o locales/messages.pot .
```
*(Lembre-se do `.` no final do comando, que indica o diretório atual).*

**2. Inicializar o Ficheiro da Nova Língua**

Este comando cria a estrutura de pastas e o ficheiro `.po` para a nova língua, usando o modelo gerado no passo anterior. Substitua `[codigo_da_lingua]` pelo código apropriado (ex: `fr`, `es`, `de`).

```bash
uv run pybabel init -i locales/messages.pot -d locales -l [codigo_da_lingua]
```
**Exemplo para Francês:**
```bash
uv run pybabel init -i locales/messages.pot -d locales -l fr
```

**Próximo Passo:** Agora, edite o novo ficheiro criado em `locales/[codigo_da_lingua]/LC_MESSAGES/messages.po` e preencha as traduções nas linhas `msgstr ""`.

---

### B. Como Atualizar as Línguas Existentes

Este é o processo mais comum. Execute-o sempre que **adicionar ou modificar strings traduzíveis** no código Python.

**1. Re-extrair as Strings para Atualizar o Modelo**

Primeiro, atualize o ficheiro `messages.pot` com as últimas alterações do código.

```bash
uv run pybabel extract -F babel.cfg -o locales/messages.pot .
```

**2. Sincronizar os Ficheiros de Tradução**

Este comando compara o `messages.pot` atualizado com todos os ficheiros `.po` existentes e adiciona as novas strings que precisam de ser traduzidas, sem apagar as traduções já feitas.

```bash
uv run pybabel update -i locales/messages.pot -d locales
```

**Próximo Passo:** Edite os ficheiros `.po` em cada pasta de língua para traduzir as novas strings que foram adicionadas.

---

### C. Como Compilar as Traduções

Após traduzir ou atualizar os ficheiros `.po`, é necessário compilá-los para o formato binário `.mo` que a aplicação utiliza em tempo de execução.

**1. Compilar Todos os Ficheiros de Língua**

Este comando encontra todos os ficheiros `.po` na pasta `locales` e gera o ficheiro `.mo` correspondente para cada um.

```bash
uv run pybabel compile -d locales
```

**Importante:** Após compilar, você precisa de **gerar novamente o executável com o PyInstaller** para que ele inclua os novos ficheiros `.mo` no pacote final.

---