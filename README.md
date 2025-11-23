# Guia do Utilizador - Ferramenta de Lançamento Contabilístico

Siga estes passos cuidadosamente para garantir o funcionamento correto da ferramenta.

### 1. Configuração Inicial (Apenas na Primeira Utilização)

Antes de usar a ferramenta pela primeira vez, precisa de gerar e configurar as suas credenciais de acesso à API.

1.  **Gerar Credenciais:** Aceda ao seguinte endereço no seu navegador: `[COLOQUE AQUI O URL PÚBLICO DA SUA API]`
2.  Siga as instruções na página para gerar as suas credenciais pessoais. Ser-lhe-ão fornecidos três valores: `APP_KEY`, `CLIENT_ID`, e `API_SECRET`. Copie estes três valores.
3.  **Configurar a Ferramenta:** Na mesma pasta onde está o ficheiro Excel, encontrará um ficheiro chamado **`config.ini`**.
4.  Abra o `config.ini` com o Bloco de Notas ou outro editor de texto.
5.  Cole os valores que copiou nos campos correspondentes. O ficheiro deve ficar semelhante a isto:

    ```ini
    # Ficheiro de Configuração da Aplicação
    API_URL="http://endereco_da_api/graphql"
    
    # Suas credenciais geradas
    APP_KEY="valor_que_voce_copiou_aqui"
    CLIENT_ID="outro_valor_que_voce_copiou_aqui"
    API_SECRET="o_seu_segredo_copiado_aqui"
    ```
6.  Guarde e feche o ficheiro `config.ini`. Este passo só precisa de ser feito uma vez.

---

### 2. Preparação do Ficheiro de Trabalho

1.  Abra o ficheiro modelo **`journal_entry_template.xlsm`**.
2.  Imediatamente, vá ao menu **Ficheiro > Guardar Como**.
3.  Escolha um local e dê um nome ao seu ficheiro (ex: `Lancamentos_Novembro_2025`).
4.  **PASSO MAIS IMPORTANTE:** No campo **"Tipo"**, certifique-se de que a opção **"Livro do Excel Ativado com Macros (*.xlsm)"** está selecionada.
5.  Clique em **Guardar**.

> **AVISO:** Se guardar o ficheiro como um `.xlsx` normal, os botões de automação deixarão de funcionar. O ficheiro **deve** ser sempre do tipo `.xlsm`.

### 3. Preenchimento dos Dados

Preencha as linhas da folha de cálculo com os dados do lançamento.

*   **Agrupamento Automático:** Para que o sistema agrupe os lançamentos automaticamente, deixe a coluna **`Group By`** vazia.
*   **Agrupamento Manual:** Para definir os seus próprios grupos, insira um identificador (ex: "A", "GRUPO1") na primeira linha de cada novo grupo na coluna **`Group By`**.

### 4. Execução da Ferramenta

1.  Certifique-se de que os programas **`SageJournalEntry.exe`** e **`config.ini`** estão na mesma pasta que o seu ficheiro `.xlsm`.
2.  Clique no botão **"Create entries"** para validar e enviar os dados para a API.
3.  Aguarde a mensagem de "Success" ou "Error". As colunas `Document`, `Status` e `Warning` serão atualizadas automaticamente.
4.  (Opcional) Após a criação de um lançamento, pode usar o botão **"Update status"** para atualizar o estado dos documentos existentes.

---

# User Guide - Journal Entry Tool

Please follow these steps carefully to ensure the tool functions correctly.

### 1. First-Time Setup: Generating and Configuring Credentials

Before using the tool for the first time, you need to generate and configure your API access credentials.

1.  **Generate Credentials:** Access the following URL in your browser: `[INSERT YOUR PUBLIC API URL HERE]`
2.  Follow the instructions on the page to generate your personal credentials. You will be provided with three values: `APP_KEY`, `CLIENT_ID`, and `API_SECRET`. Copy these three values.
3.  **Configure the Tool:** In the same folder where the Excel file is located, you will find a file named **`config.ini`**.
4.  Open `config.ini` with Notepad or another text editor.
5.  Paste the values you copied into the corresponding fields. The file should look similar to this:

    ```ini
    # Application Configuration File
    API_URL="http://api_address_here/graphql"
    
    # Your generated credentials
    APP_KEY="value_you_copied_here"
    CLIENT_ID="another_value_you_copied_here"
    API_SECRET="your_secret_copied_here"
    ```
6.  Save and close the `config.ini` file. This step only needs to be done once.

---

### 2. Preparing Your Working File

1.  Open the template file **`journal_entry_template.xlsm`**.
2.  Immediately, go to the **File > Save As** menu.
3.  Choose a location and give your file a name (e.g., `Journal_Entries_November_2025`).
4.  **MOST IMPORTANT STEP:** In the **"Save as type"** field, make sure the option **"Excel Macro-Enabled Workbook (*.xlsm)"** is selected.
5.  Click **Save**.

> **WARNING:** If you save the file as a standard `.xlsx` workbook, the automation buttons will stop working. The file **must** always be an `.xlsm` type.

### 3. Filling in the Data

Fill in the spreadsheet rows with the journal entry data.

*   **Automatic Grouping:** To have the system group entries automatically, leave the **`Group By`** column empty.
*   **Manual Grouping:** To define your own groups, enter an identifier (e.g., "A", "GROUP1") in the first row of each new group in the **`Group By`** column.

### 4. Running the Tool

1.  Ensure that the **`SageJournalEntry.exe`** program and the **`config.ini`** file are in the same folder as your `.xlsm` file.
2.  Click the **"Create entries"** button to validate and send the data to the API.
3.  Wait for the "Success" or "Error" message. The `Document`, `Status`, and `Warning` columns will be updated automatically.
4.  (Optional) After an entry has been created, you can use the **"Update status"** button to update the status of existing documents.