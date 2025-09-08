# Automação de Cadastro de Previsão de Vendas - TOTVS Protheus

## Descrição

Este projeto é uma ferramenta de automação de processos (RPA) desenvolvida em Python para automatizar o cadastro de "Previsão de Vendas" no sistema TOTVS Protheus (WebApp). A aplicação possui uma interface gráfica amigável construída com CustomTkinter, que permite ao usuário fornecer suas credenciais, selecionar um arquivo Excel (`.xlsx`) com os dados e iniciar/parar o processo de automação.

A automação é executada em segundo plano para não travar a interface, e o usuário recebe notificações por e-mail (via Outlook) sobre o sucesso ou a falha da operação.

## Funcionalidades Principais

- **Interface Gráfica Intuitiva:** Criada com CustomTkinter para uma experiência de usuário moderna e agradável.
- **Automação em Background:** Utiliza `threading` para que a interface permaneça responsiva durante a execução do robô.
- **Controle de Execução:** Botões para "Iniciar" e "Parar" a automação a qualquer momento.
- **Processamento de Dados via Excel:** Lê os dados diretamente de uma planilha `.xlsx` utilizando a biblioteca Pandas.
- **Navegação Complexa:** O robô (Selenium) é capaz de navegar pela complexa estrutura do Protheus WebApp, incluindo a manipulação de **Shadow DOM** e **iframes**.
- **Notificações por E-mail:** Envia um e-mail de status ao final da execução (sucesso ou erro) através da integração com o Microsoft Outlook.

## Demonstração da Interface

A interface da aplicação é composta pelos seguintes elementos:

1.  **Campos de Entrada:** Para o e-mail de notificação, usuário e senha do Microsiga.
2.  **Modelo do Arquivo:** Uma imagem exibe o layout esperado da planilha Excel para garantir o formato correto dos dados.
3.  **Seleção de Arquivo:** Um botão para localizar e carregar o arquivo `.xlsx` com as previsões.
4.  **Botões de Ação:** "Cadastrar Previsões" para iniciar o robô e "Parar Execução" para interrompê-lo.
5.  **Barra de Status:** Uma label de texto que informa o andamento da automação em tempo real.

## Tecnologias Utilizadas

- **Python 3.x**
- **Selenium:** Para a automação do navegador web.
- **CustomTkinter:** Para a criação da interface gráfica.
- **Pandas:** Para leitura e manipulação de dados do arquivo Excel.
- **OpenPyXL:** Dependência do Pandas para ler arquivos `.xlsx`.
- **Pillow (PIL):** Para exibir a imagem do modelo na interface.
- **pywin32:** Para a integração com o Microsoft Outlook e envio de e-mails.

## Pré-requisitos

Antes de executar o projeto, certifique-se de que você tem instalado:

- [Python 3.8+](https://www.python.org/downloads/)
- Navegador [Google Chrome](https://www.google.com/chrome/) atualizado.
- Microsoft Outlook (versão Desktop) instalado e configurado na máquina.

## Instalação e Configuração

Siga os passos abaixo para configurar o ambiente de desenvolvimento:

1.  **Clone o repositório:**
    ```bash
    git clone [https://docs.github.com/articles/referencing-and-citing-content](https://docs.github.com/articles/referencing-and-citing-content)
    cd [nome-da-pasta-do-projeto]
    ```

2.  **Crie e ative um ambiente virtual (Recomendado):**
    ```bash
    # Criar o ambiente
    python -m venv venv

    # Ativar no Windows
    .\venv\Scripts\activate

    # Ativar no macOS/Linux
    source venv/bin/activate
    ```

3.  **Instale as dependências:**
    Copie as bibliotecas abaixo e instale todas com um único comando.
    ```bash
    pip install selenium customtkinter pandas openpyxl pillow pywin32
    ```

## Como Usar

1.  **Prepare a Planilha:** Certifique-se de que seu arquivo `.xlsx` com as previsões de vendas está formatado conforme o modelo exibido na imagem da aplicação.
2.  **Execute a Aplicação:**
    ```bash
    python main.py
    ```
3.  **Preencha os Campos:** Insira o e-mail para notificação, seu usuário e senha do Microsiga.
4.  **Selecione o Arquivo:** Clique em "Buscar Arquivo" e escolha a planilha preparada.
5.  **Inicie a Automação:** Clique em "Cadastrar Previsões".
6.  **Acompanhe o Status:** A label de status na parte inferior da janela informará o progresso. Os logs detalhados também aparecerão no terminal.
7.  **Pare a Execução (se necessário):** O botão "Parar Execução" pode ser usado para interromper o processo de forma segura.
8.  **Verifique seu E-mail:** Ao final, um e-mail de sucesso ou de erro será enviado para o endereço fornecido.

## Estrutura do Projeto

.
├── venv/                     # Pasta do ambiente virtual
├── img/                      # Pasta para imagens da UI
│   └── modelo_tab_prev_vend.png
├── main.py                   # Arquivo principal, responsável pela interface (UI)
├── acesso_microsiga.py       # Contém toda a lógica de automação (Selenium)
└── README.md                 # Este arquivo


## Detalhes Técnicos

O desafio técnico principal deste projeto é a interação com a interface WebApp do TOTVS Protheus, que utiliza extensivamente **Shadow DOM**. Para superar isso, foram implementadas duas funções auxiliares em `acesso_microsiga.py`:

- `_find_first_element()`: Utiliza JavaScript (`querySelector`) injetado para navegar pela árvore do Shadow DOM e retornar o **primeiro** elemento encontrado.
- `_find_all_elements()`: Utiliza JavaScript (`querySelectorAll`) para retornar uma **lista** de todos os elementos que correspondem a um seletor, permitindo a iteração sobre menus e listas de resultados.
