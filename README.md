# 093-e-147-50mais
### Um projeto de automação em python,  feito para ajudar a acelerar o processo de aquisição de relatórios mensais todo dia do site da empresa 50 mais.

## O Que o código Faz?

1.  Abre o Google Chrome e navega para o sistema.
2.  Realiza o login com usuário e senha.
3.  Navega para a tela de relatórios (atalho 093 e 147).
4.  Seleciona filtros (datas do mês corrente).
5.  Clica em "Pesquisar" e depois em "Excel" para baixar o relatório.
6.  Identifica o arquivo baixado (que é um HTML disfarçado de `.xls`).
7.  Cria um novo arquivo `.xls` limpo, com a planilha interna renomeada para "BD" e com formatação de números e datas.
8.  O arquivo final é salvo na pasta `downloads_excel` com o nome `Relatorio_Mensal_AAAA-MM_atalho_147.xls` e `Relatorio_Mensal_AAAA-MM_atalho_093.xls`, sobrepondo caso ja tenha um arquivo nesse mês.

## Execução

- **Manualmente:** Dê um duplo clique no arquivo `executar_automacao_93_147.bat`.
- **Automaticamente:** Configure o Agendador de Tarefas do Windows para executar o arquivo `executar_automacao_93_147.bat` diariamente no horário desejado.

## Requisitos

- **Software:**
  - Python 3.11+
  - Google Chrome

- **Bibliotecas Python:**
  - selenium
  - webdriver-manager
  - pandas
  - lxml
  - xlrd
  - xlwt

## Configuração/Como usar

1.  **Clone este repositório:**
    ```sh
    git clone https://github.com/Wheremst/093-e-147-50mais
    ```
2.  **Instale as dependências:**
    ```sh
    pip install -r requirements.txt
    ```
3.  **Credenciais:**
    - Abra o arquivo `.py` e altere as variáveis `USUARIO` e `SENHA` com suas credenciais reais.
    - Também altere o caminho que deseja para o download dos arquivos excel
4.  **Agendamento:**
    - O arquivo `executar_automacao_93_147` é usado para agendar a execução. **É necessário editar este arquivo** e atualizar os caminhos para o seu `python.exe` e para a pasta do script, conforme o seu computador.

