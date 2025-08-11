# ==============================================================================
#      De Bruno Rabelo ou Wheremst (https://github.com/Wheremst)
# ==============================================================================

#Automatizar nosso navegador
from selenium import webdriver
# Gerenciador automático de drivers
from webdriver_manager.chrome import ChromeDriverManager 
# Responsável por iniciar e parar o arquivo executável do driver
from selenium.webdriver.chrome.service import Service
# Fornece as diferentes maneiras de encontrar elementos em uma página da web
from selenium.webdriver.common.by import By
# Permite o uso de teclas especiais (Enter, Tab, Shift, etc.)
from selenium.webdriver.common.keys import Keys
# Permite que o programa espere um botão ser capaz de ser clicado
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# Pausas forçadas no nosso codigo
import time
# Módulos para manipular pastas e arquivos do sistema
import os
#Manipulação de datas
import datetime
import calendar
#Tradução de datas
import locale
# Manipulação do excel
import pandas as pd
import xlwt
import xlwings as xw

# Configura a localidade para Português do Brasil para que o nome do mês seja traduzido
# Em alguns sistemas, pode ser necessário usar 'Portuguese' em vez de 'pt_BR.UTF-8'
try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
except locale.Error:
    print("Localidade 'pt_BR.UTF-8' não encontrada, tentando 'Portuguese'.")
    locale.setlocale(locale.LC_TIME, 'Portuguese')

hoje = datetime.date.today()
nome_mes_atual = hoje.strftime('%B').capitalize()
ano_atual = hoje.year
numero_mes_atual = hoje.month
mes_ano_atual = hoje.strftime('%Y-%m')

# Usa a função monthrange para encontrar o número de dias no mês atual
# Retorna uma tupla: (dia da semana do dia 1, número de dias no mês)
# Pegamos o segundo valor com o índice [1]
dias_no_mes = calendar.monthrange(ano_atual, numero_mes_atual)[1]

chrome_options = webdriver.ChromeOptions()

#Configurar o navegador para lidar com downloads de arquivos
# --- PASSO 1: CONFIGURAR AS OPÇÕES DO CHROME ---

# Definir o caminho da pasta onde os downloads serão salvos
diretorio_desejado = r"CAMINHO/ATÉ/A/PASTA"

# 2. Defina o nome da pasta de destino
pasta_destino = "NOME DA PASTA"
caminho_download_absoluto = os.path.abspath(os.path.join(diretorio_desejado, pasta_destino))

# Criar a pasta se ela não existir
if not os.path.exists(caminho_download_absoluto):
    os.makedirs(caminho_download_absoluto)

# Criar o dicionário de preferências usando o caminho absoluto
prefs = {
    # Use o caminho absoluto na configuração
    "download.default_directory": caminho_download_absoluto, 
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True
}

# Aplicar as preferências
chrome_options.add_experimental_option("prefs", prefs)


# --- PASSO 2: INICIAR O NAVEGADOR COM AS OPÇÕES ---
navegador = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()), 
    options=chrome_options  # Certifique-se de que este argumento está aqui
)

time.sleep(2)
# Ir ate o site
navegador.get("https://50maiscs.brudam.com.br/index.php")

# Preencher o campo de USUARIO e SENHA
campo_usuario = WebDriverWait(navegador, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="user"]'))
)
campo_usuario.send_keys("USUARIO")#Trocar as credenciais para funcionar

campo_senha = WebDriverWait(navegador, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="password"]'))
)
campo_senha.send_keys("SENHA") #Trocar as credenciais para funcionar

botao_acessar = WebDriverWait(navegador, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="acessar"]'))
)
botao_acessar.click()

# Espera o campo de código do atalho aparecer
campo_codigo_opcao = WebDriverWait(navegador, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="codigo_opcao"]'))
)
campo_codigo_opcao.send_keys("093")#digita "093"
campo_codigo_opcao.send_keys(Keys.RETURN)#pressiona Enter

# --- OPÇÕES PARA PUXAR O RELATÓRIO CERTO ---
unidade = WebDriverWait(navegador, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="unidade"]'))
)
unidade.click()
unidade.send_keys(Keys.ARROW_UP * 4, Keys.ENTER)

unidade_CTE = WebDriverWait(navegador, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="unidade_cte"]'))
)
unidade_CTE.click()
unidade_CTE.send_keys(Keys.ARROW_UP * 4, Keys.ENTER)

# Variáveis de data
data_inicial = hoje.strftime('01%m%Y')
data_final = hoje.strftime(f'{dias_no_mes}%m%Y')

# Espera o campo de data inicial ficar clicável, limpa e envia a nova data
campo_data_inicial = WebDriverWait(navegador, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="data_1"]'))
)
campo_data_inicial.send_keys(Keys.BACK_SPACE * 8)
campo_data_inicial.send_keys(data_inicial)

# Espera o campo de data final ficar clicável, limpa e envia a nova data
campo_data_final = WebDriverWait(navegador, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="data_2"]'))
)
campo_data_final.send_keys(Keys.BACK_SPACE * 8)
campo_data_final.send_keys(data_final)

#Pausa forçada para ter certeza
time.sleep(3)

navegador.find_element(By.XPATH, '//*[@id="PESQUISAR"]').click()
#Pausa forçada para ter certeza
time.sleep(3)

# ==============================================================================
#      INÍCIO DO BLOCO DE DOWNLOAD E MANIPULAÇÃO DE ARQUIVO
# ==============================================================================
try:
    # 1. PEGAR A "FOTO" DA PASTA ANTES DO DOWNLOAD
    arquivos_antes = set(os.listdir(caminho_download_absoluto))
    print(f"Arquivos na pasta antes do download: {arquivos_antes or 'Nenhum'}")

    # 2. CLICAR NO BOTÃO DE DOWNLOAD
    download_excel = WebDriverWait(navegador, 10).until(
    EC.element_to_be_clickable((By.XPATH, '/html/body/center/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td[3]/img'))
    )
    download_excel.click()
    
    # 2. ESPERAR O DOWNLOAD TERMINAR
    print("Aguardando o download ser concluído...")
    tempo_max_espera = 90
    nome_original_arquivo = None
    for _ in range(tempo_max_espera):
        time.sleep(1)
        arquivos_depois = set(os.listdir(caminho_download_absoluto))
        novos_arquivos = arquivos_depois - arquivos_antes
        if novos_arquivos and not any(f.endswith('.crdownload') for f in novos_arquivos):
            nome_original_arquivo = novos_arquivos.pop()
            print(f"Download concluído! Arquivo original: {nome_original_arquivo}")
            break
    
    if not nome_original_arquivo:
        raise Exception("O download não foi concluído no tempo esperado.")

    # 3. DEFINIR NOMES E CAMINHOS FINAIS
    nome_arquivo_final = f"Relatorio_Mensal_{mes_ano_atual}_atalho_093.xls"
    nome_planilha_final = "BD"
    caminho_original = os.path.join(caminho_download_absoluto, nome_original_arquivo)
    caminho_final = os.path.join(caminho_download_absoluto, nome_arquivo_final)

    # 4. LER O ARQUIVO COM PANDAS
    print(f"Lendo dados de '{nome_original_arquivo}' como HTML...")
    lista_de_tabelas = pd.read_html(
        caminho_original, 
        encoding='utf-8',
        decimal=',',      # Diz que a vírgula é o separador decimal
        thousands='.'     # Diz que o ponto é o separador de milhar
    )
    dataframe = lista_de_tabelas[0]
    if dataframe.empty:
        raise ValueError("O arquivo foi lido, mas não continha dados.")

    # =================================================================
    #          ETAPA DE HIGIENIZAÇÃO DOS DADOS
    # =================================================================
    print("Higienizando os dados para garantir compatibilidade com .xls...")
    
    # Cria uma cópia do dataframe para trabalhar
    dataframe_higienizado = dataframe.copy()

    # Itera sobre cada coluna para tratar os tipos de dados
    for coluna in dataframe_higienizado.columns:
        # Verifica se a coluna é do tipo data/hora
        if pd.api.types.is_datetime64_any_dtype(dataframe_higienizado[coluna]):
            # Converte as datas para formato de texto (string)
            # e preenche valores de data vazios (NaT) com texto vazio ''
            dataframe_higienizado[coluna] = dataframe_higienizado[coluna].apply(
                lambda x: x.strftime('%d/%m/%Y %H:%M:%S') if pd.notna(x) else ''
            )
            
    # Converte quaisquer outros valores nulos/NaN restantes para texto vazio
    dataframe_higienizado = dataframe_higienizado.fillna('')
    # =================================================================

    print("Dados higienizados. Iniciando escrita manual...")
    
    # 5. CRIAR UMA PASTA DE TRABALHO .xls EM BRANCO
    workbook_wt = xlwt.Workbook(encoding='utf-8')
    planilha_wt = workbook_wt.add_sheet(nome_planilha_final)
    
    # 6. ESCREVER CABEÇALHOS
    for col_num, valor in enumerate(dataframe_higienizado.columns):
        planilha_wt.write(0, col_num, str(valor))
        
    # 7. ESCREVER OS DADOS HIGIENIZADOS
    for row_num, row in enumerate(dataframe_higienizado.values.tolist(), 1):
        for col_num, cell_value in enumerate(row):
            planilha_wt.write(row_num, col_num, str(cell_value))

    # 8. SOBREPOR E SALVAR
    if os.path.exists(caminho_final):
        os.remove(caminho_final)
    workbook_wt.save(caminho_final)
    
    # 9. LIMPEZA
    os.remove(caminho_original)
    
    print(f"\nPROCESSO CONCLUÍDO COM SUCESSO!")
    print(f"Arquivo final salvo em: {caminho_final}")

except Exception as e:
    print(f"\nOCORREU UM ERRO NO PROCESSO FINAL: {e}")

finally:
    print("\nFechando o navegador.")
    time.sleep(3)
    navegador.quit()

