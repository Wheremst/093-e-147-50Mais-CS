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
# responsavel pela manipulação do excel
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

# Usa a função monthrange para encontrar o número de dias no mês atual
# Retorna uma tupla: (dia da semana do dia 1, número de dias no mês)
# Pegamos o segundo valor com o índice [1]
dias_no_mes = calendar.monthrange(ano_atual, numero_mes_atual)[1]

chrome_options = webdriver.ChromeOptions()

#Configurar o navegador para lidar com downloads de arquivos
# --- PASSO 1: CONFIGURAR AS OPÇÕES DO CHROME ---

# Definir o caminho da pasta onde os downloads serão salvos
diretorio_desejado = r"CAMINHO/ATÉ/A/PASTA"
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
    
    # 3. ESPERAR O DOWNLOAD TERMINAR
    print("Aguardando o download ser concluído...")
    tempo_max_espera = 90
    download_concluido = False
    nome_original_arquivo = None

    for _ in range(tempo_max_espera):
        time.sleep(1)
        # Verifica se não há mais arquivos temporários de download
        if not any(fname.endswith('.crdownload') for fname in os.listdir(caminho_download_absoluto)):
            arquivos_depois = set(os.listdir(caminho_download_absoluto))
            novos_arquivos = arquivos_depois - arquivos_antes
            if novos_arquivos:
                nome_original_arquivo = novos_arquivos.pop()
                download_concluido = True
                break
    
    if not download_concluido:
        raise Exception("O download não foi concluído no tempo esperado.")

    print(f"Download concluído! Nome original do arquivo: {nome_original_arquivo}")

    # 4. DEFINIR O NOME PADRÃO PARA O MÊS ATUAL
    mes_ano_atual = datetime.date.today().strftime('%Y-%m') # Formato AAAA-MM
    
    # Esse nome de arquivo é o que o programa vai procuarar nas proximas vezes que baixar o excel
    nome_arquivo_mensal = f"Relatorio_Mensal_{mes_ano_atual}_atalho_093.xls"
    print(f"O nome padrão para este mês é: {nome_arquivo_mensal}")

    # Constrói o caminho completo para o arquivo original e o de destino
    caminho_original = os.path.join(caminho_download_absoluto, nome_original_arquivo)
    caminho_destino_mensal = os.path.join(caminho_download_absoluto, nome_arquivo_mensal)

    # 5. LÓGICA DE SOBREPOSIÇÃO
    # Antes de renomear, verifica se o arquivo de destino do mês já existe
    if os.path.exists(caminho_destino_mensal):
        print(f"Arquivo '{nome_arquivo_mensal}' já existe. Removendo o antigo para sobrepor...")
        os.remove(caminho_destino_mensal) # Remove o arquivo antigo

    # 6. RENOMEAR O NOVO DOWNLOAD
    os.rename(caminho_original, caminho_destino_mensal)
    print(f"Arquivo renomeado com sucesso para '{nome_arquivo_mensal}'!")
except Exception as e:
    print(f"Ocorreu um erro no processo de download ou renomeação: {e}")
    
# Constrói o caminho completo para o arquivo Excel
caminho_completo_arquivo = os.path.join(caminho_download_absoluto, nome_arquivo_mensal)

# O novo nome que você deseja para a planilha interna
novo_nome_planilha = "BD" 

# Inicia uma instância do Excel invisível
app = xw.App(visible=False)

try:
    # DIZ AO EXCEL PARA NÃO MOSTRAR NENHUM ALERTA 
    app.display_alerts = False

    # Abre a pasta de trabalho (o arquivo .xls)
    workbook = app.books.open(caminho_completo_arquivo)

    # Pega a primeira planilha e a renomeia
    planilha = workbook.sheets[0]
    planilha.name = novo_nome_planilha

    # Salva e fecha o arquivo. O Excel vai "limpar" o arquivo ao salvar.
    workbook.save()
    workbook.close()
    
    print(f"Sucesso! Planilha renomeada para '{novo_nome_planilha}' usando xlwings.")

except Exception as e:
    print(f"Ocorreu um erro com xlwings: {e}")

finally:
    # Garante que o processo do Excel seja fechado no final
    app.quit()