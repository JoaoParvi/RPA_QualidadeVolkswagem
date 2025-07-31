import time
import pandas as pd
import urllib
import logging
from io import StringIO
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
from selenium.webdriver.common.action_chains import ActionChains
from sqlalchemy import create_engine
from datetime import date

# Configurar o serviço do WebDriver
navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# Definir a URL e as credenciais
url = 'https://www.portalredevw.com.br/portalredevw2/menu_frames.aspx'
login = "01390992470"
senha = "2835"
TABELA_SQL = "QualidadeVolkswagen_PosVendas"

# Acessar a URL
navegador.get(url)
navegador.maximize_window()

time.sleep(5)

# Acessar a URL
navegador.get(url)
wait = WebDriverWait(navegador, 20)

# Esperar o primeiro iframe 'main' e entrar nele
wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "main")))

# Dentro dele, entrar no iframe 'mainApp'
wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "mainApp")))

# Agora, o conteúdo carregado é um <frameset>, precisamos entrar no frame 'HeaderMenu'
wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "HeaderMenu")))

# Agora estamos no frame correto, preencher login
campo_login = wait.until(EC.presence_of_element_located((By.ID, "txtCPFCNPJ")))
campo_login.send_keys(login)
time.sleep(6)

campo_senha = wait.until(EC.presence_of_element_located((By.ID, "txtSenha")))
campo_senha.send_keys(senha)
campo_senha.send_keys(Keys.ENTER)
time.sleep(6)

navegador.switch_to.default_content()

wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "main")))

wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "mainApp")))

wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "App")))

campo_filtroBre =  wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#rblEmpresa > tbody > tr:nth-child(1) > td:nth-child(1) > label")))
campo_filtroBre.click()
time.sleep(3)

navegador.switch_to.default_content()

wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "divframe")))

campo_Prog =  wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#rptMenuPrincipal_ctl06_hlkItemMenu")))
campo_Prog.click()
time.sleep(3)

campo_CEM =  wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#sub_77 > table > tbody > tr > td:nth-child(1) > h1:nth-child(8) > a")))
campo_CEM.click()
time.sleep(3)

navegador.get('https://satisfacaovw.com.br/cem/Previa/')
time.sleep(10)

navegador.switch_to.active_element.send_keys(Keys.ENTER)
time.sleep(3)

campo_area = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#Area")))
campo_area.click()

campo_posvenda = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#Area > option:nth-child(2)")))
campo_posvenda.click()

grupo_filtro = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "input.nivel:nth-child(6)")))
grupo_filtro.click()

input_filtroAuxiliar = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".chosen-choices")))
input_filtroAuxiliar.click()

navegador.switch_to.active_element.send_keys("93")
time.sleep(3)

navegador.switch_to.active_element.send_keys(Keys.ENTER)
time.sleep(5) 

navegador.switch_to.active_element.send_keys(Keys.ENTER)

campo_filtrar = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#filtrar")))
campo_filtrar.click()
time.sleep(3)



tabela_elemento = WebDriverWait(navegador, 20).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, ".datagrid-btable"))
)


# Extrai o HTML da tabela
tabela_html = tabela_elemento.get_attribute('outerHTML')

campo_passarpage = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".pagination-next")))
campo_passarpage.click()
time.sleep(3)



tabela_elemento1 = WebDriverWait(navegador, 20).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, ".datagrid-btable"))
)

# Extrai o HTML da tabela
tabela_html1 = tabela_elemento1.get_attribute('outerHTML')

df_table = pd.read_html(StringIO(tabela_html))[0]
df_table1 = pd.read_html(StringIO(tabela_html1))[0]

df_tabela = pd.concat([df_table, df_table1], ignore_index=True)

# Renomeando as colunas
df_tabela.rename(columns={
    0: 'Concessionaria',
    1: 'Realizados_Atual',
    2: 'Indice_Atual',
    3: 'Realizados_Anterior',
    4: 'Indice_Anterior',
    5: 'Realizados_Retrasado',
    6: 'Indice_Retrasado',
    7: 'Realizados_Media_Trimestral',
    8: 'Indice__Media_Trimestral',
    9: 'Realizados_Media_12meses',
    10: 'Indice_Media_12meses'
}, inplace=True)


# Apagando linhas em branco
df_tabela = df_tabela.drop(1)
df_tabela = df_tabela.drop(2)
df_tabela = df_tabela.drop(10)
df_tabela = df_tabela.drop(11)

substituicoes = {
    '84 - BREMEN': 'Bremen Bequimao',
    '1488 - BREMEN': 'Bremen Afogados',
    '97 - BREMEN': 'Bremen Salvador Shopping',
    '50 - BREMEN': 'Bremen Olinda',
    '18 - BREMEN CARUARU': 'Bremen Caruaru',
    '1188 - BREMEN RECIFE': 'Bremen Recife',
    '1214 - BREMEN SALVADOR': 'Bremen Salvador',
    '1316 - BREMEN FEIRA DE SANTANA': 'Bremen Feira de Santana',
    '1346 - BREMEN SÃO LUIS': 'Bremen São Luis',
    '1359 - BREMEN TIRIRICAL': 'Bremen Tirirical'
}

# Substituindo os valores na coluna 'Concessionaria'
df_tabela['Concessionaria'] = df_tabela['Concessionaria'].replace(substituicoes)

df3 = pd.DataFrame({"Segmento": ["Pos Vendas"]})
df = pd.concat([df_tabela.reset_index(drop=True), df3.reset_index(drop=True)], axis=1)
df["Segmento"] = df["Segmento"].fillna("Pos Vendas")

# Adicionando a coluna "data_atualizacao" com a data atual
df['data_atualizacao'] = date.today()

print (df)

# Fechando o navegador após 10 segundos
time.sleep(5) 
navegador.quit()
print("Fechando navegador...")


 # === SALVA NO BANCO DE DADOS ===
try:
    print("Conectando ao banco de dados...")
    user = 'rpa_bi'
    password = 'Rp@_B&_P@rvi'
    host = '10.0.10.243'
    port = '54949'
    database = 'stage'

    params = urllib.parse.quote_plus(
        f'DRIVER=ODBC Driver 17 for SQL Server;SERVER={host},{port};DATABASE={database};UID={user};PWD={password}'
    )
    connection_str = f'mssql+pyodbc:///?odbc_connect={params}'
    engine = create_engine(connection_str)

    with engine.connect() as connection:
        df.to_sql(TABELA_SQL, con=connection, if_exists='replace', index=False)

    print(f"Dados inseridos com sucesso na tabela '{TABELA_SQL}'!")
    logging.info(f"Dados inseridos com sucesso na tabela '{TABELA_SQL}'.")

except Exception as e:
    logging.exception("Erro ao inserir dados no banco: %s", str(e))
    print("Erro ao inserir dados no banco:", str(e))
