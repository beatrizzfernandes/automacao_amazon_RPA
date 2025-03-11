from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By # MÃ³dulo para localizar elementos na pÃ¡gina
from selenium.webdriver.support.ui import WebDriverWait # Esperar condiÃ§Ãµes especÃ­ficas na pÃ¡gina
from selenium.webdriver.support import expected_conditions as EC # Esperar condiÃ§Ãµes especÃ­ficas na pÃ¡gina
import time # Pausas na automaÃ§Ã£o
from openpyxl import Workbook # Criar e manipular arquivos Excel
import smtplib # Enviar emails
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import json # Formato Json 
import logging  # Para registrar logs
import os # Interagir com o sistema operacional

print(os.getcwd()) # Exibir o diretório de trabalho atual

# ConfiguraÃ§Ã£o do logging
logging.basicConfig(
    level=logging.INFO,  # Define o nÃ­vel mÃ­nimo de log (INFO ou superior)
    format='%(asctime)s - %(levelname)s - %(message)s',  # Formato do log
    filename='automacao.log',  # Nome do arquivo de log
    filemode='a'  # Modo de abertura do arquivo (Adiciona novas mensagens sem apagar as antigas)
)

# Arquivo conf.json
with open("conf.json", "r") as f: # Abre o arquivo
    config = json.load(f) # Carrega o conteÃºdo em um dicionÃ¡rio Python

# ConfiguraÃ§Ãµes
remetente = config["email"]["remetente"]
senha = config["email"]["senha"]
destinatario = config["email"]["destinatario"]
termo_pesquisa = config["amazon"]["pesquisa"]
arquivo_excel = config["arquivo_excel"]

# FunÃ§Ã£o para enviar e-mail com anexo
def enviar_email(arquivo_anexo, destinatario, remetente, senha):
    assunto = "RelatÃ³rio de Produtos da Amazon"
    msg = MIMEMultipart() # Cria uma mensagem de email
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = assunto

    with open(arquivo_anexo, "rb") as anexo:
        part = MIMEBase('application', 'octet-stream') # Define o tipo de conteÃºdo do anexo
        part.set_payload(anexo.read())
        encoders.encode_base64(part) # Codifica o anexo em Base64
        part.add_header('Content-Disposition', f'attachment; filename="{arquivo_anexo}"')
        msg.attach(part)

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)  # Servidor SMTP do Gmail
        server.starttls() # Habilita criptografia TLS
        server.login(remetente, senha)
        server.sendmail(remetente, destinatario, msg.as_string())
        server.quit()
        logging.info("â E-mail enviado com sucesso!")
    except Exception as e:
        logging.error(f"â Erro ao enviar e-mail: {e}")

# Inicializa o WebDriver
logging.info("Iniciando o WebDriver...")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# Acessa a Amazon
logging.info("Acessando a Amazon...")
driver.get("https://www.amazon.com.br")

# Espera o CAPTCHA manual
while True:
    try:
        # Verifica se a barra de pesquisa estÃ¡ visÃ­vel 
        WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.ID, "twotabsearchtextbox"))
        )
        logging.info("â CAPTCHA resolvido! Continuando a automaÃ§Ã£o...")
        break  # Sai do loop
    except:
        logging.info("â³ Aguardando CAPTCHA ser resolvido...")
        time.sleep(3)  # Espera antes de tentar de novo

# Aguarda a barra de pesquisa e digita o produto
search_box = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "twotabsearchtextbox"))
)
search_box.click()
search_box.clear()
search_box.send_keys(termo_pesquisa)  # Usa o termo de pesquisa do conf.json

# Aguarda o botÃ£o de pesquisa e clica nele
search_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, "nav-search-submit-button"))
)
search_button.click()

# Aguarda os resultados carregarem
time.sleep(5) 

# Coleta os dados dos 10 primeiros produtos
products = []

# Localiza os elementos dos produtos
product_elements = driver.find_elements(By.CSS_SELECTOR, "div.s-main-slot div[data-component-type='s-search-result']")

for product in product_elements[:10]:  # Limita aos 10 primeiros produtos
    try:
        # Nome do produto
        name = product.find_element(By.CSS_SELECTOR, "h2 span").text
    except:
        name = "N/A"

    try:
        # PreÃ§o do produto
        price = product.find_element(By.CSS_SELECTOR, ".a-price .a-offscreen").get_attribute("innerHTML")
        # Remove caracteres indesejados (como &nbsp;)
        price = price.replace("&nbsp;", " ").strip()
    except:
        price = "N/A"

    try:
        # AvaliaÃ§Ã£o do produto
        rating = product.find_element(By.CSS_SELECTOR, "i.a-icon-star-small span.a-icon-alt").get_attribute("innerHTML")
    except:
        rating = "N/A"

    try:
        # NÃºmero de avaliaÃ§Ãµes
        num_reviews = product.find_element(By.CSS_SELECTOR, "span.a-size-base.s-underline-text").text
    except:
        num_reviews = "N/A"

    # Adiciona os dados Ã  lista de produtos
    products.append([name, price, rating, num_reviews])
# Fecha o navegador
logging.info("Fechando o navegador...")
driver.quit()

# Cria um arquivo Excel com openpyxl
wb = Workbook()  # Cria uma nova workbook
ws = wb.active  # Seleciona a planilha ativa

# Adiciona cabeÃ§alhos
ws.append(["Nome", "PreÃ§o", "AvaliaÃ§Ã£o", "NÃºmero de AvaliaÃ§Ãµes"])

# Adiciona os dados dos produtos
for product in products:
    ws.append(product)

# Salva o arquivo Excel
wb.save(arquivo_excel)
logging.info(f"â Dados salvos com sucesso em '{arquivo_excel}'!")

# Envia o e-mail com o relatÃ³rio em anexo
logging.info("Enviando e-mail com o relatÃ³rio...")
enviar_email(arquivo_excel, destinatario, remetente, senha)  