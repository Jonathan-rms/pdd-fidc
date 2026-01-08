from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time

# Configurações para rodar no servidor (Headless = sem interface gráfica)
chrome_options = Options()
chrome_options.add_argument("--headless") 
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

# Inicia o navegador
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

url = "https://calculadora-pdd-fidc.streamlit.app/"

print(f"Acessando {url}...")
driver.get(url)

# O PULO DO GATO: Esperar o JavaScript carregar e o app "bootar"
# Streamlit pode demorar para sair da hibernação
print("Aguardando carregamento (30s)...")
time.sleep(30)

# Tira um print (opcional, ajuda a debugar se der erro nos logs do Action)
print("Título da página encontrada:", driver.title)

driver.quit()
print("Processo finalizado com sucesso.")
