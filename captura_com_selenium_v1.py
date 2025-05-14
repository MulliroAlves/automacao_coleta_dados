from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import re
import timeit

# Gspread e autenticação
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Configuração do acesso ao Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(
    r"C:\Users\MDT Store\Desktop\api-planilha-cnpja-f4157f5ba0e2.json", scope)
client = gspread.authorize(creds)

# Abra a planilha (já precisa existir e estar compartilhada com a conta de serviço)
sheet = client.open("Planilha-CNPJA").sheet1

# Escreve o cabeçalho
sheet.append_row([
    "CNPJ", "Nome", "Status", "Período de Operação", "Nome Fantasia",
    "Endereço", "Natureza Jurídica", "Capital Social", "E-mail", "Telefone", "URL"
])

# Configura o Selenium
service = Service(r"C:\Program Files\Selenium\chromedriver.exe")
options = webdriver.ChromeOptions()
options.add_argument("window-size=1920,1080")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--no-sandbox")
options.add_argument("--log-level=3")

# Medir tempo de abertura
start = timeit.default_timer()
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 10)
print("⏱️ Navegador iniciado em:", round(timeit.default_timer() - start, 2), "segundos")

# Lista de CNPJs
cnpjs = ["47703730000112", "09117642000140", "21848698000170", "49629113000140", "40188631000109", "22802772000180", "30596100000193"]

def get_button_data(button_id):
    try:
        return driver.find_element(By.XPATH, f"//button[@id='{button_id}']/div/div").text.strip()
    except:
        return "Não encontrado"

def fallback_via_bs4():
    soup = BeautifulSoup(driver.page_source, "html.parser")
    texts = [elem.get_text(strip=True) for elem in soup.find_all(string=True)]

    cep_regex = re.compile(r'\d{5}-\d{3}$')
    telefone_regex = re.compile(r'\(?\d{2}\)?\s?\d{4,5}-\d{4}')
    email_regex = re.compile(r'\b[\w\.-]+@[\w\.-]+\.\w{2,}\b')
    capital_regex = re.compile(r'R\$\s?[\d\.\,]+')
    cnpj_regex = re.compile(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}')
    periodo_regex = re.compile(r'\d{2}/\d{2}/\d{4}\s*-\s*(\d{2}/\d{2}/\d{4}|Presente)', re.IGNORECASE)

    endereco = next((t for t in texts if cep_regex.search(t)), "Não encontrado")
    telefone = next((t for t in texts if telefone_regex.search(t)), "Não encontrado")
    email = next((t for t in texts if email_regex.search(t)), "Não encontrado")
    capital = next((t for t in texts if capital_regex.search(t)), "Não encontrado")
    cnpj = next((t for t in texts if cnpj_regex.search(t)), "Não encontrado")
    periodo = next((t for t in texts if periodo_regex.search(t)), "Não encontrado")

    return endereco, capital, email, telefone, cnpj, periodo

# Loop principal
for cnpj_num in cnpjs:
    print(f"\n🚀 Acessando CNPJ: {cnpj_num}...")
    url = f"https://cnpja.com/office/{cnpj_num}"
    driver.get(url)

    # Espera até o conteúdo principal aparecer
    try:
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "[data-toc-observe-for]")))
    except:
        print("⚠️ Timeout: página demorou a carregar")

    try:
        nome = driver.find_element(By.CSS_SELECTOR, "[data-toc-observe-for]").text.strip()
    except:
        nome = "Não encontrado"

    try:
        status = driver.find_element(
            By.XPATH,
            "//div[contains(@class, 'text-sm') and (contains(@class, 'text-green-600') or contains(@class, 'text-rose-400') or contains(@class, 'text-red-600'))]"
        ).text.strip()
    except:
        status = "Não encontrado"

    try:
        cnpj_formatado = driver.find_element(By.XPATH, "//dt[text()='CNPJ']/following-sibling::dd[1]").text.strip()
    except:
        cnpj_formatado = "Não encontrado"

    try:
        periodo_operacao = driver.find_element(
            By.XPATH, "//dt[text()='Período de Operação']/following-sibling::dd[1]"
        ).text.strip()
    except:
        periodo_operacao = "Não encontrado"

    nome_fantasia = get_button_data("bits-154")
    endereco = get_button_data("bits-143")
    natureza_juridica = get_button_data("bits-150")
    capital_social = get_button_data("bits-152")
    email = get_button_data("bits-157")
    telefone = get_button_data("bits-161")

    if "Não encontrado" in [endereco, capital_social, email, telefone]:
        print("⚠️  Usando BeautifulSoup para dados ausentes...")
        endereco, capital_social, email, telefone, cnpj_formatado, periodo_operacao = fallback_via_bs4()

    print("🔎 Dados coletados:")
    print(f"CNPJ: {cnpj_formatado}")
    print(f"Nome: {nome}")
    print(f"Status: {status}")
    print(f"Período de Operação: {periodo_operacao}")
    print(f"Nome Fantasia: {nome_fantasia}")
    print(f"Endereço: {endereco}")
    print(f"Natureza Jurídica: {natureza_juridica}")
    print(f"Capital Social: {capital_social}")
    print(f"E-mail: {email}")
    print(f"Telefone: {telefone}")

    sheet.append_row([
        cnpj_formatado, nome, status, periodo_operacao, nome_fantasia,
        endereco, natureza_juridica, capital_social, email, telefone, url
    ])

driver.quit()
