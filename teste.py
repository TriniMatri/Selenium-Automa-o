import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

# Caminho do driver do Edge
service = Service(r"")
options = Options()

# Variáveis de dados
CODFORN = "123" #COLUNA 1 A
nome_form = "WALT TRINITI" #COLUNA 2 B
NOTA = "1213" #COLUNA 3 C
NUMPED = "123456789" #COLUNA 4 D
VALOR = "2039" #COLUNA 5 E
vencimento = "22/03/2025" #COLUNA 7 G
DATAENVIO = "" #COLUNA 11 L
TIPOPAG = "BOLETO" #COLUNA 12 M
NUMENTRA = "1010101010" #COLUNA 13 N
DATAEMI = "18/03/2025" #COLUNA 17 R

login = ""
senha = ""



diretorio = r""  # Ajustar diretório para teste
nome_arquivo = f"03.05 NF {NOTA} F {CODFORN} {nome_form}.pdf"
file_path = os.path.join(diretorio, nome_arquivo)

# Abrindo o Forms
driver = webdriver.Edge(service=service, options=options)
options.add_argument('--disable-extensions')
driver.get("")

# Entrar na conta para conseguir acesso ao Forms
login_field = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.XPATH, '/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[3]/div/div/div/div[2]/div[2]/div/input[1]'))
)
login_field.click()
login_field.send_keys(login)

loginbutton = WebDriverWait(driver, 5).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="idSIButton9"]'))
)
loginbutton.click()
time.sleep(30)

# Preencher o campo de senha
senha_field = WebDriverWait(driver, 10).until(
    EC.visibility_of_element_located((By.XPATH, '//*[@id="passwordInput"]'))
)
ActionChains(driver).move_to_element(senha_field).click().perform()
senha_field.send_keys(senha)

# Submeter a senha
submit_button = WebDriverWait(driver,5).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="submitButton"]'))
)
submit_button.click()


opcaosetor = WebDriverWait(driver, 5).until(
    EC.presence_of_element_located((By.XPATH, "//input[@value='']"))
)
opcaosetor.click()
time.sleep(2)

# SETOR
menu_suspenso = driver.find_element(By.XPATH, '//*[@id="rc78bb8d842574508936c58a1235d1db9_placeholder_content"]')
menu_suspenso.click()
opcao_ti = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//span[text()='TI']"))
)
opcao_ti.click()
time.sleep(2)

# Tipo PEDIDO
menu_uso_consumo = driver.find_element(By.XPATH, '//*[@id="r5cb66de3650f44d28e7b0b1e880a851a_placeholder_content"]')
driver.execute_script("arguments[0].scrollIntoView(true);", menu_uso_consumo)
menu_uso_consumo = WebDriverWait(driver, 5).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="r5cb66de3650f44d28e7b0b1e880a851a_placeholder_content"]'))
)
menu_uso_consumo.click()

opcao_uso_consumo = WebDriverWait(driver, 5).until(
    EC.element_to_be_clickable((By.XPATH, "//span[text()='USO E CONSUMO']"))
)
opcao_uso_consumo.click()
time.sleep(2)

# Preenchendo os campos
campo_data = {
    "NUMPED": "//*[@id='question-list']/div[4]/div[2]/div/span/input",
    "NUMENTRA": "//*[@id='question-list']/div[5]/div[2]/div/span/input",
    "NOTA": "//*[@id='question-list']/div[6]/div[2]/div/span/input",
    "CODFORN": "//*[@id='question-list']/div[7]/div[2]/div/span/input",
    "VALOR": "//*[@id='question-list']/div[8]/div[2]/div/span/input",
    "DATAEMI": "//*[@id='question-list']/div[9]/div[2]/div/span/input"
}

# Preencher os campos automaticamente
for key, xpath in campo_data.items():
    campo_texto = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, xpath))
    )
    campo_texto.send_keys(locals()[key])  # Usa a variável dinamicamente
    time.sleep(2)

# Selecionando o tipo de pagamento
menu_TIPAG = driver.find_element(By.XPATH, '//*[@id="re46b7db9d675498babd619893a9dadec_placeholder_content"]')
menu_TIPAG.click()

opcaoTIPAG = WebDriverWait(driver, 5).until(
    EC.element_to_be_clickable((By.XPATH, f"//span[text()='{TIPOPAG}']"))
)

if TIPOPAG == "BOLETO":
    # Selecionando vencimento
    opcaoTIPAG.click()
    menu_vencimento = driver.find_element(By.XPATH, '//*[@id="rfdc43eac78b74c17b4aee0f6ce4c81f8_placeholder_content"]')
    menu_vencimento.click()

    opcaovenci = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, f"//span[text()='{vencimento}']"))
    )
    opcaovenci.click()

    # Upload do arquivo
    upload_element = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="question-list"]/div[12]/div[2]/div/button'))
    )
    input_file = driver.find_element(By.XPATH, '//*[@id="question-list"]/div[12]/div[2]/div/input')
    input_file.send_keys(file_path)
    time.sleep(5)

    # Enviar o formulário
    enviarbutton = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="form-main-content1"]/div/div/div[2]/div[3]/div/button'))
    )
    enviarbutton.click()
    time.sleep(3)

elif TIPOPAG == "DEPOSITO":  # Verificar os campos do Depósito
    opcaoTIPAG.click()

# Espera final para garantir que tudo foi enviado
time.sleep(3)
