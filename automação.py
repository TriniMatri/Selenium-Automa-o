import time
import os
from datetime import datetime
import win32com.client as win32
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

# Abrir o Excel
excel = win32.Dispatch("Excel.Application")
excel.Visible = False  # Rodar em segundo plano

# Caminho da planilha no OneDrive (ajuste para o caminho correto no seu PC)
caminho_planilha = r""
workbook = excel.Workbooks.Open(caminho_planilha)
sheet = workbook.Sheets("teste")  # Ajuste o nome da aba

# Identificar a última linha preenchida
ultima_linha = sheet.Cells(sheet.Rows.Count, 1).End(-4162).Row  

# Filtrar apenas as linhas onde STATUS = "ENVIAR AO PAGAMENTO"
dados = []
for i in range(2, ultima_linha + 1):  # Começa a leitura da linha 2 (ignorando cabeçalho)
    status = sheet.Cells(i, 7).Value  # STATUS está na 7ª coluna (coluna G)
    
    # Exibir o valor do status para depuração
    print(f"Status na linha {i}: '{status}'")  # Verifique o valor exato do status
    
    if status and isinstance(status, str) and status.strip().lower() == "ENVIAR PARA PAGAMENTO".lower():
        CODFORN = sheet.Cells(i, 1).Value  
        nome_form = sheet.Cells(i, 2).Value  
        NOTA = sheet.Cells(i, 3).Value  
        NUMPED = sheet.Cells(i, 4).Value  
        VALOR = sheet.Cells(i, 5).Value  
        vencimento = sheet.Cells(i, 8).Value  
        TIPOPAG = sheet.Cells(i, 13).Value
        NUMENTRA = sheet.Cells(i, 14).Value
        DATAEMI = sheet.Cells(i, 18).Value

        # Salvar os dados extraídos junto com a linha correspondente
        dados.append((CODFORN, nome_form, NOTA, NUMPED, VALOR, vencimento, TIPOPAG, NUMENTRA, DATAEMI, i))

# Fechar o Excel
workbook.Close()
excel.Quit()

# Configuração do Selenium
service = Service(r"")
options = Options()

# Abrindo o Forms
driver = webdriver.Edge(service=service, options=options)
options.add_argument('--disable-extensions')
driver.get("")

# Entrar na conta para conseguir acesso ao Forms
login = ""
senha = ""

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
senha_field = WebDriverWait(driver, 5).until(
    EC.visibility_of_element_located((By.XPATH, '//*[@id="passwordInput"]'))
)
ActionChains(driver).move_to_element(senha_field).click().perform()
senha_field.send_keys(senha)

# Submeter a senha
submit_button = WebDriverWait(driver, 5).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="submitButton"]'))
)
submit_button.click()

# Escolhendo BEMOL
opcaosetor = WebDriverWait(driver, 5).until(
    EC.presence_of_element_located((By.XPATH, "//input[@value='BEMOL S.A']"))
)
opcaosetor.click()
time.sleep(2)

# SETOR
menu_suspenso = driver.find_element(By.XPATH, '//*[@id="rc78bb8d842574508936c58a1235d1db9_placeholder_content"]')
menu_suspenso.click()
opcao_ti = WebDriverWait(driver, 3).until(
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

# Preenchendo os campos automaticamente para cada registro
for registro in dados:
    CODFORN, nome_form, NOTA, NUMPED, VALOR, vencimento, TIPOPAG, NUMENTRA, DATAEMI, linha = registro

    # Preencher os campos com os dados extraídos
    campo_data = {
        "NUMPED": NUMPED,
        "NUMENTRA": NUMENTRA,
        "NOTA": NOTA,
        "CODFORN": CODFORN,
        "VALOR": VALOR,
        "DATAEMI": DATAEMI
    }


    for i, (key, value) in enumerate(campo_data.items(), start=4):  # Ajuste o start para o índice correto
        xpath = f"//*[@id='question-list']/div[{i}]/div[2]/div/span/input"
        campo_texto = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        campo_texto.send_keys(value)
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

        opcaovenci = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, f"//span[text()='{vencimento}']"))
        )
        opcaovenci.click()

        # Upload do arquivo
        upload_element = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="question-list"]/div[12]/div[2]/div/button'))
        )
        input_file = driver.find_element(By.XPATH, '//*[@id="question-list"]/div[12]/div[2]/div/input')
        input_file.send_keys(f"C:\\path\\to\\file.pdf")
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

# Fechar o navegador do Selenium
driver.quit()
