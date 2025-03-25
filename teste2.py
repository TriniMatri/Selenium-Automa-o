import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import win32com.client as win32
from datetime import datetime
import decimal

# Caminho do driver do Edge
service = Service(r"")
options = Options()

#Caso o formulario a ser prenchido possua Uploas
diretorio = r""  

# Abrir o Excel
excel = win32.Dispatch("Excel.Application")
excel.Visible = False 

# Caminho da planilha no OneDrive (ajuste para o caminho correto no seu PC)
caminho_planilha = r""
workbook = excel.Workbooks.Open(caminho_planilha)
sheet = workbook.Sheets("")  # Ajuste o nome da aba

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
        nome_arquivo = f"03.05 NF {NOTA} F {CODFORN} {nome_form}.pdf"
        file_path = os.path.join(diretorio, nome_arquivo)

# Exibir os dados extraídos
for registro in dados:
    CODFORN, nome_form, NOTA, NUMPED, VALOR, vencimento, TIPOPAG, NUMENTRA, DATAEMI, linha = registro
    # Formatar a data de emissão para dd/mm/aa
    if isinstance(DATAEMI, datetime) and isinstance(vencimento, datetime):
        data_emissao_formatada = DATAEMI.strftime("%d/%m/%y")
        data_formatada = vencimento.strftime("%d/%m/%Y")
    else:
        data_emissao_formatada = str(DATAEMI)
        data_formatada = str(vencimento)

    print(f" {CODFORN}, {nome_form}, {NOTA}, {NUMPED}, {VALOR}, {data_formatada}, {TIPOPAG}, {NUMENTRA}, {data_emissao_formatada}")

# Abrindo o Forms
driver = webdriver.Edge(service=service, options=options)
options.add_argument('--disable-extensions')
driver.get("https://forms.office.com/Pages/ResponsePage.aspx?id=HIFgbJ5zRECh04xWqKUMh5axnQPELMRLlJ1ME-h4TE1UM1JKOVhZNVI4V1BaR1lKR0tSREFSMDVPMC4u")

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

# Escolhendo BEMOL
opcaosetor = WebDriverWait(driver, 5).until(
    EC.presence_of_element_located((By.XPATH, "//input[@value='BEMOL S.A']"))
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
    "data_emissao_formatada": "//*[@id='question-list']/div[9]/div[2]/div/span/input"
}

# Preencher os campos automaticamente
for key, xpath in campo_data.items():

    campo_texto = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, xpath))
    )
    
    # Obter o valor da variável específica
    value_to_send = locals()[key]  # Obtém o valor da variável
    
    # Verificar se a variável é VALOR e se é um número, então formatar com 1 casa decimal
    if key == "VALOR":
        try:
            # Tentar converter para decimal e formatar
            value_to_send = decimal.Decimal(value_to_send).quantize(decimal.Decimal('0.0'))  # Forçar 1 casa decimal
        except (decimal.InvalidOperation, ValueError):
            value_to_send = "0.0"  # Se não puder ser convertido, colocar "0.0"
    
    # Verificar se a variável é uma das que precisam ser convertidas para inteiro
    elif key in ["NUMPED", "NUMENTRA", "NOTA", "CODFORN"]:
        try:
            # Tentar converter para inteiro
            value_to_send = int(value_to_send)
        except (ValueError, TypeError):
            value_to_send = 0  # Se não puder ser convertido, colocar 0
    
    # Para as outras variáveis, mantemos o formato original
    campo_texto.send_keys(str(value_to_send))
    time.sleep(2)

menu_TIPAG = driver.find_element(By.XPATH, '//*[@id="re46b7db9d675498babd619893a9dadec_placeholder_content"]')
menu_TIPAG.click()

opcaoTIPAG = WebDriverWait(driver, 120).until(
    EC.element_to_be_clickable((By.XPATH, f"//span[text()='{TIPOPAG}']"))
)
opcaoTIPAG.click()

if TIPOPAG == "BOLETO":
   menu_vencimento = driver.find_element(By.XPATH, '//*[@id="rfdc43eac78b74c17b4aee0f6ce4c81f8_placeholder_content"]')
   menu_vencimento.click()

   opcaovenci = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, f"//span[text()='{data_formatada}']"))
    )
   opcaovenci.click()
   time.sleep(3)

    # Upload do arquivo
   upload_element = WebDriverWait(driver, 10).until(
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


elif TIPOPAG == "DEPOSITO":
    # Seleção do tipo de pagamento como depósito
    menu_deposito = WebDriverWait(driver, 10).until( 
        EC.element_to_be_clickable((By.XPATH, "//*[@id='r12a1ddbf04e74c24b8a7d90847d2498_placeholder_content']"))
    )
    menu_deposito.click()

time.sleep(2)

# Atualizando a data de envio na planilha
for registro in dados:
    linha = registro[9]  # Número da linha
    sheet.Cells(linha, 12).Value = datetime.today().strftime('%d/%m/%Y')
    print(f"Data de envio atualizada na linha {linha}")

# Finaliza o workbook
workbook.Save()
workbook.Close()