import time
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains 
import os

# Caminho do driver do Edge
service = Service(r"C:\\Users\\20263\\OneDrive - BEMOL S A\\Documentos\\edgedriver_win64\\msgedgedriver.exe")
options = Options()
NUMPED = "123456789"
NUMENTRA = "123456789"
CODFORN = "123"
NOTA = "1213"
VALOR = "2039"
DATAEMI = "22/03/2025"
login = "03346804@sempreuninorte.com.br"
senha = "28691901"
TIPOPAG = "BOLETO"
nome_form = "WALT TRINITI"
vencimento = "22/03/2025"
diretorio = r"C:\\\Users\\Erivaldo\\Documents\\trabalho trini\Notas teste" ##Ajustar diretorio para teste
nome_arquivo = f"03.05 NF {NOTA} F {CODFORN} {nome_form}.pdf"
file_path = os.path.join(diretorio, nome_arquivo)

##ABIRNDO O FORMS
driver = webdriver.Edge(options=options)
options.add_argument('--disable-extensions')
driver.get("https://forms.office.com/Pages/ResponsePage.aspx?id=HIFgbJ5zRECh04xWqKUMh5axnQPELMRLlJ1ME-h4TE1UM1JKOVhZNVI4V1BaR1lKR0tSREFSMDVPMC4u") 

##ENTRAR NA CONTA PARA CONSEGUIR ACESSO AO FORMS
login_field = WebDriverWait(driver, 20).until(
EC.element_to_be_clickable((By.XPATH, '/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[3]/div/div/div/div[2]/div[2]/div/input[1]'))
)
login_field.click()
login_field.send_keys(login)
loginbutton = WebDriverWait(driver, 20).until(
 EC.element_to_be_clickable((By.XPATH, '//*[@id="idSIButton9"]'))
)
loginbutton.click()
time.sleep(30)

senha_field = WebDriverWait(driver, 10).until(
    EC.visibility_of_element_located((By.XPATH, '//*[@id="passwordInput"]')))
ActionChains(driver).move_to_element(senha_field).click().perform()
senha_field.send_keys(senha)
submit_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="submitButton"]'))
)
submit_button.click()

##ESCOLENDO BEMOL
opcaosetor = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//input[@value='BEMOL S.A']"))
 )   
opcaosetor.click() 
time.sleep(2)

# SETOR
menu_suspenso = driver.find_element(By.XPATH, '//*[@id="rc78bb8d842574508936c58a1235d1db9_placeholder_content"]')
menu_suspenso.click()
wait = WebDriverWait(driver, 3)
opcao_ti = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='TI']")))
wait = WebDriverWait(driver, 3)
opcao_ti.click()
time.sleep(2)

# Tipo PEDIDO
menu_uso_consumo = driver.find_element(By.XPATH, '//*[@id="r5cb66de3650f44d28e7b0b1e880a851a_placeholder_content"]')
driver.execute_script("arguments[0].scrollIntoView(true);", menu_uso_consumo)
menu_uso_consumo = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="r5cb66de3650f44d28e7b0b1e880a851a_placeholder_content"]'))
)
menu_uso_consumo.click()
opcao_uso_consumo = WebDriverWait(driver, 5).until(
    EC.element_to_be_clickable((By.XPATH, "//span[text()='USO E CONSUMO']"))
)
opcao_uso_consumo.click()
time.sleep(2)

# NÚMERO DO PEDIDO
campo_texto = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//*[@id='question-list']/div[4]/div[2]/div/span/input"))
    )
campo_texto.send_keys(NUMPED)
time.sleep(2)

# NÚMERO DE ENTARDA
campo_texto = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//*[@id='question-list']/div[5]/div[2]/div/span/input"))
    )
campo_texto.send_keys(NUMENTRA)
time.sleep(2)

# NÚMERO NOTA 
campo_texto = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//*[@id='question-list']/div[6]/div[2]/div/span/input"))
    )
campo_texto.send_keys(NOTA)

# FORNACEDOR
campo_texto = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//*[@id='question-list']/div[7]/div[2]/div/span/input"))
    )
campo_texto.send_keys(CODFORN)
time.sleep(2)

# VALOR
campo_texto = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//*[@id='question-list']/div[8]/div[2]/div/span/input"))
    )
campo_texto.send_keys(VALOR)

# DATA EMISSÃO
campo_texto = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//*[@id='question-list']/div[9]/div[2]/div/span/input"))
    )
campo_texto.send_keys(DATAEMI)
time.sleep(3)


##TIPOPAG
menu_TIPAG = driver.find_element(By.XPATH, '//*[@id="re46b7db9d675498babd619893a9dadec_placeholder_content"]') 
menu_TIPAG.click() 
opcaoTIPAG = WebDriverWait(driver, 5).until(
    EC.element_to_be_clickable((By.XPATH, f"//span[text()='{TIPOPAG}']"))
)

if TIPOPAG == "BOLETO": 
    #selecionando vencimento (criar print quando data variavel não encontrada)
    opcaoTIPAG.click()
    menu_vencimento = driver.find_element(By.XPATH, '//*[@id="rfdc43eac78b74c17b4aee0f6ce4c81f8_placeholder_content"]') 
    menu_vencimento.click() 
    opcaovenci = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[text()='{}']".format(vencimento)))
    )
    opcaovenci.click()

    #upload do arquive
    upload_element = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="question-list"]/div[12]/div[2]/div/button'))
)
    input_file = driver.find_element(By.XPATH, '//*[@id="question-list"]/div[12]/div[2]/div/input') 
    time.sleep(5)
    input_file.send_keys(file_path)
    time.sleep(10)
    
    enviarbutton = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="form-main-content1"]/div/div/div[2]/div[3]/div/button'))
    )
    enviarbutton.click()
    time.sleep(3)
    
elif TIPOPAG == "DEPOSITO": #Verifiacar os campos do Depósito
    opcaoTIPAG.click()
    

opcaoTIPAG.click()

time.sleep(3)
