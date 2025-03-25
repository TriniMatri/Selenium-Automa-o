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

service = Service(r"C:\\Users\\Erivaldo\\Documents\\edgedriver_win64\\msedgedriver.exe")
options = Options()
login = "03346804@sempreuninorte.com.br"
senha = "28691901"

diretorio = r"C:\\Users\\Erivaldo\\Documents\\trabalho trini\\Notas teste"  # Ajustar diretório para teste
nome_arquivo = f"03.05 NF {NOTA} F {CODFORN} {nome_form}.pdf"
file_path = os.path.join(diretorio, nome_arquivo)

# Abrir o Excel
excel = win32.Dispatch("Excel.Application")
excel.Visible = False  # Rodar em segundo plano

# Caminho da planilha no OneDrive (ajuste para o caminho correto no seu PC)
caminho_planilha = r"C:\\Users\\Erivaldo\\OneDrive\\Documentos\\teste.xlsx"
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

# Exibir os dados extraídos
for registro in dados:
    CODFORN, nome_form, NOTA, NUMPED, VALOR, vencimento, TIPOPAG, NUMENTRA, DATAEMI, linha = registro
    # Formatar a data de emissão para dd/mm/aa
    if isinstance(DATAEMI, datetime):
        data_emissao_formatada = DATAEMI.strftime("%d/%m/%y")
    else:
        data_emissao_formatada = str(DATAEMI)

    print(f" {CODFORN}, {nome_form}, {NOTA}, {NUMPED}, {VALOR}, {vencimento}, {TIPOPAG}, {NUMENTRA}, {data_emissao_formatada}")



# Atualizar a "Data de envio" após submissão do formulário
data_hoje = datetime.today().strftime("%d/%m/%Y")

for _, _, _, _, _, _, _, _, _, linha in dados:
    sheet.Cells(linha, 12).Value = data_hoje  # Atualizar coluna "Data de envio" (coluna 12)

# Salvar e fechar a planilha
workbook.Save()
workbook.Close()
excel.Quit()
