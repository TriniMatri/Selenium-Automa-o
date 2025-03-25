import time
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import pandas as pd

file_path = 'teste.xlsx'
df = pd.read_excel('')
                   
Fornecedor = df['Fornecedor']
NOTA = df['NOTA']
NPedido = df['N° Pedido']
Valor = df['Valor']
Vencimento = df['Vencimento']
Modo_de_Pagamento = df['Modo de Pagamento']
Entrada = df['Entrada']
Data_de_emissao = df['Data de emissão']

print("Fornecedor:", Fornecedor)
print("NOTA:", NOTA)
print("N° Pedido:", NPedido)
print("Valor:", Valor)
print("Vencimento:", Vencimento)
print("Modo de Pagamento:", Modo_de_Pagamento)
print("Entrada:", Entrada)
print("Data de emissão:", Data_de_emissao)
