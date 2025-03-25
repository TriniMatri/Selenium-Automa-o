import pandas as pd
import requests
from io import BytesIO

# URL do arquivo Excel no OneDrive ou SharePoint (verifique se a URL é pública ou com as permissões adequadas)
url = "URL_PUBLICA_DA_PLANILHhttps://1drv.ms/x/s!ApqMDUsVdYgofNQ7Mkn0YdLvcro?e=W8HHbWA"

# Enviar a requisição HTTP para obter o conteúdo da planilha
response = requests.get(url)

# Verifique se a requisição foi bem-sucedida
if response.status_code == 200:
    # Ler a planilha diretamente da resposta em formato de bytes
    excel_file = BytesIO(response.content)
    
    # Usar o pandas para ler a planilha
    df = pd.read_excel(excel_file, engine='openpyxl')
    
    # Exibir as primeiras linhas da planilha
    print(df.head())
else:
    print(f"Erro ao acessar o arquivo: {response.status_code}")