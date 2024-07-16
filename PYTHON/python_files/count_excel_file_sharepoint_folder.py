import requests
from requests_ntlm import HttpNtlmAuth

# Configurações do SharePoint
sharepoint_url = "https://enaplic.sharepoint.com/sites/QP"
username = "eliezer@enaplic.com.br"
password = "9a7e5i3o1U-+"
relative_url = "/sites/QP/Documentos Compartilhados/QP ABERTA/"

# URL para obter os arquivos da pasta no SharePoint
url = f"{sharepoint_url}/_api/web/getfolderbyserverrelativeurl('{relative_url}')/files"

try:
    # Configurar a sessão HTTP com autenticação NTLM
    session = requests.Session()
    session.auth = HttpNtlmAuth(username, password)

    # Executar a requisição GET
    response = session.get(url)
    response.raise_for_status()  # Lança um erro se a resposta não for bem-sucedida

    # Processar a resposta JSON
    files = response.json()["d"]["results"]

    # Contar arquivos Excel
    excel_file_count = sum(
        1 for file in files if file["Name"].lower().endswith('.xlsx') or file["Name"].lower().endswith('.xls'))

    print(f"Número de arquivos Excel: {excel_file_count}")

except requests.exceptions.HTTPError as http_err:
    print(f"HTTP error occurred: {http_err}")
except Exception as err:
    print(f"Other error occurred: {err}")
