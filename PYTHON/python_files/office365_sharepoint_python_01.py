from shareplum import Site
from shareplum import Office365
from shareplum.site import Version

# Defina os detalhes de autenticação
site_url = 'https://enaplic.sharepoint.com'
username = 'eliezer@enaplic.com.br'
password = '9a7e5i3o1U-+'

# URL do arquivo no SharePoint
file_url = "/sites/QP/Documentos Compartilhados/QP ABERTA/QP-E7769 -  NARDINI - EQUIPAMENTOS ESMALTAÇÃO_NOVA_VERSÃO.xlsm"

# Crie uma conexão com o SharePoint
authcookie = Office365(site_url, username=username, password=password).GetCookies()
site = Site(site_url, version=Version.v365, authcookie=authcookie)

# Obtenha o arquivo do SharePoint
file = site.get_file(file_url)

# Baixe o arquivo para o diretório local
local_path = "local_path/seuarquivo.xlsx"
with open(local_path, "wb") as f:
    f.write(file.content)
