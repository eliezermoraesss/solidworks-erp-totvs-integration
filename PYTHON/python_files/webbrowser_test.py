import webbrowser
import os

# URL que queremos abrir
url = 'https://www.google.com'

# Possíveis caminhos do Firefox no Windows
firefox_paths = [
    r"C:\Program Files\Mozilla Firefox\firefox.exe",
    r"C:\Program Files (x86)\Mozilla Firefox\firefox.exe",
    os.path.expanduser(r"~\AppData\Local\Mozilla Firefox\firefox.exe"),
]

# Tenta encontrar o Firefox
firefox_path = None
for path in firefox_paths:
    if os.path.exists(path):
        firefox_path = path
        break

if firefox_path:
    # Registra o Firefox como navegador
    browser = webbrowser.Mozilla(firefox_path)
    webbrowser.register('firefox', None, browser)
    
    # Abre a URL no Firefox
    webbrowser.get('firefox').open(url)
else:
    # Se não encontrar o Firefox, usa o navegador padrão
    print("Firefox não encontrado. Abrindo no navegador padrão...")
    webbrowser.open(url)
