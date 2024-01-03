import os

def abrir_arquivo(caminho):
    try:
        # Usa a função os.startfile para abrir o arquivo com o aplicativo padrão
        os.startfile(caminho)
    except Exception as e:
        print(f"Erro ao abrir o arquivo: {e}")

# Substitua 'caminho_do_arquivo' pelo caminho completo do arquivo que você deseja abrir
caminho_do_arquivo = r'\\192.175.175.4\f\INTEGRANTES\ELIEZER\PROJETO SOLIDWORKS TOTVS\VBA\PYTHON\python-test_013_PESQUISA_PYQT5.pyw'
abrir_arquivo(caminho_do_arquivo)
