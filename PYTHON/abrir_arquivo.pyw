import subprocess

caminho_do_arquivo = r"\\192.175.175.4\f\INTEGRANTES\ELIEZER\PROJETO SOLIDWORKS TOTVS\VBA\PYTHON\python-test_013_PESQUISA_PYQT5.pyw"

try:
    subprocess.run(["start", caminho_do_arquivo], shell=True)
except Exception as e:
    print(f"Ocorreu um erro: {e}")
