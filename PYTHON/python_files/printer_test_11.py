import subprocess
import os

def imprimir_pdf(pdf_path, impressora):
    sumatra_path = r"W:\REPOSITORIOS\resources\SumatraPDF\SumatraPDF.exe"

    if not os.path.exists(pdf_path):
        print("Arquivo PDF não encontrado!")
        return

    comando = [
        sumatra_path,
        "-print-to", impressora,
        "-print-settings", "duplex,portrait",
        pdf_path
    ]

    try:
        subprocess.run(comando, check=True)
        print("Impressão enviada com sucesso!")
    except subprocess.CalledProcessError as e:
        print(f"Erro ao imprimir: {e}")

# Caminho da pasta de rede
caminho_pdf = r"V:\ORDEM_DE_PRODUCAO\OP_01507401032_M-059-020-686.pdf"
imprimir_pdf(caminho_pdf, "Samsung M337x 387x 407x Series XPS (192.175.175.87)") # HP OfficeJet Pro 7740 series PCL-3 (Rede)
