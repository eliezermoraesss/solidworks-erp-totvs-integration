import win32print
import win32api
import os

def imprimir_pdf(caminho_pdf):
    # Certifique-se de que o caminho está correto
    caminho_pdf = os.path.abspath(caminho_pdf)

    if not os.path.exists(caminho_pdf):
        print(f"Erro: O arquivo '{caminho_pdf}' não existe.")
        return

    printer_name = win32print.GetDefaultPrinter()

    try:
        win32api.ShellExecute(0, "print", caminho_pdf, None, ".", 0)
        print(f"Arquivo '{caminho_pdf}' enviado para a impressora '{printer_name}'.")
    except Exception as e:
        print(f"Erro ao imprimir: {e}")

# Chamada com raw string para evitar problemas de escape
imprimir_pdf(r"V:\ORDEM_DE_PRODUCAO\OP_01483601022_M-039-033-110.pdf")
