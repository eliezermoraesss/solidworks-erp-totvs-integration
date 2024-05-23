import inspect

def get_last_line_number():
        caminho_arquivo_python = r"\\192.175.175.4\f\INTEGRANTES\ELIEZER\PROJETO SOLIDWORKS TOTVS\VBA\PYTHON\botao-salvar-bom-solidworks-totvs_CRUD_DEV.pyw"
        with open(caminho_arquivo_python, 'r') as file:
            lines = file.readlines()
            return len(lines)
    
lastline = get_last_line_number()

print(lastline)