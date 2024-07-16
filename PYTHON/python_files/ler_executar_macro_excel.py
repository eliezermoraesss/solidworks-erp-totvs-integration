import os
import win32com.client as win32
import pandas as pd

# Caminho para o arquivo Excel
file_path = r'C:\Users\Eliezer\AppData\Local\Temp\PROJ_INDICATORS.xlsm'

# Verifica se o arquivo existe
if not os.path.exists(file_path):
    print(f"O arquivo {file_path} não foi encontrado.")
else:
    # Abre o Excel
    excel_app = win32.Dispatch('Excel.Application')
    excel_app.Visible = False  # Mantenha o Excel invisível

    try:
        # Abre a pasta de trabalho
        workbook = excel_app.Workbooks.Open(file_path)

        # Executa a macro
        excel_app.Application.Run('PROJ_INDICATORS.xlsm!Macro2')

        # Espera a macro terminar de executar
        excel_app.CalculateUntilAsyncQueriesDone()
        
        # Salva a pasta de trabalho
        workbook.Save()

        # Lê os dados da planilha "BD" em um DataFrame do Pandas
        sheet_name = "BD"
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # Exibe as primeiras linhas do DataFrame
        print(df)

    except Exception as e:
        print(f"Ocorreu um erro: {e}")
    finally:
        # Fecha a pasta de trabalho e o Excel
        workbook.Close(SaveChanges=True)
        excel_app.Quit()
