import pandas as pd
import os


def consolidar_quantidades(arquivo_excel, nome_aba, coluna_codigo, coluna_quantidade):
    # Ler a planilha de Excel
    df = pd.read_excel(arquivo_excel, sheet_name=nome_aba)

    # Verificar se as colunas especificadas existem no DataFrame
    if coluna_codigo not in df.columns or coluna_quantidade not in df.columns:
        raise ValueError("Colunas especificadas não foram encontradas na planilha")

    # Consolidar as quantidades dos códigos repetidos
    consolidado = df.groupby(coluna_codigo)[coluna_quantidade].sum().reset_index()

    return consolidado


arquivo_excel = os.path.join('\\\\192.175.175.4', 'desenvolvimento' , 'PROJETOS_DEV', 'DEV-0003-24_FUNCAO_COMERCIAL_SMARTPLIC', 'DADOS_PARA_COMPARAR.xlsx')
nome_aba = 'MP-DE-PAI-SEM-FILTRO-REV'  # Substitua pelo nome da aba da sua planilha
coluna_codigo = 'CODIGO'  # Substitua pelo nome da coluna de códigos
coluna_quantidade = 'QTD'  # Substitua pelo nome da coluna de quantidades

consolidado_df = consolidar_quantidades(arquivo_excel, nome_aba, coluna_codigo, coluna_quantidade)
print(consolidado_df)

# Se desejar salvar o resultado em um novo arquivo Excel
consolidado_df.to_excel('MP_BOM_COM_FILTRO.xlsx', index=False)
