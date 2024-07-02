import ctypes
import pandas as pd
import pyodbc

server = 'SVRERP,1433'
database = 'PROTHEUS12_R27'  # PROTHEUS12_R27 (base de produção) PROTHEUS1233_HML (base de desenvolvimento/teste)
username = 'coognicao'
password = '0705@Abc'
driver = '{SQL Server}'


def get_quantities(dataframe):
    conn = pyodbc.connect(
        f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')
    relacao_produto_quantidade = {}
    try:

        cursor = conn.cursor()

        for i, row in dataframe.iterrows():
            codigo_produto = row.iloc[0]
            query_consulta_produto = f"""
                    SELECT C2_PRODUTO,
                    B1_DESC, C2_QUANT, C2_OBS, C2_DATRF
                    FROM {database}.dbo.SC2010 op
                    INNER JOIN SB1010 prod ON C2_PRODUTO = B1_COD
                    WHERE C2_PRODUTO = '{codigo_produto.strip()}'
                    AND C2_OBS LIKE '%QP-E6963-TOMAZELLA - CONJ ROLE%'
                    AND op.D_E_L_E_T_ <> '*'
                    ORDER BY op.R_E_C_N_O_ DESC;
                    """

            cursor.execute(query_consulta_produto)
            resultado = cursor.fetchone()

            if resultado is None:
                relacao_produto_quantidade[codigo_produto] = 0
                continue

            if resultado[2] != '':
                relacao_produto_quantidade[codigo_produto] = resultado[2]

    except Exception as ex:
        # Exibe uma caixa de diálogo se a conexão ou a consulta falhar
        ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(ex)}",
                                         "Erro ao consultar o cadastro de produtos", 16 | 0)

    finally:
        # Fecha a conexão com o banco de dados se estiver aberta
        if 'conn' in locals():
            conn.close()
    return relacao_produto_quantidade


dataframe_excel = pd.read_excel('./DIFERENCA_FABRICACAO_E6963-006-A00.xlsx', sheet_name='Planilha1', header=None)
resultado = get_quantities(dataframe_excel)

quantidades_df = pd.DataFrame(list(resultado.items()), columns=['Produto', 'Quantidade'])

quantidades_df.to_excel('./quantidades_produtos.xlsx', index=False)
