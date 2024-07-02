import ctypes
import pandas as pd
import pyodbc

server = 'SVRERP,1433'
database = 'PROTHEUS12_R27'  # PROTHEUS12_R27 (base de produção) PROTHEUS1233_HML (base de desenvolvimento/teste)
username = 'coognicao'
password = '0705@Abc'
driver = '{SQL Server}'


def get_code(dataframe):
    conn = pyodbc.connect(
        f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')
    relacao_codigo_descricao = {}
    try:

        cursor = conn.cursor()

        for i, row in dataframe.iterrows():
            descricao_produto = row.iloc[0]
            query_consulta_produto = f"""
                    SELECT 
                        B1_COD, B1_DESC 
                        FROM {database}.dbo.SB1010 
                    WHERE
                        B1_DESC LIKE '%{descricao_produto.strip()}%' 
                        AND D_E_L_E_T_ <> '*';
                    """

            cursor.execute(query_consulta_produto)
            resultado_consulta = cursor.fetchone()

            if resultado_consulta is None:
                relacao_codigo_descricao[descricao_produto] = 0
                continue

            if resultado_consulta[0] != '':
                relacao_codigo_descricao[descricao_produto] = resultado_consulta[0]

    except Exception as ex:
        # Exibe uma caixa de diálogo se a conexão ou a consulta falhar
        ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o TOTVS ou consulta. Erro: {str(ex)}",
                                         "Erro ao consultar o cadastro de produtos", 16 | 0)

    finally:
        # Fecha a conexão com o banco de dados se estiver aberta
        if 'conn' in locals():
            conn.close()
            
    return relacao_codigo_descricao


dataframe_excel = pd.read_excel('./MOINHO_MARE_RASCUNHO_V2.xlsx', sheet_name='Planilha1', header=None)
resultado = get_code(dataframe_excel)

codigos_df = pd.DataFrame(list(resultado.items()), columns=['Código', 'Descrição'])

codigos_df.to_excel('./codigos_produtos_04.xlsx', index=False)
