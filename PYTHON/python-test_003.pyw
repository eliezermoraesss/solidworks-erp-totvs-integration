import pyodbc
import ctypes

# Parâmetros de conexão com o banco de dados SQL Server
server = 'SVRERP,1433'  # Substitua 'porta' pelo número da porta do seu servidor SQL Server
database = 'PROTHEUS1233_HML'
username = 'coognicao'
password = '0705@Abc'
driver = '{ODBC Driver 17 for SQL Server}'  # Verifique se este é o driver correto

# Tente estabelecer a conexão com o banco de dados
try:
    conn = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')
    
    # Cria um cursor para executar comandos SQL
    cursor = conn.cursor()

    # Query de inserção
    insert_query = """
    INSERT INTO PROTHEUS1233_HML.dbo.SBM010 (
        BM_FILIAL, BM_GRUPO, BM_DESC, BM_PICPAD, BM_PROORI, BM_CODMAR, BM_STATUS,
        BM_GRUREL, BM_TIPGRU, BM_MARKUP, BM_PRECO, BM_MARGPRE, BM_LENREL, BM_TIPMOV,
        BM_CODGRT, BM_CLASGRU, BM_FORMUL, BM_TPSEGP, BM_DTUMOV, BM_HRUMOV, BM_CONC,
        BM_CORP, BM_EVENTO, BM_LAZER, D_E_L_E_T_, R_E_C_N_O_, R_E_C_D_E_L_,
        BM_USERLGI, BM_USERLGA
    ) VALUES (
        N'0101', N'2830', N'TESTE Eliezer', N'                              ',
        N'1', N'   ', N'1', N'                                        ',
        N'1 ', 0.0, N'   ', 0.0, 0.0, N' ', N'  ', N'1', N'      ', N' ',
        N'        ', N'     ', N' ', N'F', N'F', N'F', N' ', 306, 0, N' ', N' '
    )
    """

    # Executa a query de inserção
    cursor.execute(insert_query)

    # Confirma a transação
    conn.commit()

    # Exibe uma caixa de diálogo se a inserção for bem-sucedida
    ctypes.windll.user32.MessageBoxW(0, "Inserção bem-sucedida na tabela SBM010!", "Sucesso", 1)

except pyodbc.Error as ex:
    # Exibe uma caixa de diálogo se a conexão ou a inserção falhar
    ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão ou inserção. Erro: {str(ex)}", "Erro", 0)

finally:
    # Fecha a conexão com o banco de dados
    conn.close()
