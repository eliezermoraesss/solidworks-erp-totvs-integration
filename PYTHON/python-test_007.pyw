import pyodbc
import pandas as pd
from PyQt5.QtWidgets import QApplication, QTextBrowser, QVBoxLayout, QWidget

# Parâmetros de conexão com o banco de dados SQL Server
server = 'SVRERP,1433'
database = 'PROTHEUS12_R27'
username = 'coognicao'
password = '0705@Abc'
driver = '{ODBC Driver 17 for SQL Server}'

# Caminho para o arquivo Excel (caminho bruto)
excel_file_path = r'\\192.175.175.4\f\INTEGRANTES\ELIEZER\PROJETO SOLIDWORKS TOTVS\E7047-001-002.xlsx'

# Arrays para armazenar os resultados
resultados = []

# Tente estabelecer a conexão com o banco de dados
try:
    conn = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')
    
    # Cria um cursor para executar comandos SQL
    cursor = conn.cursor()

    # Query SELECT
    select_query = """
    SELECT G1_COD, G1_QUANT 
    FROM PROTHEUS12_R27.dbo.SG1010 
    WHERE G1_COD = 'E7047-001-002' 
        AND G1_REVFIM <> 'ZZZ' 
        AND D_E_L_E_T_ <> '*'
        AND G1_REVFIM = (
            SELECT MAX(G1_REVFIM) 
            FROM PROTHEUS12_R27.dbo.SG1010 
            WHERE G1_COD = 'E7047-001-002'
            AND G1_REVFIM <> 'ZZZ'
        );
    """

    # Executa a query SELECT e obtém os resultados em um DataFrame
    df_sql = pd.read_sql(select_query, conn)

    # Carrega a planilha do Excel em um DataFrame
    df_excel = pd.read_excel(excel_file_path, sheet_name='Planilha1', header=None)
    
    # Remove espaços em branco das colunas
    df_sql['G1_COD'] = df_sql['G1_COD'].str.strip()
    df_excel.iloc[:, 0] = df_excel.iloc[:, 0].str.strip()

    # Mescla os DataFrames usando a coluna 'G1_COD'
    df_merge = pd.merge(df_sql, df_excel, left_on='G1_COD', right_on=df_excel.iloc[:, 0], how='outer', suffixes=('_SQL', '_Excel'))

    # Compara as quantidades
    df_merge['DIFERENCA_QUANTIDADE'] = df_merge['G1_QUANT'] - df_merge.iloc[:, 3]

    # Adiciona os resultados à lista
    for index, row in df_merge.iterrows():
        resultado = f"Código: {row['G1_COD']}, Quantidade SQL: {row['G1_QUANT']}, Quantidade Excel: {row.iloc[3]}, Diferença: {row['DIFERENCA_QUANTIDADE']}"
        resultados.append(resultado)

except pyodbc.Error as ex:
    # Adiciona mensagem de erro à lista de resultados
    resultados.append(f"Falha na conexão ou consulta. Erro: {str(ex)}")

finally:
    # Fecha a conexão com o banco de dados
    conn.close()

# Inicia o aplicativo PyQt5
app = QApplication([])

# Cria a janela de resultados com barra de rolagem
window = QWidget()
layout = QVBoxLayout()
text_browser = QTextBrowser()
text_browser.setPlainText("\n".join(resultados))
layout.addWidget(text_browser)
window.setLayout(layout)

# Exibe a janela
window.show()

# Executa o loop de eventos do aplicativo
app.exec_()
