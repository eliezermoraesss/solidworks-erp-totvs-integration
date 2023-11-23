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
    # Exibe uma caixa de diálogo se a conexão for bem-sucedida
    conn.close()  # Fecha a conexão com o banco de dados
    ctypes.windll.user32.MessageBoxW(0, "Conexão bem-sucedida com o banco de dados SQL Server!", "Sucesso", 1)
except pyodbc.Error as ex:
    # Exibe uma caixa de diálogo se a conexão falhar
    ctypes.windll.user32.MessageBoxW(0, f"Falha na conexão com o banco de dados. Erro: {str(ex)}", "Erro", 0)
