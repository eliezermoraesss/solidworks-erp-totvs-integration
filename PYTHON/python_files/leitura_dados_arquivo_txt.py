def ler_credencial_mssql():
    caminho_do_arquivo = r"\\192.175.175.4\f\INTEGRANTES\ELIEZER\PROJETO SOLIDWORKS TOTVS\libs-python\user-password-mssql\user-password-mssql.txt"
    try:
        with open(caminho_do_arquivo, 'r') as arquivo:
            string_lida = arquivo.read()
            username, password = string_lida.split(',')
            return username, password
            
    except FileNotFoundError:
        print("O arquivo especificado não foi encontrado.")

    except Exception as e:
        print("Ocorreu um erro ao ler o arquivo:", e)
        
print(ler_credencial_mssql())
