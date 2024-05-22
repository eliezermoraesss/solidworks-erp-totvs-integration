import sqlalchemy as sa

# Adjust the connection string with a longer login timeout
engine = sa.create_engine('mssql+pyodbc://coognicao:0705@Abc@SVRERP,1433/PROTHEUS1233_HML?driver=ODBC+Driver+17+for+SQL+Server&timeout=30')

try:
    with engine.connect():
        print("Connection successful.")
except Exception as ex:
    print(f"Connection failed: {str(ex)}")