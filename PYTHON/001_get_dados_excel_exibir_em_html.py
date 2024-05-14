import pandas as pd
from flask import Flask
import os

app = Flask(__name__)

@app.route('/')
def display_excel_data():
    # Caminho para o arquivo Excel
    file_path = os.path.join('\\\\192.175.175.4', 'f', 'INTEGRANTES', 'ELIEZER', 'PROJETO SOLIDWORKS TOTVS', 'M-048-020-285.xlsx')

    # Lendo o arquivo Excel
    df = pd.read_excel(file_path)

    # Convertendo os dados para HTML
    html_table = df.to_html()

    # Criando a página HTML
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Dados do Excel</title>
    </head>
    <body>
        <h1>Dados do Excel</h1>
        {html_table}
    </body>
    </html>
    """

    return html_content

if __name__ == '__main__':
    app.run(debug=True)
