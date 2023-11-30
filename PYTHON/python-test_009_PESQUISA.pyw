import tkinter as tk
from tkinter import ttk
import pyodbc
import pyperclip

class ConsultaApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Consulta SQL Server")

        # Variáveis para armazenar os dados da consulta
        self.codigo_var = tk.StringVar()
        self.descricao_var = tk.StringVar()

        # Criar campos para a consulta
        lbl_codigo = tk.Label(root, text="Código:")
        lbl_codigo.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        entry_codigo = tk.Entry(root, textvariable=self.codigo_var)
        entry_codigo.grid(row=0, column=1, padx=10, pady=5)

        lbl_descricao = tk.Label(root, text="Descrição:")
        lbl_descricao.grid(row=1, column=0, padx=10, pady=5, sticky="w")
        entry_descricao = tk.Entry(root, textvariable=self.descricao_var)
        entry_descricao.grid(row=1, column=1, padx=10, pady=5)

        # Botão para executar a consulta
        btn_consultar = tk.Button(root, text="Consultar", command=self.executar_consulta)
        btn_consultar.grid(row=2, column=0, columnspan=2, pady=10)

        # Criar uma barra de rolagem vertical
        scroll_y = tk.Scrollbar(root, orient="vertical")

        # Tabela para exibir os resultados
        self.tree = ttk.Treeview(root, columns=("Código", "Descrição", "XDESC2", "Tipo", "UM", "Localização", "Grupo", "ZZNOGRP", "CC", "MSBLQL", "Revatu"), yscrollcommand=scroll_y.set)
        self.tree.grid(row=3, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")
        scroll_y.grid(row=3, column=2, sticky="ns")

        # Configurar a barra de rolagem para rolar a tabela verticalmente
        scroll_y.config(command=self.tree.yview)

        # Configurar a expansão da tabela para ocupar todo o espaço disponível
        root.grid_rowconfigure(3, weight=1)
        root.grid_columnconfigure(1, weight=1)

        self.tree.heading("Código", text="Código")
        self.tree.heading("Descrição", text="Descrição")
        self.tree.heading("XDESC2", text="Descrição Complementar")
        self.tree.heading("Tipo", text="Tipo")
        self.tree.heading("UM", text="UM")
        self.tree.heading("Localização", text="Localização")
        self.tree.heading("Grupo", text="Grupo")
        self.tree.heading("ZZNOGRP", text="Desc. Grupo")
        self.tree.heading("CC", text="Centro Custo")
        self.tree.heading("MSBLQL", text="Boqueado?")
        self.tree.heading("Revatu", text="Revisão")

    def executar_consulta(self):
        # Obter os valores dos campos de consulta
        codigo = self.codigo_var.get()
        descricao = self.descricao_var.get()

        # Construir a query de consulta
        select_query = f"""
        SELECT B1_COD, B1_DESC, B1_XDESC2, B1_TIPO, B1_UM, B1_LOCPAD, B1_GRUPO, B1_ZZNOGRP, B1_CC, B1_MSBLQL, B1_REVATU
        FROM PROTHEUS12_R27.dbo.SB1010
        WHERE B1_COD LIKE '{codigo}%' AND B1_DESC LIKE '{descricao}%'
        """
        
    def copiar_campo(self, event):
        # Obter a coluna clicada
        coluna_clicada = self.tree.identify_column(event.x)
        
        # Obter o item (linha) clicado
        item_clicado = self.tree.selection()

        if item_clicado and coluna_clicada:
            # Obter o índice da coluna clicada
            indice_coluna = int(coluna_clicada.replace("#", ""))

            # Obter o valor do campo clicado
            valor_campo = self.tree.item(item_clicado, "values")[indice_coluna - 1]

            # Copiar para a área de transferência
            pyperclip.copy(str(valor_campo))

        try:
            # Estabelecer a conexão com o banco de dados
            conn = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')
            
            # Criar um cursor para executar comandos SQL
            cursor = conn.cursor()

            # Executar a consulta
            cursor.execute(select_query)

            # Limpar a tabela
            for item in self.tree.get_children():
                self.tree.delete(item)

            # Preencher a tabela com os resultados
            for i, row in enumerate(cursor.fetchall(), start=1):
                self.tree.insert("", i, values=row)

        except pyodbc.Error as ex:
            print(f"Falha na consulta. Erro: {str(ex)}")

        finally:
            # Fechar a conexão com o banco de dados
            conn.close()

if __name__ == "__main__":
    # Parâmetros de conexão com o banco de dados SQL Server
    server = 'SVRERP,1433'
    database = 'PROTHEUS12_R27'
    username = 'coognicao'
    password = '0705@Abc'
    driver = '{ODBC Driver 17 for SQL Server}'

    root = tk.Tk()
    app = ConsultaApp(root)
    root.mainloop()
