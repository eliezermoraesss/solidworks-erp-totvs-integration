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

        # Criar barras de rolagem vertical e horizontal
        scroll_y = tk.Scrollbar(root, orient="vertical")
        scroll_x = tk.Scrollbar(root, orient="horizontal")

        # Tabela para exibir os resultados
        self.tree = ttk.Treeview(root, yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        self.tree.grid(row=3, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")
        scroll_y.grid(row=3, column=2, sticky="ns")
        scroll_x.grid(row=4, column=0, columnspan=2, sticky="ew")

        # Configurar as barras de rolagem para rolar a tabela vertical e horizontalmente
        scroll_y.config(command=self.tree.yview)
        scroll_x.config(command=self.tree.xview)

        # Configurar a expansão da tabela para ocupar todo o espaço disponível
        root.grid_rowconfigure(3, weight=1)
        root.grid_columnconfigure(1, weight=1)

        # Adicionar evento de clique para copiar linha para a área de transferência
        self.tree.bind("<ButtonRelease-1>", self.copiar_linha)

    def executar_consulta(self):
        # Obter os valores dos campos de consulta
        codigo = self.codigo_var.get()
        descricao = self.descricao_var.get()

        # Construir a query de consulta
        select_query = f"""
        SELECT B1_COD, B1_DESC, B1_XDESC2, B1_TIPO, B1_UM, B1_LOCPAD, B1_GRUPO, B1_ZZNOGRP, B1_CC, B1_MSBLQL, B1_REVATU
        FROM PROTHEUS12_R27.dbo.SB1010
        WHERE B1_COD LIKE '{codigo}%' AND B1_DESC LIKE '%{descricao}%'
        """

        try:
            # Estabelecer a conexão com o banco de dados
            conn = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')
            
            # Criar um cursor para executar comandos SQL
            cursor = conn.cursor()

            # Executar a consulta
            cursor.execute(select_query)

            # Limpar a tabela
            self.tree.delete(*self.tree.get_children())

            # Obter os nomes das colunas
            colunas = [column[0] for column in cursor.description]

            # Configurar as colunas na árvore
            self.tree["columns"] = colunas
            for coluna in colunas:
                self.tree.column(coluna, width=tk.AUTO)
                self.tree.heading(coluna, text=coluna, anchor=tk.W)

            # Preencher a tabela com os resultados
            for i, row in enumerate(cursor.fetchall(), start=1):
                self.tree.insert("", i, values=row)

        except pyodbc.Error as ex:
            print(f"Falha na consulta. Erro: {str(ex)}")

        finally:
            # Fechar a conexão com o banco de dados
            conn.close()

    def copiar_linha(self, event):
        # Obter a coluna clicada
        coluna_clicada = self.tree.identify_column(event.x)
        
        # Obter o item (linha) clicado
        item_clicado = self.tree.selection()

        if item_clicado and coluna_clicada:
            # Obter o índice da coluna clicada
            indice_coluna = int(coluna_clicada.replace("#", ""))

            # Obter os valores da linha clicada
            valores_linha = self.tree.item(item_clicado, "values")

            # Obter o valor do campo clicado
            valor_campo = valores_linha[indice_coluna - 1]

            # Copiar para a área de transferência
            pyperclip.copy(str(valor_campo))

if __name__ == "__main__":
    # Parâmetros de conexão com o banco de dados SQL Server
    server = 'SVRERP,1433'
    database = 'PROTHEUS1233_HML'
    username = 'coognicao'
    password = '0705@Abc'
    driver = '{ODBC Driver 17 for SQL Server}'

    root = tk.Tk()
    app = ConsultaApp(root)
    root.mainloop()
