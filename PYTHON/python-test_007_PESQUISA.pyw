import tkinter as tk
from tkinter import ttk
import pandas as pd
import pyperclip

class FormularioApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Formulário e Tabela")

        # Variáveis para armazenar os dados do formulário
        self.nome_var = tk.StringVar()
        self.idade_var = tk.StringVar()

        # Criar campos do formulário
        lbl_nome = tk.Label(root, text="Nome:")
        lbl_nome.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        entry_nome = tk.Entry(root, textvariable=self.nome_var)
        entry_nome.grid(row=0, column=1, padx=10, pady=5)

        lbl_idade = tk.Label(root, text="Idade:")
        lbl_idade.grid(row=1, column=0, padx=10, pady=5, sticky="w")
        entry_idade = tk.Entry(root, textvariable=self.idade_var)
        entry_idade.grid(row=1, column=1, padx=10, pady=5)

        # Botão para executar a consulta
        btn_consultar = tk.Button(root, text="Consultar", command=self.executar_consulta)
        btn_consultar.grid(row=2, column=0, columnspan=2, pady=10)

        # Tabela para exibir os resultados
        self.tree = ttk.Treeview(root, columns=("Nome", "Idade"))
        self.tree.grid(row=3, column=0, columnspan=2, padx=10, pady=5)
        self.tree.heading("#0", text="ID")
        self.tree.heading("Nome", text="Nome")
        self.tree.heading("Idade", text="Idade")

        # Adicionar eventos de clique para copiar campo a campo
        self.tree.bind("<ButtonRelease-1>", self.copiar_campo)

    def executar_consulta(self):
        # Aqui você deve realizar a consulta ao banco de dados usando os dados do formulário
        # Neste exemplo, estamos criando dados fictícios
        dados = [
            {"Nome": "João", "Idade": 25},
            {"Nome": "Maria", "Idade": 30},
            # Adicione os resultados reais da sua consulta aqui
        ]

        # Limpar a tabela
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Preencher a tabela com os resultados
        for i, dado in enumerate(dados, start=1):
            self.tree.insert("", i, values=(dado["Nome"], dado["Idade"]))

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

if __name__ == "__main__":
    root = tk.Tk()
    app = FormularioApp(root)
    root.mainloop()
