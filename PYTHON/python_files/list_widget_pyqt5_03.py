from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QListWidget, QPushButton, QLineEdit, QTextEdit, QLabel, QHBoxLayout, QTableWidget, QTableWidgetItem, QHeaderView
from PyQt5.QtCore import Qt, QSize
from PyQt5.QtGui import QIcon, QFont
import sys
import pandas as pd
from datetime import datetime
import os


class MultiSelectApp(QWidget):
    def __init__(self):
        super().__init__()

        # Configuração da janela
        self.setWindowTitle("Filtro de Tabela")
        self.setGeometry(100, 100, 800, 600)

        # Layout principal
        layout = QVBoxLayout()

        # DataFrame original (simulado aqui, substituir pela consulta SQL)
        self.df_original = pd.DataFrame({
            'Coluna1': [1, 2, 3, 4],
            'Coluna2': ['A', 'B', 'C', 'D'],
            'Coluna3': [True, False, True, False]
        })
        self.df_atual = self.df_original.copy()  # Copia do DataFrame para aplicar filtros

        # Tabela para exibir o DataFrame
        self.table = QTableWidget(self)
        layout.addWidget(self.table)

        # Botão para abrir o filtro
        btn_filtro = QPushButton("Aplicar Filtros", self)
        btn_filtro.clicked.connect(self.abrir_filtro)
        layout.addWidget(btn_filtro)

        # Botão de limpar filtros
        btn_limpar = QPushButton("Limpar Filtros", self)
        btn_limpar.clicked.connect(self.limpar_filtros)
        layout.addWidget(btn_limpar)

        # Configuração inicial da tabela
        self.configurar_tabela(self.df_atual)

        # Definir o layout na janela
        self.setLayout(layout)

    def configurar_tabela(self, dataframe):
        """Configura a tabela com os dados do DataFrame"""
        self.table.setRowCount(len(dataframe))
        self.table.setColumnCount(len(dataframe.columns))
        self.table.setHorizontalHeaderLabels(dataframe.columns)

        for i, row in dataframe.iterrows():
            for j, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                self.table.setItem(i, j, item)

        # Configurar a tabela para redimensionar as colunas
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

    def abrir_filtro(self):
        """Simula a abertura do filtro e aplicação no DataFrame"""
        # Exemplo de filtro - Filtrar onde Coluna1 > 2
        self.df_atual = self.df_original[self.df_original['Coluna1'] > 2]
        self.configurar_tabela(self.df_atual)

    def limpar_filtros(self):
        """Restaura o DataFrame original na tabela"""
        self.df_atual = self.df_original.copy()
        self.configurar_tabela(self.df_atual)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MultiSelectApp()
    window.show()
    sys.exit(app.exec_())
