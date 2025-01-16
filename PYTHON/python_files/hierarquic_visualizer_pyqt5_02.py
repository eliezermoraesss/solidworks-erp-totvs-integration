import sys
import sqlite3
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QPushButton, QTableWidget, QTableWidgetItem, QTreeWidget, 
                           QTreeWidgetItem, QSplitter, QLineEdit, QLabel, QHBoxLayout)
from PyQt5.QtCore import Qt
import pandas as pd

class HierarchicalViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Visualizador Hierárquico')
        self.setGeometry(100, 100, 1200, 800)
        
        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # Criar layout para filtro
        filter_layout = QHBoxLayout()
        filter_label = QLabel("Filtrar:")
        self.filter_input = QLineEdit()
        self.filter_input.textChanged.connect(self.filter_tables)
        filter_layout.addWidget(filter_label)
        filter_layout.addWidget(self.filter_input)
        layout.addLayout(filter_layout)
        
        # Splitter para dividir tabela e árvore
        splitter = QSplitter(Qt.Horizontal)
        
        # Tabela
        self.table = QTableWidget()
        self.table.setSortingEnabled(True)
        splitter.addWidget(self.table)
        
        # Árvore
        self.tree = QTreeWidget()
        self.tree.setHeaderLabel('Hierarquia')
        splitter.addWidget(self.tree)
        
        layout.addWidget(splitter)
        
        # Carregar dados
        self.load_sample_data()  # Você pode trocar por load_data() para usar banco de dados
        
    def load_sample_data(self):
        # Dados de exemplo - substituir pelo seu read_sql
        data = {
            'Nivel': [1,1,1,1,1,1,1,1,2,2,2,2,2,2],
            'CÓDIGO': ['C-008-069-229', 'C-008-069-227', 'E7000-009-201', 'E7000-009-202', 
                      'E7000-009-264', 'E7000-009-265', 'M-033-012-166', 'M-033-012-144',
                      'C-008-039-365', 'C-008-100-012', 'C-008-105-003', 'C-008-105-003',
                      'C-008-100-008', 'C-008-100-008'],
            'PAI': ['E7047-001-182', 'E7047-001-182', 'E7047-001-182', 'E7047-001-182',
                    'E7047-001-182', 'E7047-001-182', 'E7047-001-182', 'E7047-001-182',
                    'M-033-012-144', 'M-033-012-166', 'E7000-009-265', 'E7000-009-264',
                    'E7000-009-202', 'E7000-009-201'],
            'QUANTIDADE': [24, 65, 12, 6, 3, 3, 4, 8, 1.6, 0.86, 0.3, 0.45, 7.2, 8.4],
            'REVISAO': ['002', '002', '002', '002', '002', '002', '002', '002', 
                       '003', '003', '001', '001', '001', '001']
        }
        self.df = pd.DataFrame(data)
        self.setup_table()
        self.build_tree()

    def load_data(self):
        # Exemplo de como usar com banco de dados
        conn = sqlite3.connect('seu_banco.db')
        query = """
        SELECT Nivel, CÓDIGO, PAI, QUANTIDADE, REVISAO 
        FROM sua_tabela
        ORDER BY Nivel, CÓDIGO
        """
        self.df = pd.read_sql(query, conn)
        conn.close()
        self.setup_table()
        self.build_tree()
        
    def setup_table(self):
        # Configurar colunas
        self.table.setColumnCount(len(self.df.columns))
        self.table.setHorizontalHeaderLabels(self.df.columns)
        
        # Preencher dados
        self.populate_table(self.df)
        
        # Ajustar tamanho das colunas
        self.table.resizeColumnsToContents()
        
    def populate_table(self, df):
        self.table.setRowCount(len(df))
        for i, row in df.iterrows():
            for j, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                self.table.setItem(i, j, item)
    
    def build_tree(self):
        self.tree.clear()
        
        def build_hierarchy(df, level=1, parent=None):
            result = []
            
            # Filtrar itens do nível atual e com o pai especificado (se houver)
            mask = df['Nivel'] == level
            if parent is not None:
                mask &= df['PAI'] == parent
            
            current_level = df[mask]
            
            for _, row in current_level.iterrows():
                item_data = {
                    'Nivel': row['Nivel'],
                    'CÓDIGO': row['CÓDIGO'],
                    'PAI': row['PAI'],
                    'QUANTIDADE': row['QUANTIDADE'],
                    'REVISAO': row['REVISAO']
                }
                
                # Criar item da árvore
                if parent is None:
                    tree_item = QTreeWidgetItem(self.tree)
                else:
                    tree_item = QTreeWidgetItem()
                
                # Configurar texto e dados
                tree_item.setText(0, f"{row['CÓDIGO']} (Qtd: {row['QUANTIDADE']})")
                
                # Recursivamente adicionar filhos
                children = build_hierarchy(df, level + 1, row['CÓDIGO'])
                if children:
                    if parent is None:
                        self.tree.addTopLevelItem(tree_item)
                    for child in children:
                        tree_item.addChild(child)
                        
                result.append(tree_item)
            
            return result
        
        # Construir hierarquia
        build_hierarchy(self.df)
        
        # Expandir todos os itens
        self.tree.expandAll()
        
    def filter_tables(self):
        filter_text = self.filter_input.text().lower()
        
        # Filtrar DataFrame
        if filter_text:
            filtered_df = self.df[
                self.df['CÓDIGO'].str.lower().str.contains(filter_text) |
                self.df['PAI'].str.lower().str.contains(filter_text)
            ]
        else:
            filtered_df = self.df
            
        # Atualizar tabela
        self.populate_table(filtered_df)
        
        # Atualizar árvore
        self.build_tree()  # Reconstruir árvore com dados filtrados
        
        # Destacar itens filtrados na árvore
        if filter_text:
            self.highlight_tree_items(self.tree, filter_text)
    
    def highlight_tree_items(self, tree, filter_text):
        def process_item(item):
            text = item.text(0).lower()
            if filter_text in text:
                item.setBackground(0, Qt.yellow)
            else:
                item.setBackground(0, Qt.white)
            
            for i in range(item.childCount()):
                process_item(item.child(i))
        
        for i in range(tree.topLevelItemCount()):
            process_item(tree.topLevelItem(i))

if __name__ == '__main__':
    app = QApplication(sys.argv)
    viewer = HierarchicalViewer()
    viewer.show()
    sys.exit(app.exec_())