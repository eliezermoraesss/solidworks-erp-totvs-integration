import json
import os
import sys
import webbrowser

import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QPushButton, QTableWidget, QTableWidgetItem, QTreeWidget, 
                           QTreeWidgetItem, QSplitter, QLineEdit, QLabel, QHBoxLayout)
from PyQt5.QtCore import Qt
from sqlalchemy import create_engine

from db_mssql import setup_mssql

class HierarchicalViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.engine = None
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
        
        # Botão para visualização hierárquica HTML
        self.view_button = QPushButton('Visualizar Hierarquia no Navegador')
        self.view_button.clicked.connect(self.show_hierarchy)
        layout.addWidget(self.view_button)

        # Setup banco de dados
        self.username, self.password, self.database, self.server = setup_mssql()
        self.driver = '{SQL Server}'
        
        # Carregar dados
        self.load_data()
        
    def load_data(self):
        conn_str = (f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};'
                    f'PWD={self.password}')
        self.engine = create_engine(f'mssql+pyodbc:///?odbc_connect={conn_str}')
        query = f"""
        DECLARE @CodigoPai VARCHAR(50) = 'E7047-001-001';

        WITH MaxRevisoes AS (
            SELECT
                G1_COD,
                MAX(G1_REVFIM) AS MaxRevisao
            FROM
                SG1010
            WHERE
                G1_REVFIM <> 'ZZZ'
                AND D_E_L_E_T_ <> '*'
            GROUP BY
                G1_COD
        ),
        ListMP AS (
            SELECT
                G1.G1_COD AS "PAI",
                G1.G1_COMP AS "CÓDIGO",
                G1.G1_QUANT AS "QUANTIDADE",
                1 AS Nivel,
                G1.G1_REVFIM AS "REVISAO"
            FROM
                SG1010 G1
            INNER JOIN MaxRevisoes MR ON G1.G1_COD = MR.G1_COD AND G1.G1_REVFIM = MR.MaxRevisao
            WHERE
                G1.G1_COD = @CodigoPai
                AND G1.G1_REVFIM <> 'ZZZ'
                AND G1.D_E_L_E_T_ <> '*'
            UNION ALL
            SELECT
                filho.G1_COD,
                filho.G1_COMP,
                filho.G1_QUANT * pai.QUANTIDADE,
                pai.Nivel + 1,
                filho.G1_REVFIM
            FROM
                SG1010 AS filho
            INNER JOIN ListMP AS pai ON
                filho.G1_COD = pai."CÓDIGO"
            INNER JOIN MaxRevisoes MR ON filho.G1_COD = MR.G1_COD AND filho.G1_REVFIM = MR.MaxRevisao
            WHERE
                pai.Nivel < 100
                AND filho.G1_REVFIM <> 'ZZZ'
                AND filho.D_E_L_E_T_ <> '*'
        )
        SELECT
            Nivel, 
            "CÓDIGO", 
            "PAI", 
            "QUANTIDADE", 
            REVISAO 
        FROM
            ListMP;
        """
        self.df = pd.read_sql(query, self.engine)
        self.engine.dispose()
        
        # Configurar tabela e árvore
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
            
            mask = df['Nivel'] == level
            if parent is not None:
                mask &= df['PAI'] == parent
            
            current_level = df[mask]
            
            for _, row in current_level.iterrows():
                if parent is None:
                    tree_item = QTreeWidgetItem(self.tree)
                else:
                    tree_item = QTreeWidgetItem()
                
                tree_item.setText(0, f"{row['CÓDIGO']} (Qtd: {row['QUANTIDADE']:.2f})")
                
                children = build_hierarchy(df, level + 1, row['CÓDIGO'])
                if children:
                    if parent is None:
                        self.tree.addTopLevelItem(tree_item)
                    for child in children:
                        tree_item.addChild(child)
                        
                result.append(tree_item)
            
            return result
        
        build_hierarchy(self.df)
        self.tree.expandAll()
    
    def filter_tables(self):
        filter_text = self.filter_input.text().lower()
        
        if filter_text:
            filtered_df = self.df[
                self.df['CÓDIGO'].str.lower().str.contains(filter_text) |
                self.df['PAI'].str.lower().str.contains(filter_text)
            ]
        else:
            filtered_df = self.df
            
        self.populate_table(filtered_df)
        self.build_tree()
        
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
        
    def show_hierarchy(self):
        hierarchy_data = self.create_hierarchy_data()
        html_content = self.create_html_visualization(hierarchy_data)
        temp_path = 'temp_hierarchy.html'
        with open(temp_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        webbrowser.open('file://' + os.path.realpath(temp_path))
    
    def create_hierarchy_data(self):
        data = []
        for _, row in self.df[self.df['Nivel'] == 1].iterrows():
            node = {
                'name': f"{row['CÓDIGO']} (Qtd: {row['QUANTIDADE']:.2f})",
                'children': self.get_children(row['CÓDIGO'])
            }
            data.append(node)
        return data
    
    def get_children(self, parent_code):
        children = []
        for _, row in self.df[self.df['PAI'] == parent_code].iterrows():
            child = {
                'name': f"{row['CÓDIGO']} (Qtd: {row['QUANTIDADE']:.2f})",
                'children': self.get_children(row['CÓDIGO'])
            }
            children.append(child)
        return children
    
    def create_html_visualization(self, data):
        return f'''
        <!DOCTYPE html>
        <html>
        <head>
            <title>Visualização Hierárquica</title>
            <script src="https://cdnjs.cloudflare.com/ajax/libs/d3/7.8.5/d3.min.js"></script>
            <style>
                .node circle {{
                    fill: #fff;
                    stroke: steelblue;
                    stroke-width: 3px;
                }}
                .node text {{
                    font: 12px sans-serif;
                }}
                .link {{
                    fill: none;
                    stroke: #ccc;
                    stroke-width: 2px;
                }}
            </style>
        </head>
        <body>
            <div id="tree-container"></div>
            <script>
                const treeData = {json.dumps(data)};
                
                const margin = {{top: 20, right: 120, bottom: 30, left: 120}};
                const width = 1200 - margin.left - margin.right;
                const height = 800 - margin.top - margin.bottom;
                
                const svg = d3.select("#tree-container").append("svg")
                    .attr("width", width + margin.left + margin.right)
                    .attr("height", height + margin.top + margin.bottom)
                    .append("g")
                    .attr("transform", `translate(${{margin.left}},${{margin.top}})`);
                
                const tree = d3.tree().size([height, width]);
                const root = d3.hierarchy(treeData[0]);
                const nodes = tree(root);
                
                svg.selectAll(".link")
                    .data(nodes.descendants().slice(1))
                    .enter().append("path")
                    .attr("class", "link")
                    .attr("d", d => `M${{d.y}},${{d.x}}
                        C${{(d.y + d.parent.y) / 2}},${{d.x}}
                         ${{(d.y + d.parent.y) / 2}},${{d.parent.x}}
                         ${{d.parent.y}},${{d.parent.x}}`);
                
                const node = svg.selectAll(".node")
                    .data(nodes.descendants())
                    .enter().append("g")
                    .attr("class", "node")
                    .attr("transform", d => `translate(${{d.y}},${{d.x}})`);
                
                node.append("circle")
                    .attr("r", 10);
                
                node.append("text")
                    .attr("dy", ".35em")
                    .attr("x", d => d.children ? -13 : 13)
                    .style("text-anchor", d => d.children ? "end" : "start")
                    .text(d => d.data.name);
            </script>
        </body>
        </html>
        '''

if __name__ == '__main__':
    app = QApplication(sys.argv)
    viewer = HierarchicalViewer()
    viewer.show()
    sys.exit(app.exec_())