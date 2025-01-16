import sys
import json
import os
import webbrowser
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QPushButton, QTableWidget, QTableWidgetItem, QTreeWidget, 
                           QTreeWidgetItem, QSplitter, QLineEdit, QLabel, QHBoxLayout)
from PyQt5.QtCore import Qt
from sqlalchemy import create_engine
from typing import List, Dict
from db_mssql import setup_mssql

class BOMViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.engine = None
        self.all_components = []
        self.setWindowTitle('Visualizador de Estrutura')
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
        self.table.setSortingEnabled(False)
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

        # Setup banco de dados e carregar dados
        self.setup_database()
        self.load_data()
        
    def setup_database(self):
        username, password, database, server = setup_mssql()
        driver = '{SQL Server}'
        conn_str = f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}'
        self.engine = create_engine(f'mssql+pyodbc:///?odbc_connect={conn_str}')

    def get_components(self, parent_code: str, parent_qty: float = 1.0, level: int = 1):
        query = f"""
        SELECT 
            G1_COD, G1_COMP, G1_XUM, G1_QUANT 
        FROM 
            PROTHEUS12_R27.dbo.SG1010 
        WHERE 
            G1_COD = '{parent_code}' 
            AND G1_REVFIM <> 'ZZZ' 
            AND D_E_L_E_T_ <> '*' 
            AND G1_REVFIM = (
                SELECT MAX(G1_REVFIM) 
                FROM PROTHEUS12_R27.dbo.SG1010 
                WHERE G1_COD = '{parent_code}'
                AND G1_REVFIM <> 'ZZZ' 
                AND D_E_L_E_T_ <> '*'
            )
        """
        
        try:
            df = pd.read_sql(query, self.engine)
            
            if not df.empty:
                for _, row in df.iterrows():
                    total_qty = parent_qty * row['G1_QUANT']
                    
                    component = {
                        'NIVEL': level,
                        'CODIGO': row['G1_COMP'],
                        'CODIGO_PAI': row['G1_COD'],
                        'UNIDADE': row['G1_XUM'],
                        'QTD_NIVEL': row['G1_QUANT'],
                        'QTD_TOTAL': total_qty
                    }
                    self.all_components.append(component)
                    
                    self.get_components(row['G1_COMP'], total_qty, level + 1)
                    
        except Exception as e:
            print(f"Erro ao processar código {parent_code}: {str(e)}")

    def load_data(self):
        codigo_pai = 'E7047-001-001'
        self.all_components = []
        
        # Adicionar item de nível mais alto
        self.all_components.append({
            'NIVEL': 0,
            'CODIGO': codigo_pai,
            'CODIGO_PAI': None,
            'UNIDADE': 'UN', 
            'QTD_NIVEL': 1.0,
            'QTD_TOTAL': 1.0
        })
        
        self.get_components(codigo_pai)
        self.df = pd.DataFrame(self.all_components)
        self.df.to_excel('bom_hierarquica_v6.xlsx', index=False)
        
        self.setup_table()
        self.build_tree()
        
    def setup_table(self):
        self.table.setColumnCount(len(self.df.columns))
        self.table.setHorizontalHeaderLabels(self.df.columns)
        self.populate_table(self.df)
        self.table.resizeColumnsToContents()
        
    def populate_table(self, df):
        self.table.setRowCount(len(df))
        for i, row in df.iterrows():
            for j, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                self.table.setItem(i, j, item)
    
    def build_tree(self):
        self.tree.clear()
        
        # Dicionário para rastrear o caminho completo de cada item
        processed_paths = set()
        
        def build_hierarchy(df, level=1, parent=None, path=None):
            result = []
            
            if path is None:
                path = []
            
            mask = df['NIVEL'] == level
            if parent is not None:
                mask &= df['CODIGO_PAI'] == parent
            
            current_level = df[mask]
            
            for _, row in current_level.iterrows():
                # Criar caminho único para este item incluindo seus ancestrais
                current_path = tuple(path + [row['CODIGO']])
                
                # Se este caminho específico já foi processado, pular
                if current_path in processed_paths:
                    continue
                    
                processed_paths.add(current_path)
                
                if parent is None:
                    tree_item = QTreeWidgetItem(self.tree)
                else:
                    tree_item = QTreeWidgetItem()
                
                tree_item.setText(0, f"{row['CODIGO']} (Qtd: {row['QTD_TOTAL']:.2f})")
                
                children = build_hierarchy(df, level + 1, row['CODIGO'], list(current_path))
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
                self.df['CODIGO'].str.lower().str.contains(filter_text) |
                self.df['CODIGO_PAI'].astype(str).str.lower().str.contains(filter_text)
            ]
        else:
            filtered_df = self.df
            
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
        for _, row in self.df[self.df['NIVEL'] == 0].iterrows():
            node = {
                'name': f"{row['CODIGO']} (Qtd: {row['QTD_TOTAL']:.2f})",
                'children': self.get_children(row['CODIGO'])
            }
            data.append(node)
        return data
    
    def get_children(self, parent_code):
        children = []
        for _, row in self.df[self.df['CODIGO_PAI'] == parent_code].iterrows():
            child = {
                'name': f"{row['CODIGO']} (Qtd: {row['QTD_TOTAL']:.2f})",
                'children': self.get_children(row['CODIGO'])
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

def main():
    app = QApplication(sys.argv)
    viewer = BOMViewer()
    viewer.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
