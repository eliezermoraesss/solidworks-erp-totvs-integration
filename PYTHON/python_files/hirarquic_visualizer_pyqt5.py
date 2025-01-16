import json
import os
import sys
import webbrowser

import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QTableWidget, QTableWidgetItem
from sqlalchemy import create_engine

from db_mssql import setup_mssql


class HierarchicalViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.engine = None
        self.setWindowTitle('Visualizador Hierárquico')
        self.setGeometry(100, 100, 800, 600)
        
        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # Tabela
        self.table = QTableWidget()
        layout.addWidget(self.table)
        
        # Botão para visualização hierárquica
        self.view_button = QPushButton('Visualizar Hierarquia')
        self.view_button.clicked.connect(self.show_hierarchy)
        layout.addWidget(self.view_button)

        self.username, self.password, self.database, self.server = setup_mssql()
        self.driver = '{SQL Server}'
        
        # Carregar dados
        self.load_data()
        
    def load_data(self):
        conn_str = (f'DRIVER={self.driver};SERVER={self.server};DATABASE={self.database};UID={self.username};'
                    f'PWD={self.password}')
        self.engine = create_engine(f'mssql+pyodbc:///?odbc_connect={conn_str}')
        query = f"""
        DECLARE @CodigoPai VARCHAR(50) = 'E7047-001-182'; -- Substitua pelo código pai que deseja consultar

        -- CTE para selecionar as revisões máximas
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
        -- CTE para selecionar os itens pai e seus subitens recursivamente
        ListMP AS (
            -- Selecionar o item pai inicialmente
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
            -- Selecione os subitens de cada item pai e multiplique as quantidades
            SELECT
                filho.G1_COD,
                filho.G1_COMP,
                filho.G1_QUANT * pai.QUANTIDADE, -- [DESCOMENTAR (* pai.QUANTIDADE) PARA CALCULAR QUANTIDADE TOTAL DOS FILHOS EM FUNÇÃO DOS PAIS]
                pai.Nivel + 1,
                filho.G1_REVFIM
            FROM
                SG1010 AS filho
            INNER JOIN ListMP AS pai ON
                filho.G1_COD = pai."CÓDIGO"
            INNER JOIN MaxRevisoes MR ON filho.G1_COD = MR.G1_COD AND filho.G1_REVFIM = MR.MaxRevisao
            WHERE
                pai.Nivel < 100
                -- Defina o limite máximo de recursão aqui
                AND filho.G1_REVFIM <> 'ZZZ'
                AND filho.D_E_L_E_T_ <> '*'
        )
            
            -- EXIBE O NÍVEL DOS COMPONENTES
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
        
        # Configurar tabela
        self.setup_table()
        
    def setup_table(self):
        # Configurar colunas
        self.table.setColumnCount(len(self.df.columns))
        self.table.setHorizontalHeaderLabels(self.df.columns)
        
        # Preencher dados
        self.table.setRowCount(len(self.df))
        for i, row in self.df.iterrows():
            for j, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                self.table.setItem(i, j, item)
        
        # Habilitar ordenação
        self.table.setSortingEnabled(True)
        
    def show_hierarchy(self):
        # Converter dados para formato hierárquico
        hierarchy_data = self.create_hierarchy_data()
        
        # Criar arquivo HTML temporário
        html_content = self.create_html_visualization(hierarchy_data)
        temp_path = 'temp_hierarchy.html'
        with open(temp_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        # Abrir no navegador
        webbrowser.open('file://' + os.path.realpath(temp_path))
        
    def create_hierarchy_data(self):
        data = []
        # Processar itens de nível 1
        for _, row in self.df[self.df['Nivel'] == 1].iterrows():
            node = {
                'name': row['CÓDIGO'],
                'children': self.get_children(row['CÓDIGO'])
            }
            data.append(node)
        return data
    
    def get_children(self, parent_code):
        children = []
        for _, row in self.df[self.df['PAI'] == parent_code].iterrows():
            child = {
                'name': row['CÓDIGO'],
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
                // Dados da hierarquia
                const treeData = {json.dumps(data)};
                
                // Configurar visualização
                const margin = {{top: 20, right: 90, bottom: 30, left: 90}};
                const width = 960 - margin.left - margin.right;
                const height = 500 - margin.top - margin.bottom;
                
                // Criar árvore
                const svg = d3.select("#tree-container").append("svg")
                    .attr("width", width + margin.left + margin.right)
                    .attr("height", height + margin.top + margin.bottom)
                    .append("g")
                    .attr("transform", `translate(${{margin.left}},${{margin.top}})`);
                
                const tree = d3.tree().size([height, width]);
                const root = d3.hierarchy(treeData[0]);
                const nodes = tree(root);
                
                // Adicionar links
                svg.selectAll(".link")
                    .data(nodes.descendants().slice(1))
                    .enter().append("path")
                    .attr("class", "link")
                    .attr("d", d => `M${{d.y}},${{d.x}}
                        C${{(d.y + d.parent.y) / 2}},${{d.x}}
                         ${{(d.y + d.parent.y) / 2}},${{d.parent.x}}
                         ${{d.parent.y}},${{d.parent.x}}`);
                
                // Adicionar nós
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
