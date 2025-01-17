import sys
import json
import os
import csv
import io
import webbrowser
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QPushButton, QTableWidget, QTableWidgetItem, QTreeWidget, 
                           QTreeWidgetItem, QSplitter, QLineEdit, QLabel, QHBoxLayout,
                           QAbstractItemView, QAction)
from PyQt5.QtCore import Qt
from sqlalchemy import create_engine
from db_mssql import setup_mssql

class BOMViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.engine = None
        self.codigo_pai = 'E7047-001-182'
        self.all_components = []
        self.setWindowTitle('Visualizador de Estrutura')
        
        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # Criar layout para filtro
        filter_layout = QHBoxLayout()
        
        # Filtro por Código
        filter_label_codigo = QLabel("Código:")
        self.filter_input_codigo = QLineEdit()
        self.filter_input_codigo.textChanged.connect(self.filter_tables)
        filter_layout.addWidget(filter_label_codigo)
        filter_layout.addWidget(self.filter_input_codigo)
        
        # Filtro por Descrição
        filter_label_desc = QLabel("Descrição:")
        self.filter_input_desc = QLineEdit()
        self.filter_input_desc.textChanged.connect(self.filter_tables)
        filter_layout.addWidget(filter_label_desc)
        filter_layout.addWidget(self.filter_input_desc)
        
        layout.addLayout(filter_layout)
        
        # Splitter para dividir tabela e árvore
        splitter = QSplitter(Qt.Horizontal)
        
        # Tabela
        self.table = QTableWidget()
        self.table.setSortingEnabled(False)
        splitter.addWidget(self.table)
        
        # Container para árvore e botões
        tree_container = QWidget()
        tree_layout = QVBoxLayout(tree_container)
        
        # Árvore
        self.tree = QTreeWidget()
        self.tree.setHeaderLabel('Hierarquia')
        
        # Mostrar linhas de conexão
        self.tree.setAlternatingRowColors(True)
        self.tree.setIndentation(20)  # Espaçamento da indentação
        self.tree.setRootIsDecorated(True)  # Mostra as linhas de conexão
        self.tree.setItemsExpandable(True)  # Permite expandir/recolher itens
        
        # Aumentar altura das linhas e melhorar espaçamento
        self.tree.setStyleSheet("""
            QTreeWidget::item {
                height: 30px;
                padding: 5px;
                margin: 2px;
            }
            QTreeWidget::item:selected {
                background-color: #E6E6E6;
            }
        """)
        
        tree_layout.addWidget(self.tree)
        
        # Botões de expandir/recolher
        tree_buttons_layout = QHBoxLayout()
        expand_button = QPushButton("Expandir Tudo")
        collapse_button = QPushButton("Recolher Tudo")
        toggle_table_button = QPushButton("Ocultar/Exibir Tabela")
        expand_button.clicked.connect(self.tree.expandAll)
        collapse_button.clicked.connect(self.tree.collapseAll)
        toggle_table_button.clicked.connect(self.toggle_table_visibility)
        tree_buttons_layout.addWidget(expand_button)
        tree_buttons_layout.addWidget(collapse_button)
        tree_buttons_layout.addWidget(toggle_table_button)
        tree_layout.addLayout(tree_buttons_layout)
        
        splitter.addWidget(tree_container)
        
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
            struct.G1_COD,
            struct.G1_COMP,
            prod.B1_DESC AS 'DESCRICAO',
            struct.G1_XUM, 
            struct.G1_QUANT 
        FROM 
            PROTHEUS12_R27.dbo.SG1010 struct
        INNER JOIN
            PROTHEUS12_R27.dbo.SB1010 prod
        ON 
            G1_COMP = B1_COD AND prod.D_E_L_E_T_ <> '*'
        WHERE 
            G1_COD = '{parent_code}' 
            AND G1_REVFIM <> 'ZZZ' 
            AND struct.D_E_L_E_T_ <> '*' 
            AND G1_REVFIM = (
                SELECT MAX(G1_REVFIM) 
                FROM PROTHEUS12_R27.dbo.SG1010 
                WHERE G1_COD = '{parent_code}'
                AND G1_REVFIM <> 'ZZZ' 
                AND struct.D_E_L_E_T_ <> '*'
            )
        """
        
        try:
            df = pd.read_sql(query, self.engine)
            
            if not df.empty:
                for _, row in df.iterrows():
                    total_qty = parent_qty * row['G1_QUANT']
                    
                    component = {
                        'NIVEL': level,
                        'CODIGO': row['G1_COMP'].strip(),
                        'DESCRICAO': row['DESCRICAO'].strip(),
                        'CODIGO_PAI': row['G1_COD'].strip(),
                        'UNIDADE': row['G1_XUM'].strip(),
                        'QTD_NIVEL': row['G1_QUANT'],
                        'QTD_TOTAL': total_qty
                    }
                    self.all_components.append(component)
                    
                    self.get_components(row['G1_COMP'], total_qty, level + 1)
                    
        except Exception as e:
            print(f"Erro ao processar código {parent_code}: {str(e)}")

    def load_data(self):
        codigo_pai = self.codigo_pai
        self.all_components = []
        
        # Adicionar item de nível mais alto
        # Primeiro precisamos buscar a descrição do item pai
        query = f"""
        SELECT B1_DESC as DESCRICAO
        FROM PROTHEUS12_R27.dbo.SB1010
        WHERE B1_COD = '{codigo_pai}'
        AND D_E_L_E_T_ <> '*'
        """
        try:
            df_pai = pd.read_sql(query, self.engine)
            descricao_pai = df_pai['DESCRICAO'].iloc[0] if not df_pai.empty else ''
        except:
            descricao_pai = ''
        
        self.all_components.append({
            'NIVEL': 0,
            'CODIGO': codigo_pai,
            'DESCRICAO': descricao_pai,
            'CODIGO_PAI': None,
            'UNIDADE': 'UN', 
            'QTD_NIVEL': 1.0,
            'QTD_TOTAL': 1.0
        })
        
        self.get_components(codigo_pai)
        self.df = pd.DataFrame(self.all_components)
        # Reordenar as colunas
        self.df = self.df[['NIVEL', 'CODIGO', 'CODIGO_PAI', 'DESCRICAO', 'UNIDADE', 'QTD_NIVEL', 'QTD_TOTAL']]
        self.df.to_excel('bom_hierarquica_v6.xlsx', index=False)
        
        self.setup_table()
        self.build_tree()
        
    def setup_table(self):
        self.table.setColumnCount(len(self.df.columns))
        self.table.setHorizontalHeaderLabels(self.df.columns)
        self.populate_table(self.df)
        self.table.resizeColumnsToContents()
        
        # Tornar a tabela não editável
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        
        # Permitir seleção de múltiplas linhas
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        
        # Habilitar a cópia dos dados da tabela
        self.table.setContextMenuPolicy(Qt.ActionsContextMenu)
        copy_action = QAction("Copiar", self.table)
        copy_action.triggered.connect(self.copy_selection)
        self.table.addAction(copy_action)
        
    def copy_selection(self):
        selection = self.table.selectedIndexes()
        if selection:
            rows = sorted(index.row() for index in selection)
            columns = sorted(index.column() for index in selection)
            rowcount = rows[-1] - rows[0] + 1
            colcount = columns[-1] - columns[0] + 1
            table = [[''] * colcount for _ in range(rowcount)]
            for index in selection:
                row = index.row() - rows[0]
                col = index.column() - columns[0]
                table[row][col] = self.table.item(index.row(), index.column()).text()
            stream = io.StringIO()
            csv.writer(stream, delimiter='\t').writerows(table)
            QApplication.clipboard().setText(stream.getvalue())
        
    def populate_table(self, df):
        self.table.setRowCount(len(df))
        for i, row in df.iterrows():
            for j, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                self.table.setItem(i, j, item)
    
    def build_tree(self):
        self.tree.clear()
        
        def build_tree_recursive(parent_item, parent_code, parent_qty):
            query = f"""
            SELECT 
                struct.G1_COD,
                struct.G1_COMP,
                prod.B1_DESC AS 'DESCRICAO',
                struct.G1_XUM, 
                struct.G1_QUANT 
            FROM 
                PROTHEUS12_R27.dbo.SG1010 struct
            INNER JOIN
                PROTHEUS12_R27.dbo.SB1010 prod
            ON 
                G1_COMP = B1_COD AND prod.D_E_L_E_T_ <> '*'
            WHERE 
                G1_COD = '{parent_code}' 
                AND G1_REVFIM <> 'ZZZ' 
                AND struct.D_E_L_E_T_ <> '*' 
                AND G1_REVFIM = (
                    SELECT MAX(G1_REVFIM) 
                    FROM PROTHEUS12_R27.dbo.SG1010 
                    WHERE G1_COD = '{parent_code}'
                    AND G1_REVFIM <> 'ZZZ' 
                    AND struct.D_E_L_E_T_ <> '*'
                )
            """
            
            try:
                df = pd.read_sql(query, self.engine)
                
                for _, row in df.iterrows():
                    total_qty = parent_qty * row['G1_QUANT']
                    
                    child_item = QTreeWidgetItem(parent_item)
                    child_item.setText(0, f"{row['G1_COMP'].strip()}  |  {row['DESCRICAO'].strip()}  |  {total_qty:.2f} {row['G1_XUM'].strip()}")
                    
                    build_tree_recursive(child_item, row['G1_COMP'], total_qty)
            except Exception as e:
                print(f"Erro ao processar código {parent_code}: {str(e)}")
        
        # Buscar descrição do item raiz
        query = f"""
        SELECT B1_DESC as DESCRICAO, B1_UM as UNIDADE
        FROM PROTHEUS12_R27.dbo.SB1010
        WHERE B1_COD = '{self.codigo_pai}'
        AND D_E_L_E_T_ <> '*'
        """
        try:
            df_root = pd.read_sql(query, self.engine)
            root_desc = df_root['DESCRICAO'].iloc[0] if not df_root.empty else ''
            root_unidade = df_root['UNIDADE'].iloc[0] if not df_root.empty else ''
        except:
            root_desc = ''
            root_unidade = ''
        
        # Iniciar a construção da árvore com o nó raiz
        root_item = QTreeWidgetItem(self.tree)
        root_item.setText(0, f"{self.codigo_pai} - {root_desc.strip()} - 1.00 {root_unidade.strip()}")
        build_tree_recursive(root_item, self.codigo_pai, 1.0)
        
        # Expandir todos os itens por padrão
        self.tree.expandAll()
    
    def filter_tables(self):
        filter_codigo = self.filter_input_codigo.text().lower().strip()
        filter_desc = self.filter_input_desc.text().lower().strip()
        
        # Filtrar a tabela
        for row in range(self.table.rowCount()):
            row_visible = True
            codigo = self.table.item(row, 1).text().lower()  # Coluna CODIGO
            descricao = self.table.item(row, 3).text().lower()  # Coluna DESCRICAO
            
            if filter_codigo and filter_codigo not in codigo:
                row_visible = False
            if filter_desc and filter_desc not in descricao:
                row_visible = False
                
            self.table.setRowHidden(row, not row_visible)
        
        # Filtrar a árvore
        def process_item(item):
            text = item.text(0).lower()
            item_visible = True
            
            # Verificar se o item corresponde aos critérios de filtro
            matches_filter = True
            if filter_codigo and filter_codigo not in text:
                matches_filter = False
            if filter_desc and filter_desc not in text:
                matches_filter = False
            
            # Destacar ou limpar o destaque do item
            if matches_filter and (filter_codigo or filter_desc):
                item.setBackground(0, Qt.yellow)
            else:
                item.setBackground(0, Qt.white)
            
            # Processar filhos
            for i in range(item.childCount()):
                child_visible = process_item(item.child(i))
                if child_visible:
                    item_visible = True
            
            item.setHidden(not item_visible)
            return item_visible or matches_filter
        
        for i in range(self.tree.topLevelItemCount()):
            process_item(self.tree.topLevelItem(i))

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

    def toggle_table_visibility(self):
        if self.table.isVisible():
            self.table.hide()
        else:
            self.table.show()

def main():
    app = QApplication(sys.argv)
    viewer = BOMViewer()
    viewer.showMaximized()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
