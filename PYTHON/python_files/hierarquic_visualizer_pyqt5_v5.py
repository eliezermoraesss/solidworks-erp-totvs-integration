import pandas as pd
from sqlalchemy import create_engine
from typing import List, Dict
from db_mssql import setup_mssql

class BOMBuilder:
    def __init__(self, engine):
        self.engine = engine
        self.all_components = []
        
    def get_components(self, parent_code: str, parent_qty: float = 1.0, level: int = 1) -> pd.DataFrame:
        """
        Busca recursivamente os componentes de um item e seus subcomponentes.
        
        Args:
            parent_code: Código do item pai
            parent_qty: Quantidade do item pai (para multiplicação acumulativa)
            level: Nível atual na hierarquia
        """
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
            
            # Se encontrou componentes
            if not df.empty:
                for _, row in df.iterrows():
                    # Calcular quantidade acumulada
                    total_qty = parent_qty * row['G1_QUANT']
                    
                    # Adicionar componente atual à lista
                    component = {
                        'NIVEL': level,
                        'CODIGO': row['G1_COMP'],
                        'CODIGO_PAI': row['G1_COD'],
                        'UNIDADE': row['G1_XUM'],
                        'QTD_NIVEL': row['G1_QUANT'],
                        'QTD_TOTAL': total_qty
                    }
                    self.all_components.append(component)
                    
                    # Buscar recursivamente os subcomponentes
                    self.get_components(row['G1_COMP'], total_qty, level + 1)
                    
        except Exception as e:
            print(f"Erro ao processar código {parent_code}: {str(e)}")
    
    def build_bom(self, top_level_code: str) -> pd.DataFrame:
        """
        Constrói a lista completa de materiais hierárquica.
        
        Args:
            top_level_code: Código do item de nível mais alto
        
        Returns:
            DataFrame com a estrutura completa
        """
        # Limpar lista anterior
        self.all_components = []
        
        # Adicionar item de nível mais alto
        self.all_components.append({
            'NIVEL': 0,
            'CODIGO': top_level_code,
            'CODIGO_PAI': None,
            'UNIDADE': 'UN',
            'QTD_NIVEL': 1.0,
            'QTD_TOTAL': 1.0
        })
        
        # Iniciar busca recursiva
        self.get_components(top_level_code)
        
        # Criar DataFrame
        df = pd.DataFrame(self.all_components)
        
        # Ordenar por nível e código
        df = df.sort_values(['NIVEL', 'CODIGO'])
        
        return df

def main():
    # Configuração da conexão (ajuste conforme necessário)
    username, password, database, server = setup_mssql()
    driver = '{SQL Server}'
    conn_str = (f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};'
                f'PWD={password}')
    engine = create_engine(f'mssql+pyodbc:///?odbc_connect={conn_str}')
    
    # Criar construtor de BOM
    bom_builder = BOMBuilder(engine)
    
    # Construir BOM para um código específico
    codigo_pai = 'E7047-001-182'
    df_bom = bom_builder.build_bom(codigo_pai)
    
    # Exibir resultado
    print("\nEstrutura de Materiais Hierárquica:")
    print(df_bom.to_string())
    
    # Opcionalmente, exportar para Excel
    df_bom.to_excel('bom_hierarquica.xlsx', index=False)
    
    engine.dispose()

if __name__ == "__main__":
    main()
