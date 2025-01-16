import pandas as pd

# Criar o DataFrame com os dados fornecidos
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

df = pd.DataFrame(data)

def build_hierarchy(df, level=1, parent=None):
    result = []
    
    # Filtrar itens do nível atual e com o pai especificado (se houver)
    mask = df['Nivel'] == level
    if parent is not None:
        mask &= df['PAI'] == parent
    
    current_level = df[mask]
    
    for _, row in current_level.iterrows():
        # Adicionar item atual
        result.append({
            'Nivel': row['Nivel'],
            'CÓDIGO': row['CÓDIGO'],
            'PAI': row['PAI'],
            'QUANTIDADE': row['QUANTIDADE'],
            'REVISAO': row['REVISAO']
        })
        
        # Recursivamente adicionar filhos
        children = build_hierarchy(df, level + 1, row['CÓDIGO'])
        result.extend(children)
    
    return result

# Reorganizar a tabela
reorganized = build_hierarchy(df)
result_df = pd.DataFrame(reorganized)

print("\nTabela Reorganizada:")
print(result_df.to_string(index=False))
