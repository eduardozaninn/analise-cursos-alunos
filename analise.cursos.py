import pandas as pd

# Carregar a planilha
arquivo = 'alunos.xlsx'
df = pd.read_excel(arquivo)

# Verificar os nomes das colunas
print(df.columns)

# Filtrar pelo curso de interesse
curso_interesse = "CRIAÇÃO DE MÍDIAS PARA REDES SOCIAIS"
curso_alternativo = "CRIAÇÃO DE MÍDIAS PARA REDES SOCIAIS"

# Verificar o conteúdo da coluna de curso
print(df['CRIAÇÃO DE MÍDIAS PARA REDES SOCIAIS'].unique())
print(df['Dentre as opções qual curso gostaria de fazer?'].unique())

# Filtrar pelas duas colunas
df_filtrado = df[
    (df['CRIAÇÃO DE MÍDIAS PARA REDES SOCIAIS'].fillna('').str.strip().str.upper().str.contains(curso_interesse.upper(), na=False)) |
    (df['Dentre as opções qual curso gostaria de fazer?'].fillna('').str.strip().str.upper().str.contains(curso_alternativo.upper(), na=False)) 
]
# Selecionar as colunas desejadas
df_resultado = df_filtrado[['Nome Completo', 'Whatsapp com DDD (somente números - sem espaço)', 'Carimbo de data/hora']]


# Exibir os dados filtrados
print(df_resultado)

# Salvar o resultado em uma nova planilha
df_resultado.to_excel('dados_filtrados.xlsx', index=False)
