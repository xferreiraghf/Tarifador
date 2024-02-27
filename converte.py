import pandas as pd

caminho_arquivo_txt = r'C:\Users\b11015\Desktop\Tarifador\fatura.txt'

dados = pd.read_csv(caminho_arquivo_txt, sep=';')

dados_agrupados = dados.groupby('NumNF', as_index=False).agg({'Valor': 'sum', 'MesRef': 'first'})

colunas_desejadas = ['NumNF', 'MesRef', 'Valor']
dados_agrupados = dados_agrupados[colunas_desejadas]

caminho_arquivo_excel = 'fatura.xlsx'

dados_agrupados.to_excel(caminho_arquivo_excel, index=False)

print(f'O arquivo Excel foi criado em: {caminho_arquivo_excel}')
