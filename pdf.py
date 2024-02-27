import pandas as pd
from tika import parser

# Leitura do arquivo fatura.xlsx
df = pd.read_excel('fatura.xlsx')

# Extração dos valores da coluna NumNF
num_nf_values = df['NumNF']

# Leitura do arquivo fatura.pdf
raw_pdf = parser.from_file('fatura.pdf')
pdf_content = raw_pdf['content']

# Busca no PDF
for num_nf in num_nf_values:
    if str(num_nf) in pdf_content:
        # Encontrou o dado no PDF
        # Transformar a página PDF em um arquivo .csv (implementação necessária)
        print(f"Dado {num_nf} encontrado no PDF. Transformando página em arquivo .csv...")
        # Implemente aqui a lógica para criar o arquivo CSV com os dados relevantes

# Caso não encontre o dado no PDF
else:
    print("Dado não encontrado no PDF.")

# Lembre-se de implementar a lógica para transformar a página do PDF em um arquivo .csv
