import rarfile

def descompactar_rar(arquivo_rar, diretorio_destino):
    with rarfile.RarFile(arquivo_rar, 'r') as rar_ref:
        rar_ref.extractall(diretorio_destino)
    print("Arquivo RAR descompactado com sucesso.")

nome_arquivo_rar = 'Faturas TIM'
diretorio_destino = 'C:\Users\b11015\Desktop\Tarifador'

descompactar_rar(nome_arquivo_rar, diretorio_destino)
