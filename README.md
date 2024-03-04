
# Documentação Detalhada

## Resumo

Este script Python é desenvolvido para automatizar o processamento de faturas da empresa TIM. Ele realiza várias tarefas, incluindo a extração de arquivos compactados, conversão de dados para o formato XLSX, extração de informações de CNPJ de arquivos PDF, extração de valores totais de faturas em PDF, integração de dados extraídos e geração de um arquivo final em formato XLSX. O script é estruturado em funções para modularidade e legibilidade.

## Dependências

O script requer as seguintes bibliotecas Python:

-   os
-   patoolib
-   pandas
-   PyPDF2
-   re
-   csv
-   openpyxl
-   shutil

## Funções Principais

### extrair_arquivo_rar(arquivo_rar, pasta_destino)

Esta função descompacta um arquivo RAR especificado para uma pasta de destino. Ela utiliza a biblioteca patoolib para realizar a extração. Após a extração, procura por arquivos TXT na pasta descompactada e retorna o caminho do primeiro arquivo encontrado.

### converter_para_xlsx(caminho_arquivo_txt)

Esta função converte um arquivo TXT em formato CSV para um arquivo XLSX. Ela lê o arquivo CSV usando o pandas, realiza uma agregação dos dados, seleciona as colunas desejadas e salva o resultado em um arquivo XLSX. Em seguida, renomeia o arquivo para "fatura.xlsx".

### extract_cnpj(text)

Esta função extrai números de CNPJ de uma string de texto usando expressões regulares. Retorna uma lista de todos os CNPJs encontrados.

### extract_faturas_cnpj_from_pdf(pdf_path, faturas)

Esta função extrai informações de CNPJ de um arquivo PDF especificado para as faturas fornecidas. Utiliza a biblioteca PyPDF2 para extrair o texto do PDF e expressões regulares para encontrar os CNPJs. Retorna uma lista de tuplas contendo o número da fatura e o CNPJ correspondente.

### extrair_valores(pdf_path)

Esta função extrai os valores totais das faturas de um arquivo PDF especificado. Utiliza a biblioteca PyPDF2 para extrair o texto do PDF e expressões regulares para encontrar os valores totais das faturas. Retorna um dicionário onde as chaves são os números de fatura e os valores são os valores totais correspondentes.

## Funcionalidades Principais

### Processamento Principal

O script principal é executado dentro de uma cláusula `if __name__ == "__main__":`. Ele inicia descompactando um arquivo RAR contendo faturas da TIM, converte o arquivo TXT para XLSX, e então extrai informações de CNPJ e valores totais de faturas em PDF. Em seguida, integra essas informações aos dados extraídos do arquivo XLSX original, gerando um arquivo final consolidado.

### Limpeza de Arquivos Temporários

Após a conclusão bem-sucedida do processamento principal, o script remove arquivos temporários, como o arquivo CSV intermediário e o arquivo XLSX original. Além disso, exclui todos os arquivos da pasta 'Faturas TIM', exceto o arquivo PDF original.

### Saída e Mensagens de Erro

O script fornece mensagens de status durante a execução, informando sobre o progresso e quaisquer problemas encontrados, como a falta de arquivos necessários. Ao finalizar com sucesso, exibe os caminhos dos arquivos gerados.

Esta documentação fornece uma visão geral detalhada do funcionamento e da estrutura do script, permitindo que os usuários entendam facilmente seu propósito e como usá-lo.
