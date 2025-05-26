import os
import pandas as pd

caminho_excel = r'C:\Users\Rubens\Documents\Tremed\Planilhas Tremed\Tabela_Master.xlsx'

pasta_destino = r'C:\Users\Rubens\Documents\Tremed\FornecedoresAtualização'

aba = 'FORNECEDORES'
coluna_nomes = 'NOME DO FORNECEDOR'

meses = ['JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO',
         'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO']

df = pd.read_excel(caminho_excel, sheet_name=aba)

df.columns = df.columns.str.strip()

nomes_fornecedores = df[coluna_nomes].dropna().unique()

for nome in nomes_fornecedores:
    nome_str = str(nome).strip()
    
    nome_formatado = nome_str.replace(" ", "").upper()
    
    caminho_pasta = os.path.join(pasta_destino, nome_formatado)
    
    os.makedirs(caminho_pasta, exist_ok=True)
    print(f"Pasta criada: {caminho_pasta}")
    
    for mes in meses:
        caminho_mes = os.path.join(caminho_pasta, mes)
        os.makedirs(caminho_mes, exist_ok=True)
        print(f"  Subpasta criada: {caminho_mes}")

print("Processo concluído com sucesso!")
