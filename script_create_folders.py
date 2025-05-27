import os
import pandas as pd
import unicodedata
import re

def remove_acentos(txt):
    nfkd = unicodedata.normalize('NFKD', txt)
    return "".join([c for c in nfkd if not unicodedata.combining(c)])

caminho_excel = r'C:\Users\Rubens\Documents\Tremed\Planilhas Tremed\Tabela_Master.xlsx'

pasta_destino = r'C:\Users\Rubens\Downloads\Fornecedores'

aba = 'FORNECEDORES'
coluna_codigo = 'CÓD'
coluna_nome_fornecedor = 'NOME DO FORNECEDOR'

meses = ['JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO', 
         'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO']

df = pd.read_excel(caminho_excel, sheet_name=aba)

df.columns = df.columns.str.strip()

fornecedores = df[[coluna_codigo, coluna_nome_fornecedor]].dropna()

for _, row in fornecedores.iterrows():
    codigo_int = int(row[coluna_codigo])
    codigo_str = str(codigo_int).zfill(3)

    nome = str(row[coluna_nome_fornecedor])
    nome_sem_acentos = remove_acentos(nome)
    nome_formatado = nome_sem_acentos.upper()
    nome_formatado = re.sub(r'[^A-Z0-9]+', '_', nome_formatado)
    nome_formatado = nome_formatado.strip('_')

    caminho_pasta_fornecedor = os.path.join(pasta_destino, f"{codigo_str}_{nome_formatado}")

    os.makedirs(caminho_pasta_fornecedor, exist_ok=True)
    print(f"Pasta criada: {caminho_pasta_fornecedor}")

    pasta_produtos = os.path.join(caminho_pasta_fornecedor, 'PRODUTOS')
    pasta_cotacoes = os.path.join(caminho_pasta_fornecedor, 'COTAÇÕES')

    os.makedirs(pasta_produtos, exist_ok=True)
    os.makedirs(pasta_cotacoes, exist_ok=True)

    for mes in meses:
        os.makedirs(os.path.join(pasta_cotacoes, mes), exist_ok=True)
        print(f"  Subpasta criada: {os.path.join(pasta_cotacoes, mes)}")

print("Processo concluído com sucesso!")
