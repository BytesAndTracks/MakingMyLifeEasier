import os
import pandas as pd

arquivo_abas = r'C:\Users\Rubens\Documents\Tremed\Planilhas Tremed\Master_Fornecedor.xlsx'
arquivo_coluna = r'C:\Users\Rubens\Documents\Tremed\Planilhas Tremed\Tabela_Master.xlsx'

saida_excel = r'C:\Users\Rubens\Downloads\fornecedores_unificados.xlsx'

sheets = pd.ExcelFile(arquivo_abas).sheet_names
fornecedores_abas = [nome.strip() for nome in sheets if nome.strip() != '']  # remove espaços e vazios

df = pd.read_excel(arquivo_coluna, sheet_name='FORNECEDORES')
df.columns = df.columns.str.strip()

if 'NOME DO FORNECEDOR' in df.columns:
    fornecedores_coluna = df['NOME DO FORNECEDOR'].dropna().astype(str).str.strip().unique()
else:
    print("Coluna 'NOME DO FORNECEDOR' não encontrada na aba 'FORNECEDORES'")
    fornecedores_coluna = []

todos_fornecedores = list(fornecedores_abas) + list(fornecedores_coluna)

fornecedores_unicos = sorted(set(todos_fornecedores))

df_saida = pd.DataFrame(fornecedores_unicos, columns=['FORNECEDOR'])

df_saida.to_excel(saida_excel, index=False)

print(f"Arquivo criado com sucesso: {saida_excel}")
