import pandas as pd

path_hospitalar = r"C:\Users\Rubens\Downloads\FORNECEDORES_HOSPITALAR_V1.xlsx"
path_master = r"C:\Users\Rubens\Documents\Tremed\Planilhas Tremed\Tabela_Master_formatado.xlsx"

df_hospitalar = pd.read_excel(path_hospitalar, usecols="A")
fornecedores_hospitalar = df_hospitalar.iloc[:,0].dropna().astype(str).str.strip()

df_master = pd.read_excel(path_master, sheet_name="FORNECEDORES")
fornecedores_master = df_master.iloc[:,5].dropna().astype(str).str.strip()

fornecedores_faltantes = fornecedores_hospitalar[~fornecedores_hospitalar.isin(fornecedores_master)].unique()

print(f"Encontrados {len(fornecedores_faltantes)} fornecedores faltantes.")

nova_coluna_f = fornecedores_master.tolist() + list(fornecedores_faltantes)

max_linhas = max(len(df_master), len(nova_coluna_f))
df_master_expanded = df_master.reindex(range(max_linhas))

df_master_expanded.iloc[:len(nova_coluna_f), 5] = nova_coluna_f

path_saida = r"C:\Users\Rubens\Documents\Tremed\Planilhas Tremed\Tabela_Master_formatado_atualizado.xlsx"
df_master_expanded.to_excel(path_saida, sheet_name="FORNECEDORES", index=False)

print(f"Arquivo atualizado salvo em: {path_saida}")
