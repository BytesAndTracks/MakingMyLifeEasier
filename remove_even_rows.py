import pandas as pd

input_file = r"C:\Users\Rubens\Downloads\FORNECEDORES_HOSPITALAR.xlsx"
output_file = r"C:\Users\Rubens\Downloads\FORNECEDORES_HOSPITALAR_sem_duplicatas.xlsx"

df = pd.read_excel(input_file, dtype=str)
df.fillna('', inplace=True)

indices_to_keep = []

for i in range(len(df)):
    current_name = df.iloc[i, 0].strip().upper()
    
    if i == 0:
        indices_to_keep.append(i)
        continue

    previous_name = df.iloc[i-1, 0].strip().upper()
    
    # Verifica se previous_name é substring do current_name
    # e se current_name é maior que previous_name (tem sufixo extra)
    if previous_name and previous_name in current_name and len(current_name) > len(previous_name):
        # Linha duplicada, não mantém
        continue
    else:
        indices_to_keep.append(i)

df_filtered = df.iloc[indices_to_keep].reset_index(drop=True)
df_filtered.to_excel(output_file, index=False)

print(f"Arquivo salvo sem duplicatas: {output_file}")
