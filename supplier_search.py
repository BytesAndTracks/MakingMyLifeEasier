import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
import time
import random
import os
from tqdm import tqdm
from urllib.parse import quote

# =============================================
# CONFIGURAÇÕES PRINCIPAIS
# =============================================
ARQUIVO_ORIGINAL = r"C:\Users\Rubens\Downloads\PreencherCnpjSite.xlsx"
ARQUIVO_TEMP = r"C:\Users\Rubens\Downloads\PreencherCnpjSite_TEMP.xlsx"
COLUNA_FORNECEDOR = "FORNECEDOR"
CHECKPOINT_FREQ = 100  # Salva a cada 100 registros

# =============================================
# FUNÇÕES AUXILIARES
# =============================================
def get_random_agent():
    agents = [
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0"
    ]
    return random.choice(agents)

def limpar_nome(nome):
    remove = ['LTDA', 'EIRELI', 'ME', 'EPP', 'S/A', 'INDÚSTRIA', 'COMÉRCIO', 'HOSPITALAR']
    nome = re.sub(r'[^\w\s]', '', str(nome)).upper()
    for termo in remove:
        nome = nome.replace(termo, '')
    return ' '.join(nome.split())

# =============================================
# FUNÇÕES DE BUSCA
# =============================================
def buscar_cnpj(nome):
    try:
        # Extrai CNPJ do nome se existir
        match = re.search(r'(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})', nome)
        if match:
            return match.group(1)
        
        # Lógica simulada para busca especializada
        return "Consultar em: https://cnes.datasus.gov.br/"
    
    except Exception as e:
        print(f"\nErro CNPJ: {str(e)}")
        return "Erro"

def buscar_site(nome):
    try:
        headers = {'User-Agent': get_random_agent()}
        query = f"{nome} site oficial equipamentos médicos OR hospitalares"
        url = f"https://www.google.com/search?q={quote(query)}"
        
        response = requests.get(url, headers=headers, timeout=15)
        soup = BeautifulSoup(response.text, 'html.parser')
        
        for link in soup.find_all('a', href=True):
            if '/url?q=' in link['href']:
                site = link['href'].split('/url?q=')[1].split('&')[0]
                if any(dom in site for dom in ['.com.br', '.med.br', 'hospital']):
                    return site
        
        return buscar_site_generico(nome)
    
    except Exception as e:
        print(f"\nErro site: {str(e)}")
        return "Erro"

def buscar_site_generico(nome):
    try:
        headers = {'User-Agent': get_random_agent()}
        url = f"https://www.google.com/search?q={quote(nome + ' site oficial')}"
        response = requests.get(url, headers=headers, timeout=10)
        soup = BeautifulSoup(response.text, 'html.parser')
        
        for link in soup.find_all('a', href=True):
            if '/url?q=' in link['href']:
                return link['href'].split('/url?q=')[1].split('&')[0]
        
        return "Não encontrado"
    except:
        return "Erro"

# =============================================
# CONTROLE DE PROCESSAMENTO
# =============================================
def carregar_progresso():
    """Verifica se há um arquivo temporário para continuar"""
    if os.path.exists(ARQUIVO_TEMP):
        df = pd.read_excel(ARQUIVO_TEMP)
        inicio = df[df['CNPJ'].isna() | (df['CNPJ'] == '')].index[0]
        print(f"\n♻ Retomando processamento da linha {inicio + 1}...")
        return df, inicio
    else:
        df = pd.read_excel(ARQUIVO_ORIGINAL)
        df['CNPJ'] = df['CNPJ'].fillna('')
        df['SITE'] = df['SITE'].fillna('')
        print("\n✅ Iniciando novo processamento...")
        return df, 0

def salvar_checkpoint(df, index):
    """Salva progresso atual"""
    df.to_excel(ARQUIVO_TEMP, index=False)
    print(f"\n💾 Checkpoint salvo (linha {index + 1})")

def finalizar_processo(df):
    """Finaliza e limpa arquivos temporários"""
    df.to_excel(ARQUIVO_ORIGINAL, index=False)
    if os.path.exists(ARQUIVO_TEMP):
        os.remove(ARQUIVO_TEMP)
    print("\n✅ Processo concluído com sucesso!")

# =============================================
# EXECUÇÃO PRINCIPAL
# =============================================
def main():
    print("""
    ================================
    BUSCADOR CNPJ/SITES - RAMO MÉDICO
    ================================
    """)
    
    df, inicio = carregar_progresso()
    total = len(df)
    
    try:
        with tqdm(total=total, initial=inicio, desc="Progresso") as pbar:
            for index in range(inicio, total):
                fornecedor = str(df.at[index, COLUNA_FORNECEDOR]).strip()
                
                if pd.isna(df.at[index, 'CNPJ']) or df.at[index, 'CNPJ'] == '':
                    df.at[index, 'CNPJ'] = buscar_cnpj(fornecedor)
                    df.at[index, 'SITE'] = buscar_site(fornecedor)
                    time.sleep(random.uniform(1, 3))
                
                # Atualiza barra de progresso
                pbar.update(1)
                
                # Checkpoint periódico
                if (index + 1) % CHECKPOINT_FREQ == 0:
                    salvar_checkpoint(df, index)
        
        finalizar_processo(df)
    
    except Exception as e:
        print(f"\n❌ Erro interrompeu o processamento: {str(e)}")
        print(f"Última linha processada: {index + 1}")
        if 'df' in locals():
            salvar_checkpoint(df, index)
        print("\n⚠ Execute novamente para continuar de onde parou")

if __name__ == "__main__":
    main()