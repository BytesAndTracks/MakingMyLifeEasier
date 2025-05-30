import requests
import pandas as pd
from bs4 import BeautifulSoup
import time

API_URL = "https://www.receitaws.com.br/v1/cnpj/"

def consulta_cnpj_receitaws(nome_empresa):
    # A API do ReceitaWS consulta CNPJ por número, não por nome, então essa função só tenta consultar se tiver o CNPJ
    # Portanto, não será possível consultar CNPJ apenas pelo nome com essa API gratuita.
    return None

def busca_site_google(nome_empresa):
    query = nome_empresa + " site oficial"
    url = f"https://www.google.com/search?q={query.replace(' ', '+')}"
    headers = {"User-Agent": "Mozilla/5.0"}
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        soup = BeautifulSoup(response.text, "html.parser")
        
        for g in soup.find_all('div', class_='tF2Cxc'):
            link = g.find('a', href=True)
            if link and 'google' not in link['href']:
                return link['href']
    except Exception as e:
        print(f"Erro buscando site para {nome_empresa}: {e}")
    return None

def main():
    caminho_arquivo = r"C:\Users\Rubens\Downloads\PreencherCnpjSite.xlsx"
    df = pd.read_excel(caminho_arquivo)
    
    if 'SITE' not in df.columns:
        df['SITE'] = None
    if 'CNPJ' not in df.columns:
        df['CNPJ'] = None

    for idx, row in df.iterrows():
        nome = row['FORNECEDOR']
        print(f"Buscando site para: {nome}")
        
        site = busca_site_google(nome)
        df.at[idx, 'SITE'] = site
        
        # CNPJ não preenchido por limitação da API gratuita
        
        time.sleep(5)  # delay para evitar bloqueios
        
    df.to_excel(caminho_arquivo.replace(".xlsx", "_preenchido.xlsx"), index=False)
    print(f"Arquivo salvo em: {caminho_arquivo.replace('.xlsx', '_preenchido.xlsx')}")

if __name__ == "__main__":
    main()
