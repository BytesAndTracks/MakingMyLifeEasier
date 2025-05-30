import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
import time
from urllib.parse import quote

# Configurações
CAMINHO_PLANILHA = r"C:\Users\Rubens\Downloads\PreencherCnpjSite.xlsx"
COLUNA_FORNECEDOR = "FORNECEDOR"  # Nome da coluna com os fornecedores

# --- Função para buscar CNPJ (simulada, pois não há API pública para busca por nome) ---
def buscar_cnpj(nome_empresa):
    """
    Simula uma busca por CNPJ com base no nome (limitações: APIs públicas não permitem busca direta por nome).
    Alternativas: 
    - Usar serviços pagos como CNPJ WS Pro ou Serpro.
    - Extrair CNPJ do nome quando disponível (ex: "EMPRESA (12.345.678/0001-90)").
    """
    try:
        # Extrai CNPJ se estiver no nome (ex: "EMPRESA (12.345.678/0001-90)")
        cnpj_no_nome = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", nome_empresa)
        if cnpj_no_nome:
            return cnpj_no_nome.group(1)
        
        # Se não houver CNPJ no nome, retorna "Não encontrado" (sem API válida para busca por nome)
        return "Não encontrado (busca manual necessária)"
    
    except Exception as e:
        print(f"Erro ao buscar CNPJ para {nome_empresa}: {str(e)}")
        return "Erro"

# --- Função para buscar site oficial via Google ---
def buscar_site(nome_empresa):
    """
    Faz uma busca no Google pelo "nome_empresa site oficial" e retorna o primeiro link válido.
    """
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        }
        query = f"{nome_empresa} site oficial"
        url = f"https://www.google.com/search?q={quote(query)}"
        
        response = requests.get(url, headers=headers, timeout=10)
        soup = BeautifulSoup(response.text, "html.parser")
        
        # Encontra o primeiro link não-Google nos resultados
        for link in soup.find_all("a", href=True):
            href = link["href"]
            if "url?q=" in href and "google.com" not in href:
                site = href.split("url?q=")[1].split("&")[0]
                if site.startswith("http"):
                    return site
        
        return "Não encontrado"
    
    except Exception as e:
        print(f"Erro ao buscar site para {nome_empresa}: {str(e)}")
        return "Erro"

# --- Processamento da planilha ---
def processar_planilha():
    # Carrega a planilha
    df = pd.read_excel(CAMINHO_PLANILHA)
    
    # Verifica se as colunas CNPJ e SITE existem, se não, cria
    if "CNPJ" not in df.columns:
        df["CNPJ"] = ""
    if "SITE" not in df.columns:
        df["SITE"] = ""
    
    # Itera sobre cada linha para preencher CNPJ e Site
    for index, row in df.iterrows():
        fornecedor = row[COLUNA_FORNECEDOR]
        
        # Pula se já estiver preenchido
        if pd.notna(row["CNPJ"]) and row["CNPJ"] != "":
            continue
        
        print(f"Processando: {fornecedor}")
        
        # Busca CNPJ
        df.at[index, "CNPJ"] = buscar_cnpj(fornecedor)
        
        # Busca Site
        df.at[index, "SITE"] = buscar_site(fornecedor)
        
        # Delay para evitar bloqueio (Google pode bloquear muitas requisições)
        time.sleep(2)
    
    # Salva a planilha atualizada
    df.to_excel(CAMINHO_PLANILHA, index=False)
    print("Planilha atualizada com sucesso!")

# Executa o script
if __name__ == "__main__":
    processar_planilha()