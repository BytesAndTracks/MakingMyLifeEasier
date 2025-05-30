import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
import time
from urllib.parse import quote
import random

# =============================================
# CONFIGURAÇÕES PRINCIPAIS
# =============================================
CAMINHO_PLANILHA = r"C:\Users\Rubens\Downloads\PreencherCnpjSite.xlsx"
COLUNA_FORNECEDOR = "FORNECEDOR"
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0"
]

# =============================================
# FUNÇÕES DE BUSCA ESPECIALIZADAS (SAÚDE)
# =============================================
def limpar_nome_empresa(nome):
    """Remove termos irrelevantes e padroniza o nome para busca"""
    termos_remover = [
        'LTDA', 'EIRELI', 'ME', 'EPP', 'S/A', 'INDÚSTRIA', 'COMÉRCIO',
        'HOSPITALAR', 'MÉDICO', 'SAÚDE', 'PRODUTOS', 'EQUIPAMENTOS'
    ]
    nome_limpo = re.sub(r'[^\w\s]', '', nome).upper()
    for termo in termos_remover:
        nome_limpo = nome_limpo.replace(termo, '')
    return ' '.join(nome_limpo.split())

def buscar_cnpj_saude(nome_empresa):
    """Busca otimizada para empresas de saúde"""
    try:
        # 1. Extrai CNPJ diretamente do nome quando disponível
        cnpj_match = re.search(r'(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})', nome_empresa)
        if cnpj_match:
            return cnpj_match.group(1)
        
        # 2. Consulta simulada à API de saúde (exemplo com CNES)
        nome_formatado = quote(limpar_nome_empresa(nome_empresa))
        
        # Simulação de busca em fonte de dados de saúde
        # (Implementação real exigiria acesso a API específica)
        return "CNPJ não encontrado automaticamente. Sugestão:\n" \
               f"- Consultar manualmente em: https://cnes.datasus.gov.br/pages/estabelecimentos/consulta.jsp\n" \
               f"- Buscar no Google: {nome_empresa} CNPJ"
    
    except Exception as e:
        print(f"Erro na busca de CNPJ: {str(e)}")
        return "Erro na consulta"

def buscar_site_saude(nome_empresa):
    """Busca otimizada para sites de empresas médicas/hospitalares"""
    try:
        headers = {
            "User-Agent": random.choice(USER_AGENTS),
            "Accept-Language": "pt-BR,pt;q=0.9"
        }
        
        # Termos de busca especializados para saúde
        query = f"{nome_empresa} site oficial equipamentos médicos OR hospitalares OR saúde"
        url = f"https://www.google.com/search?q={quote(query)}"
        
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Prioriza resultados com domínios relevantes
        dominios_relevantes = [
            '.com.br', '.med.br', '.hospitalar', 
            'saude', 'medicina', 'hospital'
        ]
        
        for link in soup.find_all('a', href=True):
            href = link['href']
            if '/url?q=' in href and not any(dom in href for dom in ['google.com', 'webcache']):
                site = href.split('/url?q=')[1].split('&')[0]
                if any(dom in site.lower() for dom in dominios_relevantes):
                    return site
        
        # Fallback: busca genérica se não encontrar com termos médicos
        return buscar_site_generico(nome_empresa)
    
    except Exception as e:
        print(f"Erro na busca de site: {str(e)}")
        return "Erro na consulta"

def buscar_site_generico(nome_empresa):
    """Busca de fallback para sites não encontrados com termos médicos"""
    try:
        headers = {"User-Agent": random.choice(USER_AGENTS)}
        query = f"{nome_empresa} site oficial"
        url = f"https://www.google.com/search?q={quote(query)}"
        
        response = requests.get(url, headers=headers, timeout=10)
        soup = BeautifulSoup(response.text, 'html.parser')
        
        for link in soup.find_all('a', href=True):
            if '/url?q=' in link['href']:
                site = link['href'].split('/url?q=')[1].split('&')[0]
                if site.startswith('http'):
                    return site
        
        return "Não encontrado"
    except:
        return "Erro na busca"

# =============================================
# PROCESSAMENTO DA PLANILHA
# =============================================
def processar_planilha():
    try:
        # Carrega os dados
        df = pd.read_excel(CAMINHO_PLANILHA)
        
        # Verifica e cria colunas se necessário
        for col in ['CNPJ', 'SITE']:
            if col not in df.columns:
                df[col] = ""
        
        # Processa cada registro
        total = len(df)
        for index, row in df.iterrows():
            fornecedor = str(row[COLUNA_FORNECEDOR]).strip()
            
            # Pula se já estiver preenchido
            if pd.notna(row['CNPJ']) and str(row['CNPJ']) not in ['', 'Não encontrado']:
                continue
            
            print(f"\nProcessando {index+1}/{total}: {fornecedor}")
            
            # Busca CNPJ com foco em saúde
            df.at[index, 'CNPJ'] = buscar_cnpj_saude(fornecedor)
            
            # Busca site especializado
            df.at[index, 'SITE'] = buscar_site_saude(fornecedor)
            
            # Intervalo aleatório para evitar bloqueio
            time.sleep(random.uniform(2, 5))
            
            # Salvamento incremental a cada 10 registros
            if (index + 1) % 10 == 0:
                df.to_excel(CAMINHO_PLANILHA, index=False)
                print(f"✓ Salvamento temporário realizado (linha {index+1})")
        
        # Salva o resultado final
        df.to_excel(CAMINHO_PLANILHA, index=False)
        print("\n✅ Planilha atualizada com sucesso!")
        
    except Exception as e:
        print(f"\n❌ Erro crítico: {str(e)}")
        print("Verifique o arquivo e tente novamente.")

# =============================================
# EXECUÇÃO PRINCIPAL
# =============================================
if __name__ == "__main__":
    print("""
    ====================================
    BUSCADOR DE CNPJ E SITES - SAÚDE
    ====================================
    """)
    
    processar_planilha()
    
    print("\nOperação concluída. Verifique o arquivo:", CAMINHO_PLANILHA)