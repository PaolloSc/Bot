#!/usr/bin/env python3
"""
Script alternativo para extrair ementas usando requests + BeautifulSoup.
Compatível com ambientes sem suporte a Selenium.
"""

import requests
from bs4 import BeautifulSoup
import json
import re


def extrair_ementas_requests(url):
    """
    Extrai ementas usando requests e BeautifulSoup.
    
    Args:
        url: URL do site
        
    Returns:
        Lista de dicionários contendo cabeçalho e ementa
    """
    ementas = []
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7',
        'Connection': 'keep-alive',
    }
    
    try:
        print(f"Acessando {url}...")
        response = requests.get(url, headers=headers, timeout=30)
        print(f"Status: {response.status_code}")
        
        if response.status_code == 200:
            # Salva HTML para debug
            with open("pagina.html", "w", encoding="utf-8") as f:
                f.write(response.text)
            print("HTML salvo em 'pagina.html'")
            
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # Tenta diferentes seletores para encontrar ementas
            seletores_possiveis = [
                {'name': 'div', 'class_': 'ementa'},
                {'name': 'div', 'class_': 'resultado'},
                {'name': 'div', 'class_': 'acordao'},
                {'name': 'article'},
                {'name': 'div', 'class_': 'item'},
            ]
            
            elementos_encontrados = []
            for seletor in seletores_possiveis:
                if 'class_' in seletor:
                    elementos = soup.find_all(seletor['name'], class_=re.compile(seletor['class_']))
                else:
                    elementos = soup.find_all(seletor['name'])
                
                if elementos:
                    elementos_encontrados = elementos
                    print(f"Encontrados {len(elementos)} elementos com seletor: {seletor}")
                    break
            
            # Se não encontrou com os seletores específicos, procura por padrões de texto
            if not elementos_encontrados:
                print("Tentando buscar padrões de ementa no texto...")
                # Procura por elementos que contenham palavras-chave de ementas
                palavras_chave = ['RECURSO', 'EMENTA', 'ACÓRDÃO', 'PROCESSO', 'DECISÃO']
                for tag in soup.find_all(['div', 'p', 'article', 'section']):
                    texto = tag.get_text().strip()
                    if any(palavra in texto.upper() for palavra in palavras_chave) and len(texto) > 100:
                        elementos_encontrados.append(tag)
            
            # Extrai dados
            for idx, elemento in enumerate(elementos_encontrados[:20]):
                try:
                    texto = elemento.get_text().strip()
                    if texto and len(texto) > 50:  # Ignora textos muito curtos
                        # Tenta separar cabeçalho da ementa
                        linhas = [l.strip() for l in texto.split('\n') if l.strip()]
                        if len(linhas) > 1:
                            cabecalho = linhas[0]
                            ementa = '\n'.join(linhas[1:])
                        else:
                            cabecalho = f"Ementa {idx + 1}"
                            ementa = texto
                        
                        ementas.append({
                            "cabecalho": cabecalho,
                            "ementa": ementa
                        })
                        print(f"Extraída ementa {idx + 1}")
                except Exception as e:
                    print(f"Erro ao processar elemento {idx}: {e}")
                    continue
            
            print(f"\nTotal de ementas extraídas: {len(ementas)}")
            
        elif response.status_code == 403:
            print("Erro 403: Acesso bloqueado pelo servidor.")
            print("O site pode estar bloqueando scrapers ou requer interação JavaScript.")
        else:
            print(f"Erro ao acessar o site: Status {response.status_code}")
            
    except requests.exceptions.RequestException as e:
        print(f"Erro de conexão: {e}")
    except Exception as e:
        print(f"Erro durante a extração: {e}")
    
    return ementas


def salvar_ementas(ementas, formato="txt"):
    """
    Salva as ementas extraídas em arquivo.
    
    Args:
        ementas: Lista de dicionários com ementas
        formato: Formato do arquivo (txt, json)
    """
    if formato == "txt":
        with open("ementas.txt", "w", encoding="utf-8") as f:
            for i, item in enumerate(ementas, 1):
                f.write(f"{'='*80}\n")
                f.write(f"EMENTA {i}\n")
                f.write(f"{'='*80}\n\n")
                f.write(f"CABEÇALHO:\n{item['cabecalho']}\n\n")
                f.write(f"EMENTA:\n{item['ementa']}\n\n")
        print("Ementas salvas em 'ementas.txt'")
    
    elif formato == "json":
        with open("ementas.json", "w", encoding="utf-8") as f:
            json.dump(ementas, f, ensure_ascii=False, indent=2)
        print("Ementas salvas em 'ementas.json'")


def main():
    url = "https://jurisprudencia.jt.jus.br/jurisprudencia-nacional/pesquisa"
    
    print("Iniciando extração de ementas (versão requests)...")
    ementas = extrair_ementas_requests(url)
    
    if ementas:
        salvar_ementas(ementas, formato="txt")
        salvar_ementas(ementas, formato="json")
        print("\nExtração concluída com sucesso!")
    else:
        print("\nNenhuma ementa foi extraída.")
        print("O site pode exigir JavaScript ou autenticação.")
        print("Verifique o arquivo 'pagina.html' para análise da estrutura.")


if __name__ == "__main__":
    main()
