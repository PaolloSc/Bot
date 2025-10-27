#!/usr/bin/env python3
"""
Script para extrair ementas do site de jurisprudência trabalhista.
"""

import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import json


def setup_driver():
    """Configura e retorna o driver do Selenium."""
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
    
    driver = webdriver.Chrome(options=chrome_options)
    return driver


def extrair_ementas(url, num_paginas=1):
    """
    Extrai ementas do site de jurisprudência.
    
    Args:
        url: URL do site
        num_paginas: Número de páginas a extrair
        
    Returns:
        Lista de dicionários contendo cabeçalho e ementa
    """
    driver = setup_driver()
    ementas = []
    
    try:
        print(f"Acessando {url}...")
        driver.get(url)
        
        # Aguarda carregamento da página
        wait = WebDriverWait(driver, 20)
        
        # Tenta encontrar os elementos de resultados
        time.sleep(5)  # Aguarda carregamento do JavaScript
        
        # Procura por elementos de ementa (ajustar seletores conforme necessário)
        # Tentando diferentes seletores comuns em sites de jurisprudência
        possivel_seletores = [
            "//div[contains(@class, 'ementa')]",
            "//div[contains(@class, 'resultado')]",
            "//div[contains(@class, 'acordao')]",
            "//article",
            "//div[contains(@class, 'item')]",
        ]
        
        elementos_encontrados = []
        for seletor in possivel_seletores:
            try:
                elementos = driver.find_elements(By.XPATH, seletor)
                if elementos:
                    elementos_encontrados = elementos
                    print(f"Encontrados {len(elementos)} elementos com seletor: {seletor}")
                    break
            except:
                continue
        
        if not elementos_encontrados:
            print("Nenhum elemento encontrado. Salvando HTML da página para análise...")
            with open("pagina.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            print("HTML salvo em 'pagina.html'")
        
        # Extrai dados de cada elemento encontrado
        for idx, elemento in enumerate(elementos_encontrados[:20]):  # Limita a 20 primeiros
            try:
                texto = elemento.text.strip()
                if texto:
                    # Tenta separar cabeçalho da ementa
                    linhas = texto.split('\n')
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
        
    except Exception as e:
        print(f"Erro durante a extração: {e}")
        # Salva screenshot para debug
        try:
            driver.save_screenshot("erro.png")
            print("Screenshot salvo em 'erro.png'")
        except:
            pass
    
    finally:
        driver.quit()
    
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
    
    print("Iniciando extração de ementas...")
    ementas = extrair_ementas(url, num_paginas=1)
    
    if ementas:
        salvar_ementas(ementas, formato="txt")
        salvar_ementas(ementas, formato="json")
        print("\nExtração concluída com sucesso!")
    else:
        print("\nNenhuma ementa foi extraída.")
        print("Verifique o arquivo 'pagina.html' para análise da estrutura do site.")


if __name__ == "__main__":
    main()
