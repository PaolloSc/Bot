#!/usr/bin/env python3
"""
Script para extrair ementas tentando identificar as chamadas de API.
"""

import requests
import json
from datetime import datetime


def tentar_apis_conhecidas():
    """
    Tenta acessar endpoints de API conhecidos de sistemas de jurisprudência.
    """
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Accept': 'application/json, text/plain, */*',
        'Accept-Language': 'pt-BR,pt;q=0.9',
        'Content-Type': 'application/json',
    }
    
    # Possíveis endpoints de API
    base_url = "https://jurisprudencia.jt.jus.br"
    endpoints = [
        "/api/pesquisa",
        "/api/jurisprudencia",
        "/api/acordaos",
        "/api/ementas",
        "/api/consulta",
        "/jurisprudencia-nacional/api/pesquisa",
        "/jurisprudencia-nacional/api/acordaos",
    ]
    
    print("Tentando descobrir endpoints de API...\n")
    
    for endpoint in endpoints:
        url = f"{base_url}{endpoint}"
        try:
            print(f"Testando: {url}")
            
            # Tenta GET
            response = requests.get(url, headers=headers, timeout=10)
            print(f"  GET - Status: {response.status_code}")
            
            if response.status_code == 200:
                try:
                    data = response.json()
                    print(f"  ✓ JSON válido recebido!")
                    print(f"  Keys: {list(data.keys()) if isinstance(data, dict) else 'Lista'}")
                    return url, data
                except:
                    print(f"  Resposta não é JSON")
            
            # Tenta POST com payload vazio
            if response.status_code in [404, 405]:
                response = requests.post(url, headers=headers, json={}, timeout=10)
                print(f"  POST - Status: {response.status_code}")
                
                if response.status_code == 200:
                    try:
                        data = response.json()
                        print(f"  ✓ JSON válido recebido via POST!")
                        return url, data
                    except:
                        pass
                        
        except requests.exceptions.Timeout:
            print(f"  Timeout")
        except requests.exceptions.RequestException as e:
            print(f"  Erro: {e}")
        
        print()
    
    return None, None


def criar_payload_busca():
    """
    Cria payloads de busca típicos para sistemas de jurisprudência.
    """
    payloads = [
        {
            "termo": "",
            "pagina": 1,
            "tamanhoPagina": 20
        },
        {
            "query": "",
            "page": 1,
            "size": 20
        },
        {
            "pesquisa": {
                "termo": "*",
                "pagina": 0,
                "itensPorPagina": 20
            }
        },
        {
            "filtros": {},
            "pagina": 1,
            "quantidade": 20
        }
    ]
    return payloads


def buscar_com_payload():
    """
    Tenta realizar buscas com diferentes payloads.
    """
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Accept': 'application/json',
        'Content-Type': 'application/json',
    }
    
    base_url = "https://jurisprudencia.jt.jus.br"
    endpoints_post = [
        "/api/pesquisa",
        "/jurisprudencia-nacional/api/pesquisa",
        "/api/busca",
    ]
    
    payloads = criar_payload_busca()
    
    print("\nTentando buscas com payloads...\n")
    
    for endpoint in endpoints_post:
        url = f"{base_url}{endpoint}"
        for i, payload in enumerate(payloads):
            try:
                print(f"POST {url}")
                print(f"Payload {i+1}: {json.dumps(payload, indent=2)}")
                
                response = requests.post(url, headers=headers, json=payload, timeout=15)
                print(f"Status: {response.status_code}")
                
                if response.status_code == 200:
                    try:
                        data = response.json()
                        print(f"✓ Sucesso! JSON recebido")
                        print(f"Keys: {list(data.keys()) if isinstance(data, dict) else 'Lista'}")
                        return url, data, payload
                    except:
                        print(f"Resposta não é JSON")
                        
            except Exception as e:
                print(f"Erro: {e}")
            
            print()
    
    return None, None, None


def main():
    print("="*80)
    print("TENTATIVA DE DESCOBERTA DE API - JURISPRUDÊNCIA TRABALHISTA")
    print("="*80)
    print()
    
    # Tenta descobrir endpoints
    url, data = tentar_apis_conhecidas()
    
    if data:
        print(f"\n{'='*80}")
        print("API ENCONTRADA!")
        print(f"{'='*80}")
        print(f"URL: {url}")
        print(f"\nDados recebidos:")
        print(json.dumps(data, indent=2, ensure_ascii=False)[:1000])
        
        with open("api_resposta.json", "w", encoding="utf-8") as f:
            json.dump({"url": url, "data": data}, f, indent=2, ensure_ascii=False)
        print(f"\nResposta salva em 'api_resposta.json'")
    else:
        print("\nNenhuma API descoberta com GET. Tentando POST...")
        url, data, payload = buscar_com_payload()
        
        if data:
            print(f"\n{'='*80}")
            print("API ENCONTRADA COM POST!")
            print(f"{'='*80}")
            print(f"URL: {url}")
            print(f"Payload: {json.dumps(payload, indent=2)}")
            print(f"\nDados recebidos:")
            print(json.dumps(data, indent=2, ensure_ascii=False)[:1000])
            
            with open("api_resposta.json", "w", encoding="utf-8") as f:
                json.dump({"url": url, "payload": payload, "data": data}, f, indent=2, ensure_ascii=False)
            print(f"\nResposta salva em 'api_resposta.json'")
        else:
            print("\n" + "="*80)
            print("CONCLUSÃO")
            print("="*80)
            print("""
O site usa Angular e carrega dados dinamicamente via JavaScript.
Não foi possível descobrir a API automaticamente.

PRÓXIMOS PASSOS:

1. Use as ferramentas de desenvolvedor do navegador (F12)
2. Acesse: https://jurisprudencia.jt.jus.br/jurisprudencia-nacional/pesquisa
3. Vá para a aba "Network" (Rede)
4. Realize uma busca no site
5. Procure por chamadas XHR/Fetch para identificar a API real
6. Adapte o script com a URL e payload corretos

ALTERNATIVA:
- Use Selenium/Playwright em ambiente com suporte a navegadores
- Configure um navegador headless com ChromeDriver
            """)


if __name__ == "__main__":
    main()
