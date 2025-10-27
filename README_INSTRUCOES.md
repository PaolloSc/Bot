# Instruções para Extração de Ementas

## Problema Identificado

O site **https://jurisprudencia.jt.jus.br/jurisprudencia-nacional/pesquisa** é uma aplicação Angular (SPA - Single Page Application) que carrega todo o conteúdo dinamicamente via JavaScript. Isso significa que:

1. O HTML inicial não contém as ementas
2. As ementas são carregadas via chamadas AJAX/API após a página carregar
3. Scrapers tradicionais (requests + BeautifulSoup) não funcionam
4. É necessário um navegador real ou descobrir a API

## Soluções Disponíveis

### Opção 1: Selenium (Requer Chrome/Chromium)

**Arquivo**: `extrair_ementas.py`

```bash
# Instalar dependências
pip install selenium

# Executar
python extrair_ementas.py
```

**Limitação**: Requer Google Chrome ou Chromium instalado e suporte a ChromeDriver.

### Opção 2: Descobrir e usar a API diretamente

**Arquivo**: `extrair_ementas_api.py`

Este script tenta descobrir automaticamente os endpoints da API:

```bash
python extrair_ementas_api.py
```

**Status atual**: A API existe em `/jurisprudencia-nacional/api/pesquisa` mas a documentação não é pública.

### Opção 3: Usar Playwright (Alternativa ao Selenium)

Playwright é mais moderno e pode funcionar melhor:

```bash
pip install playwright
playwright install chromium

python extrair_ementas_playwright.py
```

## Como Descobrir a API Real

1. Abra o navegador e acesse: https://jurisprudencia.jt.jus.br/jurisprudencia-nacional/pesquisa

2. Abra as Ferramentas do Desenvolvedor (F12)

3. Vá para a aba **Network** (Rede)

4. Filtre por **XHR** ou **Fetch**

5. Realize uma busca no site

6. Procure por chamadas que retornem JSON com as ementas

7. Clique na chamada e veja:
   - **URL completa**
   - **Método** (GET/POST)
   - **Headers**
   - **Payload** (se POST)
   - **Resposta** (JSON com as ementas)

8. Adapte o script `extrair_ementas_api.py` com os dados corretos

## Exemplo de Adaptação

Se você descobrir que a API é:

```
POST https://jurisprudencia.jt.jus.br/api/v1/busca
Content-Type: application/json

{
  "termo": "",
  "pagina": 0,
  "tamanho": 20
}
```

Adapte o script:

```python
url = "https://jurisprudencia.jt.jus.br/api/v1/busca"
payload = {
    "termo": "",
    "pagina": 0,
    "tamanho": 20
}

response = requests.post(url, json=payload, headers=headers)
data = response.json()
```

## Formato de Saída

Os scripts salvam as ementas em dois formatos:

### 1. ementas.txt (Texto formatado)
```
================================================================================
EMENTA 1
================================================================================

CABEÇALHO:
[Cabeçalho da ementa]

EMENTA:
[Texto completo da ementa]
```

### 2. ementas.json (JSON estruturado)
```json
[
  {
    "cabecalho": "...",
    "ementa": "..."
  }
]
```

## Arquivos do Projeto

- `extrair_ementas.py` - Script com Selenium (requer Chrome)
- `extrair_ementas_requests.py` - Tentativa com requests (não funciona para este site)
- `extrair_ementas_api.py` - Descoberta de API
- `README_INSTRUCOES.md` - Este arquivo
- `requirements.txt` - Dependências Python

## Suporte

Para sites que exigem JavaScript, as melhores opções são:

1. **Selenium** / **Playwright** - Simulam um navegador real
2. **API REST** - Se conseguir descobrir e documentar
3. **Puppeteer** (Node.js) - Alternativa em JavaScript
