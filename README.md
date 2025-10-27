# Extrator de Ementas - Jurisprudência Trabalhista

Script para extrair ementas do site de jurisprudência da Justiça do Trabalho.

## Instalação

```bash
pip install -r requirements.txt
```

## Uso

```bash
python extrair_ementas.py
```

O script irá:
1. Acessar o site de jurisprudência
2. Extrair as ementas disponíveis
3. Salvar em dois formatos:
   - `ementas.txt` - Formato texto legível
   - `ementas.json` - Formato JSON estruturado

## Formato dos Dados

Cada ementa contém:
- **Cabeçalho**: Identificação da ementa
- **Ementa**: Texto completo da ementa

## Requisitos

- Python 3.7+
- Google Chrome/Chromium instalado
- Selenium WebDriver

## Observações

- O script salva automaticamente o HTML da página para debug caso não encontre ementas
- Screenshots de erro são salvos automaticamente em caso de falha
- Os arquivos de output são ignorados pelo Git (ver `.gitignore`)
