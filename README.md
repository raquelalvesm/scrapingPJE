# Scrapping PJe

Automação para extrair dados de processos do PJe a partir de uma lista de números de processo em Excel.

## Requisitos

- Python 3.11+ (recomendado)
- Chrome instalado

## Instalação

```bash
pip install -r requirements.txt
```

## Uso

1. Coloque os números de processo em `Pasta1.xlsx`, na coluna `nr_processo`.
2. Ajuste o arquivo `scrapping_pje.py` se necessário para os caminhos locais.
3. Execute:

```bash
python scrapping_pje.py
```

4. Faça login manual no PJe quando o Chrome abrir.
5. O resultado será salvo em `resultado.xlsx`.

## Dependências

- `openpyxl`
- `selenium`
- `webdriver_manager`
- `pyperclip`
