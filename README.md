# Demo Ao Vivo (HTML interativo)

Este painel permite clicar em **mes** e **semana** para atualizar cards e graficos no topo, usando os dados exportados em `powerbi/output`.

## Como abrir rapido (sem servidor)

1. Gere o snapshot local:

```bash
python3 powerbi/exportar_dataset_powerbi.py
python3 dashboard_live/build_inline_data.py
```

2. Abra o arquivo no navegador:

```text
dashboard_live/index.html
```

## Como abrir em modo ao vivo (com CSV atualizado)

Na raiz do projeto:

```bash
python3 powerbi/exportar_dataset_powerbi.py
python3 -m http.server 8502
```

Abra no navegador:

```text
http://localhost:8502/dashboard_live/index.html
```

## O que ja esta dinamico

- Ano, mes e semana com filtros clicaveis.
- Semanas completas do mes (inclui semanas sem venda com valor zero).
- Cards de faturamento, vendas (NF unica), ticket medio.
- Caixa mes e caixa semana (aba Financeiro).
- Top vendedor e top regiao da semana.
- Barras clicaveis por mes e por semana.

## Observacao

- Se abrir direto o `index.html`, ele usa o snapshot `data.inline.js`.
- Se abrir com servidor HTTP na raiz, ele tenta ler os CSVs ao vivo de `powerbi/output`.
