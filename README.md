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

## Publicar atualizacao no GitHub + Vercel

Se voce alterou o arquivo `Relatório - Em contrução 2026 v4.backup.html`, publique assim no terminal Linux/WSL:

```bash
cd "/home/nic27/repos/Relatório/Painel em contrução"
./publish_report.sh
```

Se quiser, voce tambem pode passar a mensagem do commit:

```bash
./publish_report.sh "Atualiza relatorio abril"
```

O script faz `git add`, `git commit` e `git push origin main`. Como o projeto esta ligado ao GitHub no Vercel, o link publicado deve atualizar automaticamente depois do deploy.
