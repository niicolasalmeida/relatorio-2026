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

O script faz tudo em um comando:

- `git add`
- `git commit`
- `git push origin main`
- `npx vercel --prod --yes`

Se nao houver mudancas locais para commit, ele pula a parte do Git e executa somente o deploy em producao no Vercel.

## Importacao automatica para SQLite

O projeto agora inclui `sync_sqlite_imports.py`, que monitora e importa para SQLite os arquivos:

- `faturamento.xlsx`
- `dre.xlsx`
- `estoque.xlsx`
- `fluxo.xlsx`
- `clientes.xlsx`

Ele cria tabelas por aba, guarda metadados de sincronizacao e aplica apenas `insert/update/delete` nas linhas alteradas quando um arquivo for substituido.
Ao final de cada atualizacao com mudancas reais, ele recalcula automaticamente a tabela `Insights`.

Categorias geradas na tabela `Insights`:

- `maiores_crescimentos`
- `maiores_quedas`
- `estoque_parado`
- `fluxo`
- `margem`
- `desvios`
- `oportunidades`

Execucao unica:

```bash
python3 sync_sqlite_imports.py
```

Monitoramento continuo:

```bash
python3 sync_sqlite_imports.py --watch
```

Banco em caminho customizado:

```bash
python3 sync_sqlite_imports.py --db ./dados/relatorio.db
```

## API FastAPI somente leitura

Existe uma API Python em `api_fastapi.py` para consultar o banco SQLite importado.

Instalacao automatica recomendada:

```bash
python3 setup.py
```

O instalador faz automaticamente:

- cria `.venv`
- instala `requirements.txt`
- cria ou preserva `.env`
- cria o SQLite
- importa os Excel encontrados
- valida a API com uma chamada HTTP local
- executa testes de sanidade
- mostra um resumo final com `VERDE`, `AMARELO` e `VERMELHO`

Observacao importante:

- para a importacao completa, o projeto espera os arquivos `faturamento.xlsx`, `dre.xlsx`, `estoque.xlsx`, `fluxo.xlsx` e `clientes.xlsx` na raiz
- se apenas `Base.xlsx` existir, o instalador usa essa planilha como fallback e cria aliases de dataset no SQLite

Instalacao minima:

```bash
cp .env.example .env
python3 -m pip install -r requirements.txt
```

Configuracao inicial do `.env`:

- `SQLITE_DB_PATH=./relatorio.db`
- `OPENAI_API_KEY=sk-sua-chave`
- `OPENAI_MODEL=gpt-5-mini`

Subir a API localmente:

```bash
python3 -m uvicorn api_fastapi:app --reload
```

Se o banco estiver em outro caminho:

```bash
SQLITE_DB_PATH=./dados/relatorio.db python3 -m uvicorn api_fastapi:app --reload
```

Endpoints disponiveis:

- `GET /faturamento`
- `GET /dre`
- `GET /estoque`
- `GET /fluxo`
- `GET /clientes`
- `GET /indicadores`
- `GET /insights`
- `GET /morning-report`
- `POST /chat`

Os endpoints de dataset aceitam:

- `include_rows=true|false`
- `limit`
- `offset`

Fluxo do endpoint `POST /chat`:

1. interpreta a intencao com regras deterministicas
2. escolhe apenas os datasets e insights necessarios
3. consulta somente essas tabelas no SQLite
4. monta um contexto enxuto
5. envia esse contexto para a OpenAI
6. retorna a resposta

Variaveis de ambiente:

- `OPENAI_API_KEY`
- `OPENAI_MODEL` opcional, padrao `gpt-5-mini`
- `SQLITE_DB_PATH` opcional
- `LOG_LEVEL` opcional, padrao `INFO`

## Experiencia de primeira execucao

Ao iniciar a API:

- o arquivo `.env` e carregado automaticamente quando existir
- o caminho do SQLite e validado automaticamente
- o SQLite e criado automaticamente se ainda nao existir
- as tabelas-base `sync_files`, `sync_sheets` e `Insights` sao criadas automaticamente

Importante:

- criar o arquivo `relatorio.db` nao significa que os dados ja foram importados
- se o banco estiver vazio, a API responde com diagnostico e o comando exato para importar as planilhas

## Como configurar a OPENAI_API_KEY

Se `/chat` responder que a chave nao esta configurada:

1. copie o arquivo de exemplo:

```bash
cp .env.example .env
```

2. edite `.env` e preencha:

```env
OPENAI_API_KEY=sk-sua-chave-aqui
```

3. reinicie a API:

```bash
python3 -m uvicorn api_fastapi:app --reload
```

Alternativa temporaria no terminal:

```bash
export OPENAI_API_KEY="sk-sua-chave-aqui"
python3 -m uvicorn api_fastapi:app --reload
```

## Como criar e popular o SQLite

Se o banco nao existir, a API cria automaticamente um SQLite vazio no caminho configurado.

Para carregar dados reais:

1. coloque estes arquivos na raiz do projeto:

- `faturamento.xlsx`
- `dre.xlsx`
- `estoque.xlsx`
- `fluxo.xlsx`
- `clientes.xlsx`

2. execute:

```bash
python3 sync_sqlite_imports.py
```

Se voce tiver apenas `Base.xlsx`, tambem funciona:

```bash
python3 sync_sqlite_imports.py Base.xlsx
```

Nesse modo, o importador mapeia automaticamente as abas:

- `Base` -> `faturamento`
- `Despesas` -> `dre`
- `Estoque` -> `estoque`
- `Recebimento` -> `fluxo`
- `Financeiro` -> `clientes`

3. se quiser gravar em outro caminho:

```bash
python3 sync_sqlite_imports.py --db ./dados/relatorio.db
```

ou configure no `.env`:

```env
SQLITE_DB_PATH=./dados/relatorio.db
```

## Diagnostico e mensagens de erro

A aplicacao evita erros genericos.

Quando algo falhar, a API responde com:

- `error`: codigo curto do problema
- `diagnostic`: causa exata
- `solution`: o que fazer
- `details`: contexto tecnico complementar

Endpoints de apoio:

- `GET /debug/integration`
- `POST /debug/frontend-log`

Diagnostico local completo:

```bash
python3 doctor.py
```

O `doctor.py` verifica automaticamente:

- Python
- FastAPI
- Uvicorn
- SQLite
- `OPENAI_API_KEY`
- banco SQLite
- tabelas existentes
- arquivos Excel
- permissoes
- API
- dashboard

Quando houver erro ou pendencia, ele mostra exatamente como corrigir.

Exemplo:

```bash
curl -X POST http://127.0.0.1:8000/chat \
  -H "Content-Type: application/json" \
  -d '{"question":"Quais sao os principais riscos financeiros atuais?","max_rows_per_sheet":150}'
```

Relatorio executivo automatico:

```bash
curl http://127.0.0.1:8000/morning-report
```

## Subir a API sem FastAPI

Se o seu Python estiver sem `pip`, `FastAPI` ou `uvicorn`, voce ainda pode usar o HTML com o servidor local em Python puro:

```bash
python3 api_local.py
```

Por padrao ele sobe em:

```text
http://127.0.0.1:8000
```

Rotas suportadas nesse modo:

- `GET /`
- `GET /debug/integration`
- `GET /morning-report`
- `GET /insights`
- `GET /faturamento`
- `GET /dre`
- `GET /estoque`
- `GET /fluxo`
- `GET /clientes`
- `POST /debug/frontend-log`
- `POST /chat`

## Como usar no HTML copia

1. Suba a API:

```bash
python3 api_local.py
```

2. Abra `Relatório - Em contrução 2026 CFO Virtual.html`

3. Na aba `CFO Virtual`, confirme a base da API:

```text
http://127.0.0.1:8000
```

4. Clique em atualizar resumo ou envie uma pergunta no chat.

Importante:

- se `OPENAI_API_KEY` nao estiver preenchida no `.env`, o resumo executivo funciona, mas o `POST /chat` respondera `503` com a instrucao exata de configuracao

## Prototipo de aplicativo isolado

Existe um prototipo separado em `app-teste/`.

Ele nao altera o link atual do relatorio. Funciona como uma casca de aplicativo/PWA e carrega o relatorio publicado em um iframe.
Antes do deploy, o script copia `Relatório - Em contrução 2026 v4.backup.html` para `app-teste/report.html`.

Para publicar esse teste em um projeto Vercel separado:

```bash
cd "/home/nic27/repos/Relatório/Painel em contrução"
bash ./publish_test_app.sh
```
