#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT_DIR"

FILES=(
  ".vercelignore"
  "Relatório - Em contrução 2026 v4.backup.html"
  "Relatório - Em contrução 2026 CFO Virtual.html"
  "data.inline.js"
  "vercel.json"
  "api/report.mjs"
  "despesas.inline.js"
  "despesas-observacoes.inline.js"
  "recebimentos.inline.js"
  "projecao.inline.js"
  "publish_report.sh"
  "README.md"
)

LOCAL_NODE_BIN="${ROOT_DIR}/.tools/node/bin"
if [ -x "${LOCAL_NODE_BIN}/node" ] && [ -x "${LOCAL_NODE_BIN}/npx" ]; then
  export PATH="${LOCAL_NODE_BIN}:$PATH"
fi

if ! command -v npx >/dev/null 2>&1; then
  echo "Erro: npx nao foi encontrado. Instale Node.js/NPM ou extraia um Node local em .tools/node."
  exit 1
fi

if ! git config user.name >/dev/null || ! git config user.email >/dev/null; then
  echo "Erro: configure uma vez o nome e o email do Git com git config --global."
  exit 1
fi

if [ ! -f ".vercel/project.json" ]; then
  echo "Erro: projeto Vercel nao vinculado localmente."
  echo "Execute uma vez: npx vercel link --project relatorio-2026"
  exit 1
fi

git add -- .gitignore "${FILES[@]}"

if git diff --cached --quiet; then
  echo "Nenhuma alteracao de publicacao encontrada no Git."
else
  MESSAGE="${1:-Atualiza relatorio publicado $(date '+%Y-%m-%d %H:%M:%S')}"

  git commit -m "$MESSAGE"
  if ! git push origin main; then
    echo
    echo "Aviso: o GitHub recusou o push. O deploy do Vercel continuara normalmente."
    echo "Configure a autenticacao persistente do GitHub para sincronizar o commit depois."
  else
    echo
    echo "Push concluido."
  fi
fi

if ! npx vercel --prod --yes; then
  echo
  echo "Erro: o Vercel recusou o deploy. Execute uma vez: npx vercel login"
  echo "Depois execute este script novamente."
  exit 1
fi

echo
echo "Deploy do Vercel concluido."
