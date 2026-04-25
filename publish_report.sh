#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

FILES=(
  "Relatório - Em contrução 2026 v4.backup.html"
  "data.inline.js"
  "vercel.json"
  "api/report.mjs"
  "despesas.inline.js"
  "recebimentos.inline.js"
  "projecao.inline.js"
  "publish_report.sh"
  "README.md"
)

if ! command -v npx >/dev/null 2>&1; then
  echo "Erro: npx nao foi encontrado. Instale o Node.js/NPM antes de publicar."
  exit 1
fi

if [ ! -f ".vercel/project.json" ]; then
  echo "Erro: projeto Vercel nao vinculado localmente."
  echo "Execute uma vez: npx vercel link --project relatorio-2026"
  exit 1
fi

git add -- "${FILES[@]}"

if git diff --cached --quiet; then
  echo "Nenhuma alteracao de publicacao encontrada no Git."
else
  MESSAGE="${1:-Atualiza relatorio publicado $(date '+%Y-%m-%d %H:%M:%S')}"

  git commit -m "$MESSAGE"
  git push origin main

  echo
  echo "Push concluido."
fi

npx vercel --prod --yes

echo
echo "Deploy do Vercel concluido."
