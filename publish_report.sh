#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

FILES=(
  "Relatório - Em contrução 2026 v4.backup.html"
  "vercel.json"
  "api/report.mjs"
  "despesas.inline.js"
  "recebimentos.inline.js"
  "projecao.inline.js"
)

git add -- "${FILES[@]}"

if git diff --cached --quiet; then
  echo "Nenhuma alteracao de publicacao encontrada."
  exit 0
fi

MESSAGE="${1:-Atualiza relatorio publicado $(date '+%Y-%m-%d %H:%M:%S')}"

git commit -m "$MESSAGE"
git push origin main

echo
echo "Push concluido."
echo "O Vercel deve atualizar o link automaticamente apos o deploy da branch main."
