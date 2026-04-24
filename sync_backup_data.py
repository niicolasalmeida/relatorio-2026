#!/usr/bin/env python3
from __future__ import annotations

import re
from pathlib import Path


ROOT = Path(__file__).resolve().parent
BACKUP_HTML = ROOT / "Relatório - Em contrução 2026 v4.backup.html"


def wrapped_script(path: Path) -> str:
    return f"<script>\n{path.read_text(encoding='utf-8').rstrip()}\n</script>"


def replace_block(content: str, pattern: str, replacement: str, label: str) -> str:
    updated, count = re.subn(pattern, replacement, content, count=1)
    if count != 1:
        raise SystemExit(f"Bloco '{label}' nao encontrado para substituicao.")
    return updated


def main() -> None:
    content = BACKUP_HTML.read_text(encoding="utf-8")

    combined_block = "\n\n".join([
        (ROOT / "brasil.geojson.js").read_text(encoding="utf-8").rstrip(),
        (ROOT / "data.inline.js").read_text(encoding="utf-8").rstrip(),
        (ROOT / "despesas.inline.js").read_text(encoding="utf-8").rstrip(),
        (ROOT / "recebimentos.inline.js").read_text(encoding="utf-8").rstrip(),
        (ROOT / "projecao.inline.js").read_text(encoding="utf-8").rstrip(),
    ])

    content = replace_block(
        content,
        r"(?s)<script>\s*window\.__BRASIL_GEOJSON__ = .*?</script>\s*<script>\s*const MONTHS = \[\"Jan\"",
        f"<script>\n{combined_block}\n</script>\n  <script>\n    const MONTHS = [\"Jan\"",
        "data_embutida",
    )

    BACKUP_HTML.write_text(content, encoding="utf-8")
    print(f"Atualizado: {BACKUP_HTML}")


if __name__ == "__main__":
    main()
