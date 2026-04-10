#!/usr/bin/env python3
from __future__ import annotations

import calendar
import datetime as dt
import json
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent
BASE_XLSX = ROOT / "Base.xlsx"
OUT = ROOT / "data.inline.js"

MONTHS = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
REGION_BY_UF = {
    "AC": "Norte",
    "AL": "Nordeste",
    "AM": "Norte",
    "AP": "Norte",
    "BA": "Nordeste",
    "CE": "Nordeste",
    "DF": "Centro-Oeste",
    "ES": "Sudeste",
    "GO": "Centro-Oeste",
    "MA": "Nordeste",
    "MG": "Sudeste",
    "MS": "Centro-Oeste",
    "MT": "Centro-Oeste",
    "PA": "Norte",
    "PB": "Nordeste",
    "PE": "Nordeste",
    "PI": "Nordeste",
    "PR": "Sul",
    "RJ": "Sudeste",
    "RN": "Nordeste",
    "RO": "Norte",
    "RR": "Norte",
    "RS": "Sul",
    "SC": "Sul",
    "SE": "Nordeste",
    "SP": "Sudeste",
    "TO": "Norte",
}


def clean_text(value: object, default: str = "") -> str:
    return str(value or "").strip() or default


def as_number(value: object) -> float:
    if value in (None, "", "-"):
        return 0.0
    return float(value)


def as_date(value: object) -> dt.date | None:
    if value in (None, "", "-"):
        return None
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    text = str(value).strip()
    if not text:
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def week_of_month(value: dt.date) -> int:
    # Calendar weeks inside the month, using Sunday-Saturday boundaries.
    first_day = value.replace(day=1)
    start_offset = (first_day.weekday() + 1) % 7
    return ((value.day + start_offset - 1) // 7) + 1


def week_bounds_in_month(value: dt.date) -> tuple[dt.date, dt.date]:
    month_start = value.replace(day=1)
    month_end = value.replace(day=calendar.monthrange(value.year, value.month)[1])
    days_since_sunday = (value.weekday() + 1) % 7
    week_start = value - dt.timedelta(days=days_since_sunday)
    week_end = week_start + dt.timedelta(days=6)
    if week_start < month_start:
        week_start = month_start
    if week_end > month_end:
        week_end = month_end
    return week_start, week_end


def month_meta(value: dt.date) -> dict[str, object]:
    return {
        "Ano": value.year,
        "MesNumero": value.month,
        "MesNome": MONTHS[value.month - 1],
        "AnoMes": f"{value.year:04d}-{value.month:02d}",
        "SemanaMes": week_of_month(value),
    }


def focus_params(value: dt.date) -> dict[str, str]:
    week_index = week_of_month(value)
    start_date, end_date = week_bounds_in_month(value)
    return {
        "AnoFoco": str(value.year),
        "MesFoco": str(value.month),
        "SemanaIndice": str(week_index),
        "SemanaInicio": start_date.isoformat(),
        "SemanaFim": end_date.isoformat(),
        "SemanaLabel": f"{start_date.strftime('%d/%m')}–{end_date.strftime('%d/%m')}",
    }


def build_vendas(wb: openpyxl.Workbook) -> list[dict[str, object]]:
    ws = wb["Base"]
    headers = list(next(ws.iter_rows(min_row=1, max_row=1, values_only=True)))
    idx = {str(h).strip(): i for i, h in enumerate(headers) if h is not None}

    required = [
        "Nr. nota",
        "Pessoa",
        "Dt. emissão",
        "Tipo",
        "Linha de Receita",
        "Vl. total vendido",
        "Cidade",
        "UF",
        "Estado",
        "Detalhamento_Produto",
        "Classficação Produto",
        "Vendedor",
        "Forma de Pagamento",
        "Origem do LEAD",
    ]
    missing = [name for name in required if name not in idx]
    if missing:
        raise SystemExit(f"Colunas ausentes na aba Base: {', '.join(missing)}")

    rows: list[dict[str, object]] = []
    for sale_id, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=1):
        date_value = as_date(row[idx["Dt. emissão"]])
        if date_value is None:
            continue
        meta = month_meta(date_value)
        uf = clean_text(row[idx["UF"]]).upper()
        rows.append({
            "VendaID": str(sale_id),
            "DataEmissao": date_value.isoformat(),
            **meta,
            "NF": clean_text(row[idx["Nr. nota"]]),
            "Tipo": clean_text(row[idx["Tipo"]], "Nao informado"),
            "LinhaReceita": clean_text(row[idx["Linha de Receita"]], "Nao informado"),
            "Valor": round(as_number(row[idx["Vl. total vendido"]]), 2),
            "Cliente": clean_text(row[idx["Pessoa"]], "Nao informado"),
            "Cidade": clean_text(row[idx["Cidade"]], "Nao informado"),
            "UF": uf,
            "Estado": clean_text(row[idx["Estado"]], "Nao informado"),
            "DetalhamentoProduto": clean_text(row[idx["Detalhamento_Produto"]], "Nao informado"),
            "ClassificacaoProduto": clean_text(row[idx["Classficação Produto"]], "Nao informado"),
            "Vendedor": clean_text(row[idx["Vendedor"]]),
            "FormaPagamento": clean_text(row[idx["Forma de Pagamento"]], "Não informado"),
            "OrigemLead": clean_text(row[idx["Origem do LEAD"]], "Não informado"),
            "Regiao": REGION_BY_UF.get(uf, "Nao informado"),
        })
    return rows


def build_financeiro(wb: openpyxl.Workbook) -> list[dict[str, object]]:
    ws = wb["Financeiro"]
    headers = list(next(ws.iter_rows(min_row=1, max_row=1, values_only=True)))
    idx = {str(h).strip(): i for i, h in enumerate(headers) if h is not None}

    required = [
        "Nome Cliente",
        "Tipo de Pagamento",
        "Valor da Entrada",
        "Data da Entrada",
        "Quantidade de Parcelas",
        "Valor da Parcela",
        "Data da 1ª Parcela",
    ]
    missing = [name for name in required if name not in idx]
    if missing:
        raise SystemExit(f"Colunas ausentes na aba Financeiro: {', '.join(missing)}")

    rows: list[dict[str, object]] = []
    event_id = 1
    for row in ws.iter_rows(min_row=2, values_only=True):
        client = clean_text(row[idx["Nome Cliente"]], "Nao informado")
        payment_type = clean_text(row[idx["Tipo de Pagamento"]], "Nao informado")
        entry_value = round(as_number(row[idx["Valor da Entrada"]]), 2)
        entry_date = as_date(row[idx["Data da Entrada"]])
        parcel_count = int(as_number(row[idx["Quantidade de Parcelas"]]))
        parcel_value = round(as_number(row[idx["Valor da Parcela"]]), 10)
        first_parcel_date = as_date(row[idx["Data da 1ª Parcela"]])

        if entry_date and entry_value > 0:
          meta = month_meta(entry_date)
          rows.append({
              "EventoID": str(event_id),
              "Cliente": client,
              "EventoTipo": "Entrada",
              "DataEvento": entry_date.isoformat(),
              **meta,
              "Valor": entry_value,
              "TipoPagamentoOrigem": payment_type,
              "QuantidadeParcelasOrigem": str(parcel_count),
              "ValorParcelaOrigem": parcel_value,
              "DataEntradaOrigem": entry_date.isoformat(),
              "DataPrimeiraParcelaOrigem": first_parcel_date.isoformat() if first_parcel_date else "",
          })
          event_id += 1

        if first_parcel_date and parcel_count > 0 and parcel_value > 0:
            for parcel_idx in range(parcel_count):
                target_year = first_parcel_date.year + ((first_parcel_date.month - 1 + parcel_idx) // 12)
                target_month = ((first_parcel_date.month - 1 + parcel_idx) % 12) + 1
                target_day = min(
                    first_parcel_date.day,
                    calendar.monthrange(target_year, target_month)[1],
                )
                parcel_date = dt.date(target_year, target_month, target_day)
                meta = month_meta(parcel_date)
                rows.append({
                    "EventoID": str(event_id),
                    "Cliente": client,
                    "EventoTipo": "Parcela",
                    "DataEvento": parcel_date.isoformat(),
                    **meta,
                    "Valor": parcel_value,
                    "TipoPagamentoOrigem": payment_type,
                    "QuantidadeParcelasOrigem": str(parcel_count),
                    "ValorParcelaOrigem": parcel_value,
                    "DataEntradaOrigem": entry_date.isoformat() if entry_date else "",
                    "DataPrimeiraParcelaOrigem": first_parcel_date.isoformat(),
                })
                event_id += 1

    return rows


def build_estoque(wb: openpyxl.Workbook) -> list[dict[str, object]]:
    ws = wb["Estoque"]
    headers = list(next(ws.iter_rows(min_row=1, max_row=1, values_only=True)))
    idx = {str(h).strip(): i for i, h in enumerate(headers) if h is not None}

    required = [
        "Empresa",
        "Setor",
        "Grupo de produto",
        "Família de produto",
        "Código",
        "Nome do produto",
        "Quantidade",
        "Qt. futura",
        "Qt. requisitada",
        "Preço de Custo Splan",
        "CUSTO",
    ]
    missing = [name for name in required if name not in idx]
    if missing:
        raise SystemExit(f"Colunas ausentes na aba Estoque: {', '.join(missing)}")

    rows_by_code: dict[str, dict[str, object]] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        code = clean_text(row[idx["Código"]])
        if not code:
            continue

        quantity = as_number(row[idx["Quantidade"]])
        future_qty = as_number(row[idx["Qt. futura"]])
        requested_qty = as_number(row[idx["Qt. requisitada"]])
        unit_cost = as_number(row[idx["Preço de Custo Splan"]])
        total_cost = as_number(row[idx["CUSTO"]])
        sector = clean_text(row[idx["Setor"]]).upper()

        item = rows_by_code.setdefault(code, {
            "Empresa": clean_text(row[idx["Empresa"]], "Nao informado"),
            "Codigo": code,
            "Produto": clean_text(row[idx["Nome do produto"]], "Nao informado"),
            "Familia": clean_text(row[idx["Família de produto"]], "Nao informado"),
            "Grupo": clean_text(row[idx["Grupo de produto"]], "Nao informado"),
            "Quantidade": 0.0,
            "EstoquePrincipal": 0.0,
            "EstoqueDemonstracao": 0.0,
            "QtFutura": 0.0,
            "QtRequisitada": 0.0,
            "CustoUnitario": 0.0,
            "CustoTotal": 0.0,
        })

        item["Quantidade"] = round(float(item["Quantidade"]) + quantity, 2)
        item["QtFutura"] = round(float(item["QtFutura"]) + future_qty, 2)
        item["QtRequisitada"] = round(float(item["QtRequisitada"]) + requested_qty, 2)
        item["CustoTotal"] = round(float(item["CustoTotal"]) + total_cost, 2)
        if unit_cost > 0:
            item["CustoUnitario"] = round(unit_cost, 2)

        if "PRINCIPAL" in sector:
            item["EstoquePrincipal"] = round(float(item["EstoquePrincipal"]) + quantity, 2)
        elif "DEMONSTRA" in sector:
            item["EstoqueDemonstracao"] = round(float(item["EstoqueDemonstracao"]) + quantity, 2)

    return [rows_by_code[code] for code in sorted(rows_by_code)]


def main() -> None:
    if not BASE_XLSX.exists():
        raise SystemExit(f"Arquivo nao encontrado: {BASE_XLSX}")

    wb = openpyxl.load_workbook(BASE_XLSX, read_only=True, data_only=True)
    vendas = build_vendas(wb)
    financeiro = build_financeiro(wb)
    estoque = build_estoque(wb)

    if vendas:
        focus_date = max(dt.date.fromisoformat(row["DataEmissao"]) for row in vendas)
    elif financeiro:
        focus_date = max(dt.date.fromisoformat(row["DataEvento"]) for row in financeiro)
    else:
        focus_date = dt.date.today()

    payload = {
        "vendas": vendas,
        "financeiro": financeiro,
        "estoque": estoque,
        "parametros": [focus_params(focus_date)],
    }
    OUT.write_text(
        "window.__DASHBOARD_DATA__ = " + json.dumps(payload, ensure_ascii=False) + ";\n",
        encoding="utf-8",
    )
    print(f"Gerado: {OUT}")


if __name__ == "__main__":
    main()
