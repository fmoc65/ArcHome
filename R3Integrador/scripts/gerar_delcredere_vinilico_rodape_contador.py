#!/usr/bin/env python3
"""Gera seis rascunhos Del Credere com vinílicos e rodapés para validação fiscal."""

from __future__ import annotations

import hashlib
import json
import re
import unicodedata
import re
from copy import copy
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import PatternFill


ROOT = Path(__file__).resolve().parents[1]
SOURCE = Path("/home/fernando/Projetos/Work/ARCHOME/Planilhas/delcredere/100 TABELA DEL CREDERE SP ATUALIZADA.xlsx")
TEMPLATE = Path("/home/fernando/Projetos/Work/ARCHOME/Planilhas/IMPORTACAO_ERP_VINILICO_20260701_200706_CONTADOR.xlsx")
VINILICO_DIR = ROOT / "Saida" / "DELCREDERE_VINILICO_20260909"
OUT_DIR = ROOT / "Saida" / "DELCREDERE_VINILICO_RODAPE_CONTADOR_20261005"
TABLES = {"DEL5": 18, "DEL10": 19, "DEL15": 20, "DEL20": 21, "DEL25": 22, "DEL30": 23}
RED = PatternFill(fill_type="solid", fgColor="FFFF0000")
WHITE_FONT = copy(load_workbook(TEMPLATE, read_only=False).active[1][1].font)
WHITE_FONT = copy(WHITE_FONT)
WHITE_FONT.color = "FFFFFFFF"

# Colunas mandatórias conforme validação do layout e as tratadas fiscalmente
# como obrigatórias para o rodapé. AD (30) é explicitamente CSOSN.
# Legacy required-field indexes from the existing generator are zero-based;
# convert them to worksheet columns and explicitly add AD/CSOSN.
MANDATORY = {2, 4, 6, 7, 8, 13, 14, 15, 16, 19, 20, 24, 25, 26, 27, 28,
             29, 30, 31, 32, 39, 40, 51, 57, 58, 59, 30}
RODAPE_FISCAL = {13, 18, 19, 20, 21, 24, 26, 27, 28, 29, 30, 31, 32,
                 37, 38, 39, 57, 58, 59}


def clean(value: object) -> str:
    return "" if value is None else str(value).replace("***", "").strip()


def erp_text(value: object, *, model: bool = False) -> str | None:
    """Return plain ASCII product text; encode decimal dimensions with P (e.g. 19P2)."""
    if value is None:
        return None
    value = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode("ascii")
    if model:
        value = re.sub(r"(?<=\d)[,.](?=\d)", "P", value)
        value = re.sub(r"[^A-Za-z0-9]", "", value)
        return value.upper() or None
    value = re.sub(r"(?<=\d)[,.](?=\d)", "P", value)
    value = re.sub(r"[^A-Za-z0-9]+", " ", value)
    value = " ".join(value.split())
    return value or None


def money(value: object) -> Decimal:
    return Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def sale_price(value: object) -> Decimal:
    return (money(value) * Decimal("1.0536")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def linear_meters(value: object) -> float | None:
    """Extract the numeric meters-per-box value from e.g. '67,20 metro linear'."""
    if value is None:
        return None
    match = re.search(r"\d+(?:[.,]\d+)?", str(value))
    return float(Decimal(match.group(0).replace(",", "."))) if match else None


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def load_rodapes() -> dict[str, dict[str, object]]:
    ws = load_workbook(SOURCE, read_only=True, data_only=True).active
    products: dict[str, dict[str, object]] = {}
    fmt = color = ""
    for row in ws.iter_rows(min_row=293, max_row=305, values_only=True):
        ref = clean(row[1])
        if not ref.startswith("RP"):
            continue
        if row[0]:
            fmt = str(row[0]).replace(" ", "").upper()
        if row[4]:
            color = str(row[4]).strip().lower()
        # Colunas 18–23 da fonte: valores publicados para DEL 5–30.
        prices = {name: row[col - 1] for name, col in TABLES.items()}
        if len(prices) != 6 or any(v is None for v in prices.values()):
            raise ValueError(f"Preços incompletos para {ref}")
        products[ref] = {
            "formato": fmt, "cor": color, "acabamento": clean(row[5]).lower(),
            "derivacao": clean(row[9]), "m_linear_caixa": linear_meters(row[10]),
            "pecas_caixa": row[11], "peso_bruto_caixa": row[15], "precos": prices,
        }
    if len(products) != 13:
        raise ValueError(f"Esperados 13 rodapés; encontrados {len(products)}")
    return products


def populate_rodape(ws, start_row: int, ref: str, product: dict[str, object], tier: str) -> None:
    row = start_row
    for col in range(1, ws.max_column + 1):
        source_cell = ws.cell(row - 1, col)
        target_cell = ws.cell(row, col)
        if source_cell.has_style:
            target_cell._style = copy(source_cell._style)
        if source_cell.number_format:
            target_cell.number_format = source_cell.number_format
        if source_cell.alignment:
            target_cell.alignment = copy(source_cell.alignment)
        target_cell.value = None

    factory = money(product["precos"][tier])
    description = f"RODAPE DE POLIESTIRENO {product['formato']} {product['cor']} {product['acabamento']}"
    values = {
        2: ref, 4: description, 5: f"RODAPE {product['formato']} {product['cor']} {product['acabamento']}",
        6: "RODAPE", 7: "POLIESTIRENO", 8: f"VILLAGRES {tier.removeprefix('DEL')}",
        9: "RODAPE DE POLIESTIRENO", 10: product["formato"], 12: product["cor"],
        15: float(sale_price(factory)), 16: float(factory), 17: 0,
        25: product["m_linear_caixa"], 34: product["peso_bruto_caixa"],
        35: product["pecas_caixa"], 41: f"Derivação {product['derivacao']}; preço da tabela por metro linear. Dados fiscais/unidade pendentes do contador.",
    }
    for col, value in values.items():
        ws.cell(row, col).value = value
    # Force/keep the ERP contract: column AD (30) is CSOSN and is left blank
    # for the accountant on rodapé rows; no vinyl CSOSN is copied over.
    ws.cell(1, 30).value = "CSOSN"
    ws.cell(row, 30).value = None
    for col in (2, 6, 7, 8, 9, 10, 12, 13, 14, 24, 26, 27, 28, 29, 30, 31, 32, 40, 51, 52, 53, 57, 58, 59, 60):
        ws.cell(row, col).number_format = "@"
    # Highlights every missing mandatory value in the complete import layout.
    for col in sorted(MANDATORY | RODAPE_FISCAL):
        cell = ws.cell(row, col)
        if cell.value in (None, ""):
            cell.fill = copy(RED)
            cell.font = copy(WHITE_FONT)


def main() -> None:
    source_ws = load_workbook(SOURCE, read_only=True, data_only=True).active
    vinilicos: dict[str, dict[str, object]] = {}
    fmt = color = ""
    m2_box = weight_box = None
    for row in source_ws.iter_rows(min_row=276, max_row=291, values_only=True):
        ref = clean(row[1])
        if not ref:
            continue
        if row[0] is not None:
            fmt = str(row[0]).replace(" ", "").upper()
        if row[4] is not None:
            color = str(row[4]).strip().lower()
        if row[10] is not None:
            m2_box = row[10]
        if row[15] is not None:
            weight_box = row[15]
        vinilicos[ref] = {"formato": fmt, "cor": color, "m2_caixa": m2_box,
                          "peso_bruto_caixa": weight_box}
    if len(vinilicos) != 16:
        raise ValueError(f"Esperados 16 vinílicos; encontrados {len(vinilicos)}")
    rodapes = load_rodapes()
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    generated = {}
    for tier, price_col in TABLES.items():
        original = load_workbook(VINILICO_DIR / f"IMPORTACAO_ERP_DELCREDERE_{tier}_VINILICO.xlsx")
        ws = original.active
        if ws.max_column != 60 or ws.cell(1, 30).value != "CSOSN":
            raise ValueError("Layout alterado: coluna AD deve ser CSOSN")
        # Read this band's values directly from the authoritative structured source rows.
        by_ref = {}
        for src_row in range(293, 306):
            ref = clean(source_ws.cell(src_row, 2).value)
            if ref in rodapes:
                rodapes[ref]["precos"] = {tier: source_ws.cell(src_row, price_col).value}
                by_ref[ref] = rodapes[ref]
        if len(by_ref) != 13:
            raise ValueError(f"Faixa {tier}: esperados 13 rodapés, encontrados {len(by_ref)}")
        for offset, (ref, p) in enumerate(by_ref.items(), start=ws.max_row + 1):
            populate_rodape(ws, offset, ref, p, tier)
        # Apply the same missing-field cue to any other mandatory blanks in
        # the vinyl rows so the handoff file is consistent as a whole.
        for row_num in range(2, ws.max_row + 1):
            for col in (4, 5, 6, 7, 8, 9, 11, 12):
                ws.cell(row_num, col).value = erp_text(ws.cell(row_num, col).value)
            ws.cell(row_num, 10).value = erp_text(ws.cell(row_num, 10).value, model=True)
            # This note was only an internal warning and contains punctuation
            # that the ERP rejects; the red fiscal blanks already mark it.
            ws.cell(row_num, 41).value = None
            for col in MANDATORY | RODAPE_FISCAL:
                cell = ws.cell(row_num, col)
                if cell.value in (None, ""):
                    cell.fill = copy(RED)
                    cell.font = copy(WHITE_FONT)
        ws.auto_filter.ref = ws.dimensions
        output = OUT_DIR / f"IMPORTACAO_ERP_DELCREDERE_{tier}_VINILICO_RODAPE_PARA_CONTADOR.xlsx"
        original.save(output)
        check = load_workbook(output, read_only=True, data_only=True).active
        rows = list(check.iter_rows(min_row=2, values_only=True))
        if check.max_column != 60 or len(rows) != 29 or check.cell(1, 30).value != "CSOSN":
            raise ValueError(f"Falha estrutural na saída {tier}")
        if any(not r[1] for r in rows) or sum(str(r[1]).startswith("RP") for r in rows) != 13:
            raise ValueError(f"Referências incorretas em {tier}")
        if any(check.cell(r, 30).value not in (None, "") for r in range(18, 31)):
            raise ValueError(f"CSOSN AD deveria ficar pendente para os 13 rodapés em {tier}")
        generated[output.name] = {"sha256": sha256(output), "linhas": len(rows), "vinilicos": 16, "rodapes": 13}

    report = {
        "status": "RASCUNHO_PARA_PREENCHIMENTO_DO_CONTADOR_NAO_IMPORTAR_AINDA",
        "fonte": str(SOURCE), "faixas": list(TABLES), "produtos_por_faixa": {"vinilicos": 16, "rodapes": 13},
        "total_por_arquivo": 29, "colunas": 60, "coluna_AD": "CSOSN",
        "regra_preco": "preço da faixa Del Credere x 1,0536 (regra aplicada aos arquivos de vinílico já existentes)",
        "higienizacao_texto": "Descricao, descricao comercial e demais campos textuais de produto foram convertidos para ASCII sem acentos nem pontuacao; virgula decimal em dimensoes textuais foi representada por P. Observacao removida. Cabecalhos, codigos e campos fiscais foram preservados.",
        "rodape": "Dados comerciais preenchidos a partir da tabela; campos fiscais obrigatórios em branco destacados em vermelho para o contador; CSOSN em AD explicitamente em branco.",
        "arquivos": generated,
    }
    (OUT_DIR / "RELATORIO_AUDITORIA_DELCREDERE_VINILICO_RODAPE.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(OUT_DIR)


if __name__ == "__main__":
    main()
