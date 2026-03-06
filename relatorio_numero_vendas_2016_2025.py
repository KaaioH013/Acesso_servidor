from datetime import datetime
from pathlib import Path

import pandas as pd

from src.conexao import get_engine

OUTPUT_DIR = Path("exports")
OUTPUT_DIR.mkdir(exist_ok=True)


def query_vendas_mensal(engine) -> pd.DataFrame:
    sql = """
    SELECT
        YEAR(p.DTPEDIDO) AS ANO,
        MONTH(p.DTPEDIDO) AS MES,
        COUNT(DISTINCT p.CODIGO) AS QTD_VENDAS
    FROM VE_PEDIDO p
    WHERE p.DTPEDIDO >= '2016-01-01'
      AND p.DTPEDIDO < '2026-01-01'
      AND p.STATUS <> 'C'
    GROUP BY YEAR(p.DTPEDIDO), MONTH(p.DTPEDIDO)
    ORDER BY YEAR(p.DTPEDIDO), MONTH(p.DTPEDIDO)
    """
    return pd.read_sql(sql, engine)


def montar_grade_mensal(df: pd.DataFrame) -> pd.DataFrame:
    calendario = pd.DataFrame(
        {
            "MES_REF": pd.date_range("2016-01-01", "2025-12-01", freq="MS"),
        }
    )
    calendario["ANO"] = calendario["MES_REF"].dt.year
    calendario["MES"] = calendario["MES_REF"].dt.month

    out = calendario.merge(df, on=["ANO", "MES"], how="left")
    out["QTD_VENDAS"] = pd.to_numeric(out["QTD_VENDAS"], errors="coerce").fillna(0).astype(int)
    out["MES_NOME"] = out["MES_REF"].dt.strftime("%b")
    out["MES_ANO"] = out["MES_REF"].dt.strftime("%m/%Y")

    cols = ["ANO", "MES", "MES_NOME", "MES_ANO", "QTD_VENDAS"]
    return out[cols].copy()


def montar_resumo_anual(mensal: pd.DataFrame) -> pd.DataFrame:
    resumo = (
        mensal.groupby("ANO", as_index=False)
        .agg(
            QTD_VENDAS_ANO=("QTD_VENDAS", "sum"),
            MEDIA_MENSAL=("QTD_VENDAS", "mean"),
            MAX_MENSAL=("QTD_VENDAS", "max"),
            MIN_MENSAL=("QTD_VENDAS", "min"),
        )
    )
    resumo["MEDIA_MENSAL"] = resumo["MEDIA_MENSAL"].round(2)
    return resumo


def salvar_excel(mensal: pd.DataFrame, anual: pd.DataFrame) -> Path:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = OUTPUT_DIR / f"relatorio_numero_vendas_2016_2025_{ts}.xlsx"

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        mensal.to_excel(writer, sheet_name="Vendas_Mensal", index=False)
        anual.to_excel(writer, sheet_name="Resumo_Anual", index=False)

    return out


def main():
    engine = get_engine()
    base = query_vendas_mensal(engine)
    mensal = montar_grade_mensal(base)
    anual = montar_resumo_anual(mensal)
    out = salvar_excel(mensal, anual)

    print("Relatorio de numero de vendas gerado")
    print(f"Arquivo: {out}")
    print(f"Meses no periodo: {len(mensal)}")
    print(f"Total vendas 2016-2025: {int(mensal['QTD_VENDAS'].sum())}")


if __name__ == "__main__":
    main()
