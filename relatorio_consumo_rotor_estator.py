import argparse
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

import fase4_dashboard as f4
from src.conexao import get_engine

BASE_MINIMOS = Path("exports") / "Rotores e Estatores com estoque minimo.xlsx"
OUTPUT_DIR = Path("exports")
OUTPUT_DIR.mkdir(exist_ok=True)


def parse_args():
    p = argparse.ArgumentParser(description="Consumo mensal/anual de rotores e estatores com sugestao de estoque minimo")
    p.add_argument("--arquivo-base", type=str, default=str(BASE_MINIMOS), help="Planilha base de materiais")
    p.add_argument("--meses-historico", type=int, default=24, help="Meses para historico de consumo")
    p.add_argument("--tp-estoque", type=str, default="AL", help="Tipo de estoque no MT_ESTOQUE")
    p.add_argument(
        "--cobertura-meses-sugerida",
        type=float,
        default=2.0,
        help="Cobertura em meses para calculo do minimo sugerido",
    )
    return p.parse_args()


def carregar_base(caminho: Path) -> pd.DataFrame:
    if not caminho.exists():
        raise FileNotFoundError(f"Arquivo base nao encontrado: {caminho}")

    df = pd.read_excel(caminho)
    cols = {c.upper().strip(): c for c in df.columns}
    required = ["MATERIAL", "DESCRICAO", "UNDEST", "VMINIMO"]
    faltando = [c for c in required if c not in cols]
    if faltando:
        raise ValueError(f"Colunas obrigatorias ausentes: {faltando}")

    out = df[[cols["MATERIAL"], cols["DESCRICAO"], cols["UNDEST"], cols["VMINIMO"]]].copy()
    out.columns = ["MATERIAL", "DESCRICAO", "UNDEST", "VMINIMO_BASE"]
    out["MATERIAL"] = out["MATERIAL"].astype(str).str.strip()
    out["DESCRICAO"] = out["DESCRICAO"].astype(str).str.strip()
    out["UNDEST"] = out["UNDEST"].astype(str).str.strip()
    out["VMINIMO_BASE"] = pd.to_numeric(out["VMINIMO_BASE"], errors="coerce").fillna(0)

    out = out.drop_duplicates(subset=["MATERIAL"], keep="first").reset_index(drop=True)
    out["TIPO_ITEM"] = out["DESCRICAO"].str.upper().map(
        lambda d: "ROTOR" if "ROTOR" in d else ("ESTATOR" if "ESTATOR" in d else "OUTRO")
    )
    return out


def query_consumo(engine, meses_historico: int) -> pd.DataFrame:
    meses_historico = max(1, int(meses_historico))
    sql = f"""
    SELECT
        CAST(i.MATERIAL AS varchar(30)) AS MATERIAL,
        CAST(DATEFROMPARTS(YEAR(p.DTPEDIDO), MONTH(p.DTPEDIDO), 1) AS date) AS MES_REF,
        YEAR(p.DTPEDIDO) AS ANO,
        MONTH(p.DTPEDIDO) AS MES,
        SUM(ISNULL(i.QTDE, 0)) AS CONSUMO_QTDE,
        SUM(ISNULL(i.VLRTOTAL, 0)) AS CONSUMO_VALOR,
        COUNT(DISTINCT p.CODIGO) AS PEDIDOS
    FROM VE_PEDIDOITENS i
    JOIN VE_PEDIDO p ON p.CODIGO = i.PEDIDO
    WHERE p.DTPEDIDO >= DATEADD(MONTH, -{meses_historico}, CAST(GETDATE() AS date))
      AND p.STATUS <> 'C'
      AND i.STATUS <> 'C'
      AND i.FLAGSUB <> 'S'
      AND i.MATERIAL NOT LIKE '8%'
      AND i.TPVENDA NOT IN ({f4.TPVENDA_STR})
      AND p.CODIGO NOT IN (
          SELECT p2.CODIGO
          FROM VE_PEDIDO p2
          JOIN FN_FORNECEDORES f2 ON f2.CODIGO = p2.CLIENTE
          WHERE f2.UF = 'EX'
      )
    GROUP BY
        CAST(i.MATERIAL AS varchar(30)),
        DATEFROMPARTS(YEAR(p.DTPEDIDO), MONTH(p.DTPEDIDO), 1),
        YEAR(p.DTPEDIDO),
        MONTH(p.DTPEDIDO)
    """
    return pd.read_sql(sql, engine)


def query_estoque(engine, tp_estoque: str) -> pd.DataFrame:
    tp = str(tp_estoque).strip().upper().replace("'", "''")
    sql = f"""
    SELECT
        CAST(MATERIAL AS varchar(30)) AS MATERIAL,
        SUM(ISNULL(QTDEREAL, 0)) AS ESTOQUE_ATUAL
    FROM MT_ESTOQUE
    WHERE TPESTOQUE = '{tp}'
    GROUP BY CAST(MATERIAL AS varchar(30))
    """
    return pd.read_sql(sql, engine)


def montar_relatorio(df_base: pd.DataFrame, df_consumo: pd.DataFrame, df_estoque: pd.DataFrame, cobertura_meses: float):
    consumo = df_consumo.copy()
    if not consumo.empty:
        consumo["MES_REF"] = pd.to_datetime(consumo["MES_REF"], errors="coerce")
        consumo["CONSUMO_QTDE"] = pd.to_numeric(consumo["CONSUMO_QTDE"], errors="coerce").fillna(0)
        consumo["CONSUMO_VALOR"] = pd.to_numeric(consumo["CONSUMO_VALOR"], errors="coerce").fillna(0)

    mensal = (
        consumo.groupby(["MATERIAL", "MES_REF", "ANO", "MES"], dropna=False)
        .agg(CONSUMO_QTDE=("CONSUMO_QTDE", "sum"), CONSUMO_VALOR=("CONSUMO_VALOR", "sum"), PEDIDOS=("PEDIDOS", "sum"))
        .reset_index()
        .sort_values(["MATERIAL", "MES_REF"])
    )

    anual = (
        consumo.groupby(["MATERIAL", "ANO"], dropna=False)
        .agg(CONSUMO_QTDE_ANO=("CONSUMO_QTDE", "sum"), CONSUMO_VALOR_ANO=("CONSUMO_VALOR", "sum"), PEDIDOS_ANO=("PEDIDOS", "sum"))
        .reset_index()
        .sort_values(["MATERIAL", "ANO"])
    )

    agg = (
        mensal.groupby("MATERIAL", dropna=False)
        .agg(
            CONSUMO_TOTAL_QTDE=("CONSUMO_QTDE", "sum"),
            CONSUMO_TOTAL_VALOR=("CONSUMO_VALOR", "sum"),
            MESES_COM_CONSUMO=("MES_REF", "nunique"),
            MEDIA_MENSAL_QTDE=("CONSUMO_QTDE", "mean"),
            PICO_MENSAL_QTDE=("CONSUMO_QTDE", "max"),
            ULTIMO_CONSUMO_MES=("MES_REF", "max"),
        )
        .reset_index()
    )

    base = df_base.merge(agg, on="MATERIAL", how="left").merge(df_estoque, on="MATERIAL", how="left")

    for c in [
        "CONSUMO_TOTAL_QTDE",
        "CONSUMO_TOTAL_VALOR",
        "MESES_COM_CONSUMO",
        "MEDIA_MENSAL_QTDE",
        "PICO_MENSAL_QTDE",
        "ESTOQUE_ATUAL",
    ]:
        base[c] = pd.to_numeric(base[c], errors="coerce").fillna(0)

    base["ESTOQUE_ATUAL"] = base["ESTOQUE_ATUAL"].round(2)
    base["MEDIA_MENSAL_QTDE"] = base["MEDIA_MENSAL_QTDE"].round(2)

    base["MINIMO_SUGERIDO_CONSUMO"] = np.ceil(base["MEDIA_MENSAL_QTDE"] * float(cobertura_meses))
    base["MINIMO_SUGERIDO_FINAL"] = np.maximum(base["VMINIMO_BASE"], base["MINIMO_SUGERIDO_CONSUMO"])
    base["MINIMO_SUGERIDO_FINAL"] = base["MINIMO_SUGERIDO_FINAL"].round(0)

    base["GAP_REPOSICAO"] = (base["MINIMO_SUGERIDO_FINAL"] - base["ESTOQUE_ATUAL"]).clip(lower=0).round(2)
    base["COBERTURA_MESES"] = base.apply(
        lambda r: (r["ESTOQUE_ATUAL"] / r["MEDIA_MENSAL_QTDE"]) if r["MEDIA_MENSAL_QTDE"] > 0 else pd.NA,
        axis=1,
    )
    base["COBERTURA_MESES"] = pd.to_numeric(base["COBERTURA_MESES"], errors="coerce").round(2)

    base["STATUS"] = base.apply(
        lambda r: "RUPTURA"
        if r["ESTOQUE_ATUAL"] <= 0
        else ("ABAIXO_SUGERIDO" if r["ESTOQUE_ATUAL"] < r["MINIMO_SUGERIDO_FINAL"] else "OK"),
        axis=1,
    )

    base["PRIORIDADE"] = base.apply(
        lambda r: "CRITICA"
        if r["STATUS"] == "RUPTURA"
        else ("ALTA" if (r["STATUS"] == "ABAIXO_SUGERIDO" and pd.notna(r["COBERTURA_MESES"]) and r["COBERTURA_MESES"] <= 1) else ("MEDIA" if r["STATUS"] == "ABAIXO_SUGERIDO" else "OK")),
        axis=1,
    )

    ordem = {"CRITICA": 1, "ALTA": 2, "MEDIA": 3, "OK": 4}
    base["_ord"] = base["PRIORIDADE"].map(ordem).fillna(9)
    base = base.sort_values(["_ord", "GAP_REPOSICAO", "CONSUMO_TOTAL_QTDE"], ascending=[True, False, False]).drop(columns=["_ord"])

    criticos = base[base["PRIORIDADE"].isin(["CRITICA", "ALTA", "MEDIA"])].copy()

    resumo_tipo = (
        base.groupby(["TIPO_ITEM", "STATUS"], dropna=False)
        .agg(
            Itens=("MATERIAL", "nunique"),
            Gap_Total=("GAP_REPOSICAO", "sum"),
            Consumo_Total=("CONSUMO_TOTAL_QTDE", "sum"),
        )
        .reset_index()
        .sort_values(["TIPO_ITEM", "Itens"], ascending=[True, False])
    )

    resumo_geral = pd.DataFrame(
        [
            {"Metrica": "Itens_base", "Valor": len(base)},
            {"Metrica": "Itens_rotor", "Valor": int((base["TIPO_ITEM"] == "ROTOR").sum())},
            {"Metrica": "Itens_estator", "Valor": int((base["TIPO_ITEM"] == "ESTATOR").sum())},
            {"Metrica": "Itens_criticos", "Valor": len(criticos)},
            {"Metrica": "Gap_total_reposicao", "Valor": float(base["GAP_REPOSICAO"].sum())},
            {"Metrica": "Consumo_total_qtde", "Valor": float(base["CONSUMO_TOTAL_QTDE"].sum())},
        ]
    )

    return mensal, anual, base, criticos, resumo_tipo, resumo_geral


def gerar_excel(mensal, anual, base, criticos, resumo_tipo, resumo_geral, meses_historico):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = OUTPUT_DIR / f"relatorio_consumo_rotor_estator_{meses_historico}m_{ts}.xlsx"

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        resumo_geral.to_excel(writer, sheet_name="Resumo_Geral", index=False)
        resumo_tipo.to_excel(writer, sheet_name="Resumo_Tipo_Status", index=False)
        criticos.to_excel(writer, sheet_name="Criticos", index=False)
        base.to_excel(writer, sheet_name="Base_Item", index=False)
        mensal.to_excel(writer, sheet_name="Consumo_Mensal", index=False)
        anual.to_excel(writer, sheet_name="Consumo_Anual", index=False)

    return out


def main():
    args = parse_args()

    df_base = carregar_base(Path(args.arquivo_base))
    engine = get_engine()

    df_consumo = query_consumo(engine, args.meses_historico)
    df_estoque = query_estoque(engine, args.tp_estoque)

    mensal, anual, base, criticos, resumo_tipo, resumo_geral = montar_relatorio(
        df_base,
        df_consumo,
        df_estoque,
        args.cobertura_meses_sugerida,
    )

    out = gerar_excel(mensal, anual, base, criticos, resumo_tipo, resumo_geral, args.meses_historico)

    print("Relatorio consumo rotor/estator gerado")
    print(f"Arquivo: {out}")
    print(f"Itens base: {len(base)}")
    print(f"Itens criticos: {len(criticos)}")
    print(f"Gap total: {base['GAP_REPOSICAO'].sum():,.2f}")


if __name__ == "__main__":
    main()
