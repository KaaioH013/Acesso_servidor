import argparse
from datetime import datetime
from pathlib import Path

import pandas as pd

import fase4_dashboard as f4
from src.conexao import get_engine

INPUT_DEFAULT = Path("exports") / "Rotores e Estatores com estoque minimo.xlsx"
OUTPUT_DIR = Path("exports")
OUTPUT_DIR.mkdir(exist_ok=True)


def parse_args():
    p = argparse.ArgumentParser(description="Relatorio de armazem: estoque minimo x saldo x consumo")
    p.add_argument("--arquivo-minimos", type=str, default=str(INPUT_DEFAULT), help="Planilha de estoque minimo")
    p.add_argument("--dias-consumo", type=int, default=90, help="Janela de consumo em dias")
    p.add_argument("--tp-estoque", type=str, default="AL", help="Tipo de estoque MT_ESTOQUE (ex: AL)")
    return p.parse_args()


def carregar_minimos(caminho: Path) -> pd.DataFrame:
    if not caminho.exists():
        raise FileNotFoundError(f"Arquivo de minimos nao encontrado: {caminho}")

    df = pd.read_excel(caminho)
    cols = {c.upper().strip(): c for c in df.columns}
    required = ["MATERIAL", "DESCRICAO", "UNDEST", "VMINIMO"]
    faltando = [c for c in required if c not in cols]
    if faltando:
        raise ValueError(f"Colunas obrigatorias ausentes na planilha: {faltando}")

    out = df[[cols["MATERIAL"], cols["DESCRICAO"], cols["UNDEST"], cols["VMINIMO"]].copy()]
    out.columns = ["MATERIAL", "DESCRICAO", "UNDEST", "VMINIMO"]

    out["MATERIAL"] = out["MATERIAL"].astype(str).str.strip()
    out["DESCRICAO"] = out["DESCRICAO"].astype(str).str.strip()
    out["UNDEST"] = out["UNDEST"].astype(str).str.strip()
    out["VMINIMO"] = pd.to_numeric(out["VMINIMO"], errors="coerce").fillna(0)

    out = out.drop_duplicates(subset=["MATERIAL"], keep="first").reset_index(drop=True)
    return out


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


def query_consumo(engine, dias: int) -> pd.DataFrame:
    dias = max(1, int(dias))
    sql = f"""
    SELECT
        CAST(i.MATERIAL AS varchar(30)) AS MATERIAL,
        SUM(ISNULL(i.QTDE, 0)) AS CONSUMO_QTDE,
        SUM(ISNULL(i.VLRTOTAL, 0)) AS CONSUMO_VALOR,
        COUNT(DISTINCT p.CODIGO) AS PEDIDOS
    FROM VE_PEDIDOITENS i
    JOIN VE_PEDIDO p ON p.CODIGO = i.PEDIDO
    WHERE p.DTPEDIDO >= DATEADD(DAY, -{dias}, CAST(GETDATE() AS date))
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
    GROUP BY CAST(i.MATERIAL AS varchar(30))
    """
    return pd.read_sql(sql, engine)


def query_ultima_compra(engine) -> pd.DataFrame:
    sql = """
    SELECT
        CAST(MATERIAL AS varchar(30)) AS MATERIAL,
        MAX(DATA) AS DT_ULTIMA_COMPRA
    FROM MT_MOVIMENTACAO
    WHERE EVENTO = 3
    GROUP BY CAST(MATERIAL AS varchar(30))
    """
    return pd.read_sql(sql, engine)


def classificar_status(estoque: float, minimo: float) -> str:
    if estoque <= 0:
        return "RUPTURA"
    if estoque < minimo:
        return "ABAIXO_MINIMO"
    return "OK"


def classificar_prioridade(status: str, cobertura_dias) -> str:
    if status == "RUPTURA":
        return "CRITICA"
    if status == "ABAIXO_MINIMO":
        if pd.notna(cobertura_dias) and cobertura_dias <= 15:
            return "ALTA"
        return "MEDIA"
    return "OK"


def montar_relatorio(df_min: pd.DataFrame, df_est: pd.DataFrame, df_cons: pd.DataFrame, df_compra: pd.DataFrame, dias: int):
    base = df_min.merge(df_est, on="MATERIAL", how="left")
    base = base.merge(df_cons, on="MATERIAL", how="left")
    base = base.merge(df_compra, on="MATERIAL", how="left")

    base["ESTOQUE_ATUAL"] = pd.to_numeric(base["ESTOQUE_ATUAL"], errors="coerce").fillna(0)
    base["CONSUMO_QTDE"] = pd.to_numeric(base["CONSUMO_QTDE"], errors="coerce").fillna(0)
    base["CONSUMO_VALOR"] = pd.to_numeric(base["CONSUMO_VALOR"], errors="coerce").fillna(0)
    base["PEDIDOS"] = pd.to_numeric(base["PEDIDOS"], errors="coerce").fillna(0).astype(int)

    base["CONSUMO_DIA_MEDIO"] = (base["CONSUMO_QTDE"] / max(1, dias)).round(4)
    base["COBERTURA_DIAS"] = base.apply(
        lambda r: (r["ESTOQUE_ATUAL"] / r["CONSUMO_DIA_MEDIO"]) if r["CONSUMO_DIA_MEDIO"] > 0 else pd.NA,
        axis=1,
    )
    base["COBERTURA_DIAS"] = pd.to_numeric(base["COBERTURA_DIAS"], errors="coerce").round(1)

    base["GAP_REPOSICAO"] = (base["VMINIMO"] - base["ESTOQUE_ATUAL"]).clip(lower=0)

    base["STATUS_ESTOQUE"] = base.apply(lambda r: classificar_status(r["ESTOQUE_ATUAL"], r["VMINIMO"]), axis=1)
    base["PRIORIDADE"] = base.apply(lambda r: classificar_prioridade(r["STATUS_ESTOQUE"], r["COBERTURA_DIAS"]), axis=1)

    base["SUGESTAO_REPOSICAO"] = base["GAP_REPOSICAO"].round(2)
    base["DT_ULTIMA_COMPRA"] = pd.to_datetime(base["DT_ULTIMA_COMPRA"], errors="coerce")

    ordem_prioridade = {"CRITICA": 1, "ALTA": 2, "MEDIA": 3, "OK": 4}
    base["_ord"] = base["PRIORIDADE"].map(ordem_prioridade).fillna(9)
    base = base.sort_values(["_ord", "GAP_REPOSICAO", "CONSUMO_QTDE"], ascending=[True, False, False]).drop(columns=["_ord"])

    criticos = base[base["PRIORIDADE"].isin(["CRITICA", "ALTA", "MEDIA"])].copy()
    abaixo_min = base[base["STATUS_ESTOQUE"].isin(["RUPTURA", "ABAIXO_MINIMO"])].copy()

    resumo = pd.DataFrame(
        [
            {"Metrica": "Itens_base", "Valor": len(base)},
            {"Metrica": "Itens_ruptura", "Valor": int((base["STATUS_ESTOQUE"] == "RUPTURA").sum())},
            {"Metrica": "Itens_abaixo_minimo", "Valor": int((base["STATUS_ESTOQUE"] == "ABAIXO_MINIMO").sum())},
            {"Metrica": "Itens_ok", "Valor": int((base["STATUS_ESTOQUE"] == "OK").sum())},
            {"Metrica": "Gap_total_reposicao", "Valor": float(base["GAP_REPOSICAO"].sum())},
            {"Metrica": f"Consumo_total_{dias}d", "Valor": float(base["CONSUMO_QTDE"].sum())},
        ]
    )

    return resumo, criticos, abaixo_min, base


def gerar_excel(resumo: pd.DataFrame, criticos: pd.DataFrame, abaixo_min: pd.DataFrame, base: pd.DataFrame, dias: int) -> Path:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = OUTPUT_DIR / f"relatorio_armazem_minimos_{dias}d_{ts}.xlsx"

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        resumo.to_excel(writer, sheet_name="Resumo", index=False)
        criticos.to_excel(writer, sheet_name="Criticos", index=False)
        abaixo_min.to_excel(writer, sheet_name="Abaixo_Minimo", index=False)
        base.to_excel(writer, sheet_name="Base_Completa", index=False)

    return out


def main():
    args = parse_args()
    arquivo_minimos = Path(args.arquivo_minimos)

    engine = get_engine()
    df_min = carregar_minimos(arquivo_minimos)
    df_est = query_estoque(engine, args.tp_estoque)
    df_cons = query_consumo(engine, args.dias_consumo)
    df_compra = query_ultima_compra(engine)

    resumo, criticos, abaixo_min, base = montar_relatorio(df_min, df_est, df_cons, df_compra, args.dias_consumo)
    out = gerar_excel(resumo, criticos, abaixo_min, base, args.dias_consumo)

    print("Relatorio de armazem gerado")
    print(f"Arquivo: {out}")
    print(f"Itens base: {len(base)}")
    print(f"Ruptura: {(base['STATUS_ESTOQUE'] == 'RUPTURA').sum()}")
    print(f"Abaixo minimo: {(base['STATUS_ESTOQUE'] == 'ABAIXO_MINIMO').sum()}")
    print(f"Gap total reposicao: {base['GAP_REPOSICAO'].sum():,.2f}")


if __name__ == "__main__":
    main()
