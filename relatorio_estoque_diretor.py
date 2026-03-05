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


MESES_NOME = {
    1: "Jan",
    2: "Fev",
    3: "Mar",
    4: "Abr",
    5: "Mai",
    6: "Jun",
    7: "Jul",
    8: "Ago",
    9: "Set",
    10: "Out",
    11: "Nov",
    12: "Dez",
}


def parse_args():
    p = argparse.ArgumentParser(description="Relatorio enxuto de estoque para diretoria")
    p.add_argument("--arquivo-base", type=str, default=str(BASE_MINIMOS), help="Planilha base de materiais")
    p.add_argument("--meses-historico", type=int, default=24, help="Historico de consumo em meses")
    p.add_argument("--tp-estoque", type=str, default="AL", help="Tipo de estoque no MT_ESTOQUE")
    p.add_argument("--cobertura-rotor", type=float, default=2.0, help="Cobertura alvo (meses) para rotor")
    p.add_argument("--cobertura-estator", type=float, default=2.0, help="Cobertura alvo (meses) para estator")
    p.add_argument(
        "--piso-base-fracao",
        type=float,
        default=0.35,
        help="Fracao do VMINIMO_BASE usada como piso mensal na aba de sazonalidade (0 a 1)",
    )
    return p.parse_args()


def carregar_base(caminho: Path) -> pd.DataFrame:
    if not caminho.exists():
        raise FileNotFoundError(f"Arquivo base nao encontrado: {caminho}")

    df = pd.read_excel(caminho)
    cols = {c.upper().strip(): c for c in df.columns}
    required = ["MATERIAL", "DESCRICAO", "VMINIMO"]
    faltando = [c for c in required if c not in cols]
    if faltando:
        raise ValueError(f"Colunas obrigatorias ausentes: {faltando}")

    out = df[[cols["MATERIAL"], cols["DESCRICAO"], cols["VMINIMO"]]].copy()
    out.columns = ["MATERIAL", "DESCRICAO", "VMINIMO_BASE"]
    out["MATERIAL"] = out["MATERIAL"].astype(str).str.strip()
    out["DESCRICAO"] = out["DESCRICAO"].astype(str).str.strip()
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
        YEAR(p.DTPEDIDO) AS ANO,
        MONTH(p.DTPEDIDO) AS MES,
        SUM(ISNULL(i.QTDE, 0)) AS CONSUMO_QTDE
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
    GROUP BY CAST(i.MATERIAL AS varchar(30)), YEAR(p.DTPEDIDO), MONTH(p.DTPEDIDO)
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


def montar_aba_minimo(base: pd.DataFrame, consumo: pd.DataFrame, estoque: pd.DataFrame, cob_rotor: float, cob_estator: float):
    agg = (
        consumo.groupby("MATERIAL", dropna=False)
        .agg(CONSUMO_TOTAL_QTDE=("CONSUMO_QTDE", "sum"), MESES_COM_DADO=("MES", "count"))
        .reset_index()
    )

    df = base.merge(agg, on="MATERIAL", how="left").merge(estoque, on="MATERIAL", how="left")
    df["CONSUMO_TOTAL_QTDE"] = pd.to_numeric(df["CONSUMO_TOTAL_QTDE"], errors="coerce").fillna(0)
    df["MESES_COM_DADO"] = pd.to_numeric(df["MESES_COM_DADO"], errors="coerce").fillna(0)
    df["ESTOQUE_ATUAL"] = pd.to_numeric(df["ESTOQUE_ATUAL"], errors="coerce").fillna(0)

    df["MEDIA_MENSAL"] = df.apply(
        lambda r: (r["CONSUMO_TOTAL_QTDE"] / r["MESES_COM_DADO"]) if r["MESES_COM_DADO"] > 0 else 0,
        axis=1,
    )

    def cobertura(tipo):
        if tipo == "ROTOR":
            return cob_rotor
        if tipo == "ESTATOR":
            return cob_estator
        return max(cob_rotor, cob_estator)

    df["COBERTURA_ALVO_MESES"] = df["TIPO_ITEM"].map(cobertura)
    df["MINIMO_SUGERIDO_CONSUMO"] = np.ceil(df["MEDIA_MENSAL"] * df["COBERTURA_ALVO_MESES"])
    df["MINIMO_SUGERIDO_ANUAL"] = np.maximum(df["VMINIMO_BASE"], df["MINIMO_SUGERIDO_CONSUMO"])
    df["GAP_REPOSICAO"] = (df["MINIMO_SUGERIDO_ANUAL"] - df["ESTOQUE_ATUAL"]).clip(lower=0)

    df["STATUS"] = df.apply(
        lambda r: "RUPTURA" if r["ESTOQUE_ATUAL"] <= 0 else ("ABAIXO" if r["ESTOQUE_ATUAL"] < r["MINIMO_SUGERIDO_ANUAL"] else "OK"),
        axis=1,
    )

    ordem = {"RUPTURA": 1, "ABAIXO": 2, "OK": 3}
    df["_ord"] = df["STATUS"].map(ordem).fillna(9)

    df = df.sort_values(["_ord", "GAP_REPOSICAO", "MEDIA_MENSAL"], ascending=[True, False, False]).drop(columns=["_ord"])

    cols = [
        "MATERIAL",
        "DESCRICAO",
        "TIPO_ITEM",
        "ESTOQUE_ATUAL",
        "VMINIMO_BASE",
        "MEDIA_MENSAL",
        "MINIMO_SUGERIDO_ANUAL",
        "GAP_REPOSICAO",
        "STATUS",
    ]
    return df[cols].copy()


def montar_aba_sazonalidade(
    base: pd.DataFrame,
    consumo: pd.DataFrame,
    estoque: pd.DataFrame,
    cob_rotor: float,
    cob_estator: float,
    piso_base_fracao: float,
):
    itens = base[base["TIPO_ITEM"].isin(["ROTOR", "ESTATOR"])].copy()
    itens = itens.merge(estoque, on="MATERIAL", how="left")
    itens["ESTOQUE_ATUAL"] = pd.to_numeric(itens["ESTOQUE_ATUAL"], errors="coerce").fillna(0)

    if itens.empty:
        return pd.DataFrame(
            columns=["MATERIAL", "DESCRICAO", "TIPO_ITEM", "ESTOQUE_ATUAL", "VMINIMO_BASE"]
            + [MESES_NOME[m] for m in range(1, 13)]
        )

    def cobertura(tipo):
        if tipo == "ROTOR":
            return cob_rotor
        if tipo == "ESTATOR":
            return cob_estator
        return max(cob_rotor, cob_estator)

    itens["COBERTURA_ALVO_MESES"] = itens["TIPO_ITEM"].map(cobertura)
    piso_base_fracao = float(np.clip(piso_base_fracao, 0.0, 1.0))
    itens["PISO_BASE_MENSAL"] = np.ceil(itens["VMINIMO_BASE"] * piso_base_fracao)

    c = consumo.merge(itens[["MATERIAL"]], on="MATERIAL", how="inner")

    if c.empty:
        for m in range(1, 13):
            itens[MESES_NOME[m]] = itens["PISO_BASE_MENSAL"]
    else:
        # Media de consumo por material em cada mes do calendario, ao longo do historico.
        media_mes_item = (
            c.groupby(["MATERIAL", "MES"], dropna=False)["CONSUMO_QTDE"]
            .mean()
            .reset_index()
            .pivot(index="MATERIAL", columns="MES", values="CONSUMO_QTDE")
            .reset_index()
        )

        itens = itens.merge(media_mes_item, on="MATERIAL", how="left")

        for m in range(1, 13):
            if m not in itens.columns:
                itens[m] = 0

            min_mes = np.ceil(pd.to_numeric(itens[m], errors="coerce").fillna(0) * itens["COBERTURA_ALVO_MESES"])
            itens[MESES_NOME[m]] = np.maximum(itens["PISO_BASE_MENSAL"], min_mes)

    out_cols = ["MATERIAL", "DESCRICAO", "TIPO_ITEM", "ESTOQUE_ATUAL", "VMINIMO_BASE"] + [MESES_NOME[m] for m in range(1, 13)]
    out = itens[out_cols].copy()
    out = out.sort_values(["TIPO_ITEM", "DESCRICAO", "MATERIAL"], ascending=[True, True, True]).reset_index(drop=True)
    return out


def gerar_excel(aba_minimo: pd.DataFrame, aba_sazonal: pd.DataFrame) -> Path:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = OUTPUT_DIR / f"relatorio_estoque_diretor_{ts}.xlsx"

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        aba_minimo.to_excel(writer, sheet_name="Minimo_Sugerido_Anual", index=False)
        aba_sazonal.to_excel(writer, sheet_name="Sazonalidade", index=False)

    return out


def main():
    args = parse_args()

    base = carregar_base(Path(args.arquivo_base))
    engine = get_engine()
    consumo = query_consumo(engine, args.meses_historico)
    estoque = query_estoque(engine, args.tp_estoque)

    aba_minimo = montar_aba_minimo(base, consumo, estoque, args.cobertura_rotor, args.cobertura_estator)
    aba_sazonal = montar_aba_sazonalidade(
        base,
        consumo,
        estoque,
        args.cobertura_rotor,
        args.cobertura_estator,
        args.piso_base_fracao,
    )

    out = gerar_excel(aba_minimo, aba_sazonal)

    print("Relatorio diretor gerado")
    print(f"Arquivo: {out}")
    print(f"Itens: {len(aba_minimo)}")
    print(f"Ruptura: {(aba_minimo['STATUS'] == 'RUPTURA').sum()}")
    print(f"Abaixo: {(aba_minimo['STATUS'] == 'ABAIXO').sum()}")


if __name__ == "__main__":
    main()
