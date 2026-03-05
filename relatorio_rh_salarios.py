from datetime import datetime
from pathlib import Path
import argparse

import pandas as pd

from src.conexao import get_engine

OUTPUT_DIR = Path("exports")
OUTPUT_DIR.mkdir(exist_ok=True)


def carregar_quadro_atual(engine) -> pd.DataFrame:
    sql = """
    SELECT
        c.CODIGO AS CONTRATACAO,
        c.FUNCIONARIO,
        f.NOME AS NOME_FUNCIONARIO,
        f.CKATIVO AS ATIVO_FUNC,
        c.CKATIVO AS ATIVO_CONTRATO,
        c.DTADMISSAO,
        c.DTDEMISSAO,
        c.DEPARTAMENTO,
        c.SETOR,
        c.CARGO,
        c.FUNCAO,
        c.SALREF,
        c.VLRSALARIO1,
        c.VLRSALARIO2,
        c.FILIAL
    FROM RH_CONTRATACAO c
    JOIN RH_FUNCIONARIO f ON f.CODIGO = c.FUNCIONARIO
    """
    df = pd.read_sql(sql, engine)

    for col in ["VLRSALARIO1", "VLRSALARIO2"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df["SALARIO_REFERENCIA"] = df["VLRSALARIO1"].fillna(df["VLRSALARIO2"])
    df["DTADMISSAO"] = pd.to_datetime(df["DTADMISSAO"], errors="coerce")
    df["DTDEMISSAO"] = pd.to_datetime(df["DTDEMISSAO"], errors="coerce")

    return df


def carregar_historico_salarial(engine) -> pd.DataFrame:
    sql = """
    SELECT
        hs.CODIGO,
        hs.CONTRATACAO,
        hs.FUNCIONARIO,
        hs.DTCADASTRO,
        hs.USERDATE,
        hs.STATUS,
        hs.AUTORIZADO,
        hs.VLRSALARIO1,
        hs.VLRSALARIO2,
        hs.MOTIVO,
        hs.SALREF,
        hs.ORIGEM,
        c.DEPARTAMENTO,
        c.SETOR,
        c.CARGO,
        c.FUNCAO,
        f.NOME AS NOME_FUNCIONARIO,
        f.CKATIVO AS ATIVO_FUNC
    FROM RH_CONTRATACAOSALARIO hs
    LEFT JOIN RH_CONTRATACAO c ON c.CODIGO = hs.CONTRATACAO
    LEFT JOIN RH_FUNCIONARIO f ON f.CODIGO = hs.FUNCIONARIO
    """
    df = pd.read_sql(sql, engine)

    for col in ["VLRSALARIO1", "VLRSALARIO2"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df["SALARIO_EVENTO"] = df["VLRSALARIO1"].fillna(df["VLRSALARIO2"])
    df["DT_EVENTO"] = pd.to_datetime(df["DTCADASTRO"], errors="coerce")
    df["USERDATE"] = pd.to_datetime(df["USERDATE"], errors="coerce")

    df = df.sort_values(["CONTRATACAO", "DT_EVENTO", "CODIGO"]).reset_index(drop=True)
    df["SALARIO_ANTERIOR"] = df.groupby("CONTRATACAO")["SALARIO_EVENTO"].shift(1)
    df["DELTA_SALARIO"] = df["SALARIO_EVENTO"] - df["SALARIO_ANTERIOR"]
    df["DELTA_PCT"] = (df["DELTA_SALARIO"] / df["SALARIO_ANTERIOR"]) * 100

    return df


def gerar_resumo(quadro: pd.DataFrame, historico: pd.DataFrame, meses: int):
    hoje = pd.Timestamp.today().normalize()
    dt_corte = hoje - pd.DateOffset(months=meses)

    quadro_ativo = quadro[(quadro["ATIVO_CONTRATO"] == "S") | (quadro["DTDEMISSAO"].isna())].copy()

    resumo_geral = pd.DataFrame(
        {
            "Metrica": [
                "Funcionarios_unicos",
                "Contratos_ativos",
                "Salario_referencia_total",
                "Salario_referencia_medio",
                f"Eventos_salariais_{meses}m",
                f"Reajustes_positivos_{meses}m",
            ],
            "Valor": [
                quadro_ativo["FUNCIONARIO"].nunique(),
                quadro_ativo["CONTRATACAO"].nunique(),
                quadro_ativo["SALARIO_REFERENCIA"].sum(min_count=1),
                quadro_ativo["SALARIO_REFERENCIA"].mean(),
                historico[historico["DT_EVENTO"] >= dt_corte].shape[0],
                historico[(historico["DT_EVENTO"] >= dt_corte) & (historico["DELTA_SALARIO"] > 0)].shape[0],
            ],
        }
    )

    resumo_setor = (
        quadro_ativo.groupby(["DEPARTAMENTO", "SETOR"], dropna=False)
        .agg(
            Funcionarios=("FUNCIONARIO", "nunique"),
            Contratos=("CONTRATACAO", "nunique"),
            Salario_Total=("SALARIO_REFERENCIA", "sum"),
            Salario_Medio=("SALARIO_REFERENCIA", "mean"),
        )
        .reset_index()
        .sort_values("Salario_Total", ascending=False)
    )

    reajustes_periodo = historico[historico["DT_EVENTO"] >= dt_corte].copy()
    reajustes_periodo = reajustes_periodo.sort_values(["DT_EVENTO", "DELTA_SALARIO"], ascending=[False, False])

    maiores_altas = reajustes_periodo[reajustes_periodo["DELTA_SALARIO"] > 0].head(100).copy()

    return resumo_geral, resumo_setor, reajustes_periodo, maiores_altas, quadro_ativo


def gerar_alertas(
    quadro_ativo: pd.DataFrame,
    historico: pd.DataFrame,
    meses_sem_reajuste: int,
    limite_delta_pct: float,
):
    hoje = pd.Timestamp.today().normalize()
    dt_corte_sem_reajuste = hoje - pd.DateOffset(months=meses_sem_reajuste)

    base_eventos = historico.dropna(subset=["DT_EVENTO", "SALARIO_EVENTO"]).copy()

    quedas = base_eventos[base_eventos["DELTA_SALARIO"] < 0].copy()
    quedas = quedas.sort_values(["DT_EVENTO", "DELTA_SALARIO"], ascending=[False, True])

    fora_curva = base_eventos[
        (base_eventos["DELTA_SALARIO"].notna())
        & (base_eventos["SALARIO_ANTERIOR"].notna())
        & (base_eventos["DELTA_PCT"].abs() >= limite_delta_pct)
    ].copy()
    fora_curva = fora_curva.sort_values("DELTA_PCT", key=lambda s: s.abs(), ascending=False)

    ultimo_evento = (
        base_eventos.sort_values(["CONTRATACAO", "DT_EVENTO", "CODIGO"])
        .groupby("CONTRATACAO", as_index=False)
        .tail(1)
        [["CONTRATACAO", "DT_EVENTO", "SALARIO_EVENTO", "DELTA_SALARIO", "DELTA_PCT"]]
        .rename(
            columns={
                "DT_EVENTO": "DT_ULTIMO_REAJUSTE",
                "SALARIO_EVENTO": "SALARIO_ULTIMO_EVENTO",
                "DELTA_SALARIO": "DELTA_ULTIMO_EVENTO",
                "DELTA_PCT": "DELTA_PCT_ULTIMO_EVENTO",
            }
        )
    )

    sem_reajuste = quadro_ativo.merge(ultimo_evento, on="CONTRATACAO", how="left")
    sem_reajuste = sem_reajuste[
        sem_reajuste["DT_ULTIMO_REAJUSTE"].isna() | (sem_reajuste["DT_ULTIMO_REAJUSTE"] < dt_corte_sem_reajuste)
    ].copy()
    sem_reajuste["MESES_SEM_REAJUSTE"] = (
        (hoje - sem_reajuste["DT_ULTIMO_REAJUSTE"]).dt.days / 30.44
    ).round(1)
    sem_reajuste.loc[sem_reajuste["DT_ULTIMO_REAJUSTE"].isna(), "MESES_SEM_REAJUSTE"] = None
    sem_reajuste = sem_reajuste.sort_values(["MESES_SEM_REAJUSTE", "DTADMISSAO"], ascending=[False, True])

    return quedas, fora_curva, sem_reajuste


def main():
    parser = argparse.ArgumentParser(description="Relatório inicial RH - salários e reajustes")
    parser.add_argument("--meses", type=int, default=12, help="Janela de meses para destacar reajustes")
    parser.add_argument(
        "--meses-sem-reajuste",
        type=int,
        default=12,
        help="Considera alerta quando último reajuste é mais antigo que esta janela",
    )
    parser.add_argument(
        "--limite-delta-pct",
        type=float,
        default=30.0,
        help="Limite percentual absoluto para marcar reajuste fora da curva",
    )
    args = parser.parse_args()

    engine = get_engine()

    quadro = carregar_quadro_atual(engine)
    historico = carregar_historico_salarial(engine)

    resumo_geral, resumo_setor, reajustes_periodo, maiores_altas, quadro_ativo = gerar_resumo(
        quadro, historico, args.meses
    )
    quedas, fora_curva, sem_reajuste = gerar_alertas(
        quadro_ativo,
        historico,
        args.meses_sem_reajuste,
        args.limite_delta_pct,
    )

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_file = OUTPUT_DIR / f"rh_salarios_{ts}.xlsx"

    with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
        resumo_geral.to_excel(writer, sheet_name="Resumo_Geral", index=False)
        resumo_setor.to_excel(writer, sheet_name="Resumo_Setor", index=False)
        quadro_ativo.to_excel(writer, sheet_name="Quadro_Ativo", index=False)
        historico.to_excel(writer, sheet_name="Historico_Salarial", index=False)
        reajustes_periodo.to_excel(writer, sheet_name=f"Eventos_{args.meses}m", index=False)
        maiores_altas.to_excel(writer, sheet_name="Top_Altas", index=False)
        quedas.to_excel(writer, sheet_name="Alerta_Queda", index=False)
        fora_curva.to_excel(writer, sheet_name="Alerta_ForaCurva", index=False)
        sem_reajuste.to_excel(writer, sheet_name="Alerta_SemReajuste", index=False)

    print("Relatório RH concluído")
    print(f"Arquivo: {out_file}")
    print(f"Quadro ativo (linhas): {len(quadro_ativo)}")
    print(f"Histórico salarial (linhas): {len(historico)}")
    print(f"Eventos últimos {args.meses}m: {len(reajustes_periodo)}")
    print(f"Alertas queda salarial: {len(quedas)}")
    print(f"Alertas fora da curva (|delta%| >= {args.limite_delta_pct:.1f}%): {len(fora_curva)}")
    print(f"Alertas sem reajuste > {args.meses_sem_reajuste}m: {len(sem_reajuste)}")


if __name__ == "__main__":
    main()
