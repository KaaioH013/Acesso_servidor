from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

from relatorio_528_replicado import query_detalhe
from src.conexao import get_engine

OUTPUT_DIR = Path("exports")
OUTPUT_DIR.mkdir(exist_ok=True)


@dataclass(frozen=True)
class JanelaAnalise:
    dt_ini: date
    dt_fim: date


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Analise de faturamento (MoM/YoY) + previsao do proximo mes"
    )
    p.add_argument("--usuario", type=int, default=124, help="Codigo do usuario nos filtros do ERP")
    p.add_argument("--anos", type=int, default=5, help="Quantidade de anos para historico")
    p.add_argument("--formato", choices=["xlsx", "csv", "ambos"], default="ambos")
    return p.parse_args()


def _primeiro_dia_mes(d: date) -> date:
    return d.replace(day=1)


def _add_months(d: date, months: int) -> date:
    ts = pd.Timestamp(d)
    out = ts + pd.DateOffset(months=months)
    return date(out.year, out.month, out.day)


def calcular_janela(anos: int) -> JanelaAnalise:
    hoje = date.today()
    inicio_mes_atual = _primeiro_dia_mes(hoje)
    dt_fim = inicio_mes_atual - timedelta(days=1)
    inicio_ultimo_mes = _primeiro_dia_mes(dt_fim)
    dt_ini = _add_months(inicio_ultimo_mes, -(anos * 12 - 1))
    return JanelaAnalise(dt_ini=dt_ini, dt_fim=dt_fim)


def montar_base_mensal(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["Dt_Emissao"] = pd.to_datetime(out["Dt_Emissao"])
    out["MesRef"] = out["Dt_Emissao"].dt.to_period("M").dt.to_timestamp()

    out["Vlr_Saidas"] = np.where(out["Tipo"] == "S", out["Vlr_Total_Nota"], 0.0)
    out["Vlr_Devolucoes"] = np.where(out["Tipo"] == "D", out["Vlr_Total_Nota"], 0.0)
    out["Vlr_Liquido"] = out["Vlr_Saidas"] - out["Vlr_Devolucoes"]

    mensal = (
        out.groupby("MesRef", as_index=False)
        .agg(
            Saidas=("Vlr_Saidas", "sum"),
            Devolucoes=("Vlr_Devolucoes", "sum"),
            Faturamento_Liquido=("Vlr_Liquido", "sum"),
        )
        .sort_values("MesRef")
    )

    mensal["Ano"] = mensal["MesRef"].dt.year
    mensal["Mes"] = mensal["MesRef"].dt.month
    mensal["Mes_Ano"] = mensal["MesRef"].dt.strftime("%m/%Y")
    mensal["Crescimento_MoM_%"] = mensal["Faturamento_Liquido"].pct_change() * 100.0
    mensal["Crescimento_YoY_%"] = mensal["Faturamento_Liquido"].pct_change(12) * 100.0
    return mensal


def _matriz_regressao(mes_ref: pd.Series) -> np.ndarray:
    n = len(mes_ref)
    trend = np.arange(n, dtype=float)
    meses = mes_ref.dt.month.to_numpy()

    cols = [np.ones(n), trend]
    for m in range(2, 13):
        cols.append((meses == m).astype(float))
    return np.column_stack(cols)


def prever_proximo_mes(mensal: pd.DataFrame) -> pd.DataFrame:
    serie = mensal[["MesRef", "Faturamento_Liquido"]].copy()
    y = serie["Faturamento_Liquido"].to_numpy(dtype=float)
    X = _matriz_regressao(serie["MesRef"])

    beta, *_ = np.linalg.lstsq(X, y, rcond=None)
    y_hat = X @ beta
    resid = y - y_hat

    # Desvio padrao residual para intervalo pratico de previsao.
    sigma = float(np.std(resid, ddof=1)) if len(resid) > 1 else 0.0

    prox_mes_ref = (serie["MesRef"].max() + pd.offsets.MonthBegin(1)).to_pydatetime()
    prox_mes = pd.Series(pd.to_datetime([prox_mes_ref]))
    X_next = _matriz_regressao(prox_mes)
    pred = float((X_next @ beta).ravel()[0])

    hist = serie.set_index("MesRef")["Faturamento_Liquido"]
    prev_12 = hist.get(pd.Timestamp(prox_mes_ref) - pd.DateOffset(years=1), np.nan)
    yoy_prev = ((pred / prev_12) - 1.0) * 100.0 if pd.notna(prev_12) and prev_12 != 0 else np.nan

    out = pd.DataFrame(
        [
            {
                "Mes_Previsao": pd.Timestamp(prox_mes_ref).strftime("%m/%Y"),
                "Faturamento_Previsto": pred,
                "Limite_Inferior_95": pred - 1.96 * sigma,
                "Limite_Superior_95": pred + 1.96 * sigma,
                "Faturamento_Mes_Ano_Anterior": float(prev_12) if pd.notna(prev_12) else np.nan,
                "Crescimento_YoY_Previsto_%": yoy_prev,
            }
        ]
    )
    return out


def salvar_relatorios(mensal: pd.DataFrame, previsao: pd.DataFrame, janela: JanelaAnalise, formato: str) -> list[Path]:
    ts = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
    prefixo = f"analise_faturamento_{janela.dt_ini}_{janela.dt_fim}_{ts}"
    saidas: list[Path] = []

    if formato in ("xlsx", "ambos"):
        xlsx = OUTPUT_DIR / f"{prefixo}.xlsx"
        with pd.ExcelWriter(xlsx, engine="openpyxl") as writer:
            mensal.to_excel(writer, sheet_name="Base_Mensal", index=False)
            mensal[["MesRef", "Mes_Ano", "Faturamento_Liquido", "Crescimento_MoM_%"]].to_excel(
                writer, sheet_name="Crescimento_MoM", index=False
            )
            mensal[["MesRef", "Mes_Ano", "Faturamento_Liquido", "Crescimento_YoY_%"]].to_excel(
                writer, sheet_name="Crescimento_YoY", index=False
            )
            previsao.to_excel(writer, sheet_name="Previsao_Proximo_Mes", index=False)
        saidas.append(xlsx)

    if formato in ("csv", "ambos"):
        csv_base = OUTPUT_DIR / f"{prefixo}_base_mensal.csv"
        csv_prev = OUTPUT_DIR / f"{prefixo}_previsao.csv"
        mensal.to_csv(csv_base, index=False, sep=";", encoding="utf-8-sig")
        previsao.to_csv(csv_prev, index=False, sep=";", encoding="utf-8-sig")
        saidas.extend([csv_base, csv_prev])

    return saidas


def main() -> None:
    args = parse_args()
    janela = calcular_janela(args.anos)

    engine = get_engine()
    df = query_detalhe(engine, janela.dt_ini.isoformat(), janela.dt_fim.isoformat(), args.usuario)

    mensal = montar_base_mensal(df)
    previsao = prever_proximo_mes(mensal)
    saidas = salvar_relatorios(mensal, previsao, janela, args.formato)

    media_mom = float(mensal["Crescimento_MoM_%"].mean(skipna=True))
    media_yoy = float(mensal["Crescimento_YoY_%"].mean(skipna=True))
    prev = previsao.iloc[0]

    print("Analise de faturamento concluida")
    print(f"Janela analisada: {janela.dt_ini} a {janela.dt_fim}")
    print(f"Meses analisados: {len(mensal)}")
    print(f"Media Crescimento MoM: {media_mom:.2f}%")
    print(f"Media Crescimento YoY: {media_yoy:.2f}%")
    print(
        "Previsao proximo mes "
        f"({prev['Mes_Previsao']}): R$ {prev['Faturamento_Previsto']:,.2f} "
        f"(IC95%: R$ {prev['Limite_Inferior_95']:,.2f} a R$ {prev['Limite_Superior_95']:,.2f})"
    )
    if pd.notna(prev["Crescimento_YoY_Previsto_%"]):
        print(f"YoY previsto: {prev['Crescimento_YoY_Previsto_%']:.2f}%")
    print("Arquivos gerados:")
    for arq in saidas:
        print(f" - {arq}")


if __name__ == "__main__":
    main()
