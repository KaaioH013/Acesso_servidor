import argparse
from dataclasses import dataclass

from relatorio_528_replicado import montar_resumo, query_detalhe
from src.conexao import get_engine


@dataclass(frozen=True)
class CasoRegressao:
    nome: str
    dt_ini: str
    dt_fim: str
    usuario: int
    saidas: float
    devolucoes: float
    liquido: float


# Baseline:
# - jan_2026 confirmado contra CSV do 528 exportado pelo usuario.
# - demais meses fechados foram congelados a partir da mesma logica validada.
CASOS_PADRAO = [
    CasoRegressao(
        nome="nov_2025",
        dt_ini="2025-11-01",
        dt_fim="2025-11-30",
        usuario=124,
        saidas=2492916.61,
        devolucoes=21422.66,
        liquido=2471493.95,
    ),
    CasoRegressao(
        nome="dez_2025",
        dt_ini="2025-12-01",
        dt_fim="2025-12-31",
        usuario=124,
        saidas=1960362.18,
        devolucoes=14509.34,
        liquido=1945852.84,
    ),
    CasoRegressao(
        nome="jan_2026",
        dt_ini="2026-01-01",
        dt_fim="2026-01-31",
        usuario=124,
        saidas=2330512.18,
        devolucoes=22722.58,
        liquido=2307789.60,
    ),
    CasoRegressao(
        nome="fev_2026",
        dt_ini="2026-02-01",
        dt_fim="2026-02-28",
        usuario=124,
        saidas=1870412.43,
        devolucoes=30280.74,
        liquido=1840131.69,
    ),
]


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Valida regressao do relatorio 528")
    p.add_argument("--tolerancia", type=float, default=0.01, help="Diferenca maxima aceita por total")
    return p.parse_args()


def _extrair_totais(resumo):
    saidas = float(resumo.loc[resumo["Tipo"] == "S", "Vlr_Total_Nota"].sum())
    devol = float(resumo.loc[resumo["Tipo"] == "D", "Vlr_Total_Nota"].sum())
    liquido = float(resumo.loc[resumo["Tipo"] == "TOTAL_LIQ", "Vlr_Total_Nota"].sum())
    return saidas, devol, liquido


def _ok(esperado: float, obtido: float, tolerancia: float) -> bool:
    return abs(esperado - obtido) <= tolerancia


def main() -> int:
    args = parse_args()
    engine = get_engine()

    falhas = 0
    print("Iniciando regressao 528...")

    for caso in CASOS_PADRAO:
        df = query_detalhe(engine, caso.dt_ini, caso.dt_fim, caso.usuario)
        resumo = montar_resumo(df)
        saidas, devol, liquido = _extrair_totais(resumo)

        ok_s = _ok(caso.saidas, saidas, args.tolerancia)
        ok_d = _ok(caso.devolucoes, devol, args.tolerancia)
        ok_l = _ok(caso.liquido, liquido, args.tolerancia)

        status = "OK" if (ok_s and ok_d and ok_l) else "FALHA"
        print(f"\n[{status}] Caso: {caso.nome}")
        print(f"  Saidas      esperado={caso.saidas:,.2f} obtido={saidas:,.2f}")
        print(f"  Devolucoes  esperado={caso.devolucoes:,.2f} obtido={devol:,.2f}")
        print(f"  Total Liqu. esperado={caso.liquido:,.2f} obtido={liquido:,.2f}")

        if not (ok_s and ok_d and ok_l):
            falhas += 1

    if falhas:
        print(f"\nRegressao 528 finalizada com {falhas} falha(s).")
        return 1

    print("\nRegressao 528 finalizada sem divergencias.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
