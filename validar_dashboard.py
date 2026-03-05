import argparse
from datetime import datetime
from pathlib import Path

import pandas as pd

import fase4_dashboard as dash
from src.conexao import get_engine

OUTPUT_DIR = Path("exports")
OUTPUT_DIR.mkdir(exist_ok=True)


def _status(ok: bool, alerta: bool = False) -> str:
    if ok:
        return "OK"
    return "ALERTA" if alerta else "FALHA"


def _add(rows: list[dict], modo: str, check: str, ok: bool, detalhe: str, alerta: bool = False):
    rows.append(
        {
            "Modo": modo,
            "Check": check,
            "Status": _status(ok, alerta=alerta),
            "Detalhe": detalhe,
        }
    )


def _validar_modo(engine, sem_contrato: bool) -> tuple[list[dict], dict]:
    rows: list[dict] = []
    dados, modo = dash.coletar_dados_dashboard(engine, sem_contrato=sem_contrato)

    cart = dados["df_cart"]
    cot = dados["df_cotacoes"]
    sla = dados["sla_hoje"]
    mtd = dados["mtd_atual"]

    _add(rows, modo, "Carteira carregada", len(cart) >= 0, f"Itens carteira: {len(cart)}")

    semaforo_total = int(cart["Semaforo"].isin(["Atrasado", "Urgente", "No prazo", "Sem prazo"]).sum()) if not cart.empty else 0
    _add(
        rows,
        modo,
        "Semáforo consistente",
        semaforo_total == len(cart),
        f"Marcados no semáforo: {semaforo_total} | Total carteira: {len(cart)}",
    )

    l_carteira = int((cart["Status_Item"] == "L").sum()) if not cart.empty else 0
    _add(
        rows,
        modo,
        "SLA backlog vs carteira",
        int(sla["backlog_itens"]) == l_carteira,
        f"Backlog SLA: {int(sla['backlog_itens'])} | Itens L carteira: {l_carteira}",
    )

    _add(
        rows,
        modo,
        "SLA não negativo",
        min(
            int(sla["entradas_itens"]),
            int(sla["resolvidas_itens"]),
            int(sla["backlog_itens"]),
        ) >= 0,
        f"Entradas: {sla['entradas_itens']} | Resolvidas: {sla['resolvidas_itens']} | Backlog: {sla['backlog_itens']}",
    )

    _add(
        rows,
        modo,
        "MTD >= Hoje",
        float(mtd["faturamento_mtd"]) >= float(mtd["faturamento_hoje"]),
        f"MTD: {mtd['faturamento_mtd']:.2f} | Hoje: {mtd['faturamento_hoje']:.2f}",
    )

    _add(
        rows,
        modo,
        "Projeção MTD consistente",
        float(mtd["run_rate_dia"]) >= 0 and float(mtd["projecao_mes"]) >= 0,
        f"Run-rate: {mtd['run_rate_dia']:.2f} | Projeção: {mtd['projecao_mes']:.2f}",
    )

    if cot.empty:
        _add(rows, modo, "Cotações no mês fechado", False, "Sem dados de cotações no período.", alerta=True)
    else:
        status_class = cot["Status_Orc"].map(dash.classificar_status_orc)
        encerr = int((status_class == "ENCERRADA").sum())
        perd = int((status_class == "PERDIDA").sum())
        abertas = int((status_class == "ABERTA").sum())
        total = int(len(cot))

        _add(
            rows,
            modo,
            "Status de cotações consistente",
            (encerr + perd + abertas) == total,
            f"Enc:{encerr} Perd:{perd} Aber:{abertas} | Total:{total}",
        )

        conv_geral = (encerr / total * 100.0) if total else 0.0
        finalizadas = encerr + perd
        conv_final = (encerr / finalizadas * 100.0) if finalizadas else 0.0

        _add(
            rows,
            modo,
            "Conversão geral em faixa",
            0.0 <= conv_geral <= 100.0,
            f"Conversão geral: {conv_geral:.2f}%",
        )
        _add(
            rows,
            modo,
            "Conversão finalizada em faixa",
            0.0 <= conv_final <= 100.0,
            f"Conversão finalizada: {conv_final:.2f}%",
        )

        ano = dash.REF_ANO
        mes = dash.REF_MES
        ultimo = __import__("calendar").monthrange(ano, mes)[1]
        filtro_orc = dash.build_filtro_clientes_contrato_orc("f", ["PETROBRAS", "PETROLEO"] if sem_contrato else [])
        sql_aprov = f"""
            SELECT COUNT(*) AS QtdEncerradas
            FROM VE_ORCAMENTOS o
            LEFT JOIN FN_FORNECEDORES f ON f.CODIGO = o.CODCLIENTE
            WHERE o.DTCADASTRO BETWEEN '{ano}-{mes:02d}-01' AND '{ano}-{mes:02d}-{ultimo:02d}'
              AND o.FILIAL IN ({dash.FILIAIS_ORC_STR})
              AND o.STATUS = 'E'
              {filtro_orc}
        """
        qtd_sql = int(pd.read_sql(sql_aprov, engine).iloc[0]["QtdEncerradas"] or 0)
        _add(
            rows,
            modo,
            "Encerradas SQL vs dashboard",
            qtd_sql == encerr,
            f"SQL: {qtd_sql} | Dashboard: {encerr}",
        )

    resumo = {
        "modo": modo,
        "itens_carteira": int(len(cart)),
        "cotacoes_total": int(len(cot)),
        "sla_backlog": int(sla["backlog_itens"]),
        "mtd_faturamento": float(mtd["faturamento_mtd"]),
        "hoje_faturamento": float(mtd["faturamento_hoje"]),
    }
    return rows, resumo


def main():
    parser = argparse.ArgumentParser(description="Validação de consistência do dashboard comercial")
    parser.add_argument("--somente", choices=["base", "sem", "ambos"], default="ambos")
    args = parser.parse_args()

    engine = get_engine()

    modos = []
    if args.somente in ("base", "ambos"):
        modos.append(False)
    if args.somente in ("sem", "ambos"):
        modos.append(True)

    all_rows: list[dict] = []
    resumos: list[dict] = []

    for sem_contrato in modos:
        rows, resumo = _validar_modo(engine, sem_contrato=sem_contrato)
        all_rows.extend(rows)
        resumos.append(resumo)

    df_valid = pd.DataFrame(all_rows)
    df_resumo = pd.DataFrame(resumos)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo = OUTPUT_DIR / f"validacao_dashboard_{ts}.xlsx"

    with pd.ExcelWriter(arquivo, engine="openpyxl") as writer:
        df_valid.to_excel(writer, sheet_name="Validacao", index=False)
        df_resumo.to_excel(writer, sheet_name="Resumo", index=False)

    ok = int((df_valid["Status"] == "OK").sum())
    alerta = int((df_valid["Status"] == "ALERTA").sum())
    falha = int((df_valid["Status"] == "FALHA").sum())

    print("Validação do dashboard concluída")
    print(f"Arquivo: {arquivo}")
    print(f"OK: {ok} | ALERTA: {alerta} | FALHA: {falha}")


if __name__ == "__main__":
    main()
