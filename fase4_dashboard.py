"""
fase4_dashboard.py — Dashboard HTML Interativo (Fase 4 do Roadmap)

Gera um único arquivo dashboard.html com:
  - Cards KPI: carteira, faturamento mês, margem média, alertas, aguardando NF
    - Tabela operacional: prioridades do dia para ação comercial
  - Gráfico de linha: evolução mensal 24 meses com comparativo YoY
  - Gráfico de barras horizontais: ranking de vendedores
  - Treemap: participação dos clientes (Curva ABC)
  - Rosca: semáforo da carteira (atrasado / urgente / no prazo)
  - Tabela: top 10 itens aguardando NF (mais antigos primeiro)
  - Tabela: alertas críticos de margem

Uso:
    python fase4_dashboard.py           # abre no navegador ao gerar
    python fase4_dashboard.py --no-open # só gera o arquivo
    python fase4_dashboard.py --sem-contrato --no-open # exclui clientes de contrato (Petro*)

Arquivo gerado: exports/dashboard.html  (sobrescreve sempre — versão mais recente)
"""

import sys
import argparse
import calendar
from datetime import datetime, date
from pathlib import Path
from functools import lru_cache

import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots

sys.path.insert(0, "src")
from conexao import get_engine

# ── Paleta de cores ───────────────────────────────────────────────────────────
COR_AZUL       = "#1F4E79"
COR_AZUL_CLARO = "#2E75B6"
COR_VERDE      = "#70AD47"
COR_AMARELO    = "#FFC000"
COR_VERMELHO   = "#C00000"
COR_LARANJA    = "#ED7D31"
COR_CINZA      = "#A6A6A6"
COR_BG         = "#F2F2F2"

# ── Filtros padrão ────────────────────────────────────────────────────────────
TPVENDA_EXCLUIR = (7, 21, 5, 12, 24, 11, 26, 6, 15, 16, 8, 19, 9, 17, 53, 18, 65, 23)
TPVENDA_STR = ",".join(str(x) for x in TPVENDA_EXCLUIR)
FILIAIS_ORC = (1, 2)
FILIAIS_ORC_STR = ",".join(str(x) for x in FILIAIS_ORC)

FILTROS_PECAS = f"""
    AND p.STATUS <> 'C'
    AND i.STATUS NOT IN ('C','F')
    AND i.TPVENDA NOT IN ({TPVENDA_STR})
    AND i.MATERIAL NOT LIKE '8%'
    AND i.FLAGSUB <> 'S'
    AND p.CODIGO NOT IN (
        SELECT p2.CODIGO FROM VE_PEDIDO p2
        JOIN FN_FORNECEDORES f2 ON f2.CODIGO = p2.CLIENTE
        WHERE f2.UF = 'EX'
    )
"""
FILTROS_PECAS_HIST = f"""
    AND p.STATUS <> 'C'
    AND i.STATUS <> 'C'
    AND i.TPVENDA NOT IN ({TPVENDA_STR})
    AND i.MATERIAL NOT LIKE '8%'
    AND i.FLAGSUB <> 'S'
    AND p.CODIGO NOT IN (
        SELECT p2.CODIGO FROM VE_PEDIDO p2
        JOIN FN_FORNECEDORES f2 ON f2.CODIGO = p2.CLIENTE
        WHERE f2.UF = 'EX'
    )
"""


def build_filtro_clientes_contrato(padroes: list[str]) -> str:
    if not padroes:
        return ""

    condicoes = []
    for p in padroes:
        termo = p.strip().upper()
        if not termo:
            continue
        termo_sql = termo.replace("'", "''")
        condicoes.append(f"UPPER(ISNULL(f3.RAZAO, '')) LIKE '%{termo_sql}%' ")
        condicoes.append(f"UPPER(ISNULL(f3.FANTASIA, '')) LIKE '%{termo_sql}%' ")

    if not condicoes:
        return ""

    return """
    AND p.CODIGO NOT IN (
        SELECT p3.CODIGO FROM VE_PEDIDO p3
        JOIN FN_FORNECEDORES f3 ON f3.CODIGO = p3.CLIENTE
        WHERE {condicoes_sql}
    )
""".format(condicoes_sql=" OR ".join(condicoes))


def build_filtro_clientes_contrato_orc(alias_cliente: str, padroes: list[str]) -> str:
    if not padroes:
        return ""

    condicoes = []
    for p in padroes:
        termo = p.strip().upper()
        if not termo:
            continue
        termo_sql = termo.replace("'", "''")
        condicoes.append(f"UPPER(ISNULL({alias_cliente}.RAZAO, '')) LIKE '%{termo_sql}%' ")
        condicoes.append(f"UPPER(ISNULL({alias_cliente}.FANTASIA, '')) LIKE '%{termo_sql}%' ")

    if not condicoes:
        return ""

    return "AND NOT (" + " OR ".join(condicoes) + ")"

HOJE = date.today()
OUTPUT = Path("exports")
OUTPUT.mkdir(exist_ok=True)

if HOJE.month == 1:
    REF_MES = 12
    REF_ANO = HOJE.year - 1
else:
    REF_MES = HOJE.month - 1
    REF_ANO = HOJE.year

# ── Metas mensais (configurar aqui) ─────────────────────────────────────────
# Valor mensal médio esperado por ano (meta anual ÷ 12 é uma opção,
# mas aqui usamos a meta média mensal que o time trabalha)
META_MENSAL = {
    REF_ANO - 1: 2_200_000,   # meta ano anterior
    REF_ANO:     2_500_000,   # meta ano de referência
}

DATA_CORTE_REGRAS_EXTERNOS = date(2026, 1, 15)
ARQ_CIDADES_EXCLUIDAS_MG_ALEXANDRE_XLSX = Path("Cidades_nao_atendidas_mg_alexandre.xlsx")
ARQ_CIDADES_EXCLUIDAS_MG_ALEXANDRE_XLSM = Path("Cidades_nao_atendidas_mg_alexandre.xlsm")


def normalizar_texto(valor: str) -> str:
    if valor is None:
        return ""
    return " ".join(str(valor).strip().upper().split())


@lru_cache(maxsize=1)
def carregar_cidades_excluidas_mg_alexandre() -> set[str]:
    arquivo = None
    if ARQ_CIDADES_EXCLUIDAS_MG_ALEXANDRE_XLSX.exists():
        arquivo = ARQ_CIDADES_EXCLUIDAS_MG_ALEXANDRE_XLSX
    elif ARQ_CIDADES_EXCLUIDAS_MG_ALEXANDRE_XLSM.exists():
        arquivo = ARQ_CIDADES_EXCLUIDAS_MG_ALEXANDRE_XLSM

    if arquivo is None:
        return set()

    try:
        xls = pd.ExcelFile(arquivo)
        df = pd.read_excel(arquivo, sheet_name=xls.sheet_names[0])
        if df.empty:
            return set()

        col_cidade = None
        for c in df.columns:
            if "CIDAD" in normalizar_texto(c):
                col_cidade = c
                break
        if col_cidade is None:
            col_cidade = df.columns[0]

        cidades = (
            df[col_cidade]
            .dropna()
            .map(normalizar_texto)
            .loc[lambda s: s != ""]
        )
        return set(cidades.tolist())
    except Exception:
        return set()


def mapear_representante_externo(cliente_uf: str, cliente_cidade: str, dt_pedido) -> str:
    uf = normalizar_texto(cliente_uf)
    cidade = normalizar_texto(cliente_cidade)
    dt_ref = pd.to_datetime(dt_pedido, errors="coerce")
    dt_ok = dt_ref.date() if pd.notna(dt_ref) else None

    nordeste = {"AL", "BA", "CE", "MA", "PB", "PE", "PI", "RN", "SE"}
    cidades_excluidas_mg = carregar_cidades_excluidas_mg_alexandre()

    if uf == "PR":
        return "RAFAEL LUCATELI"
    if uf in {"GO", "MT"}:
        return "SEBASTIÃO"
    if uf == "PA":
        return "HELIFER"
    if uf in {"RS", "SC"}:
        return "MARCELO BENTO"
    if uf in nordeste:
        return "ALEX MENDONÇA"
    if uf == "MS" and dt_ok is not None and dt_ok >= DATA_CORTE_REGRAS_EXTERNOS:
        return "FERNANDO GUIDELI"
    if uf == "MG" and dt_ok is not None and dt_ok >= DATA_CORTE_REGRAS_EXTERNOS:
        if cidade and cidade in cidades_excluidas_mg:
            return "HELIBOMBAS"
        return "ALEXANDRE DURAZZO"

    return "HELIBOMBAS"


def classificar_status_orc(status) -> str:
    s = normalizar_texto(status)
    if s == "E":
        return "ENCERRADA"
    if s == "O":
        return "ABERTA"
    if s in {"P", "X", "C"}:
        return "PERDIDA"
    return "ABERTA"


# ── Coleta de dados ───────────────────────────────────────────────────────────

def get_evolucao_mensal(engine, filtro_cliente_extra: str = "") -> pd.DataFrame:
    ano_fim = REF_ANO
    ano_ini = ano_fim - 1
    sql = f"""
        SELECT
            YEAR(p.DTPEDIDO)  AS Ano,
            MONTH(p.DTPEDIDO) AS Mes,
            SUM(i.VLRTOTAL)   AS Faturamento,
            COUNT(DISTINCT p.CODIGO) AS Pedidos
        FROM VE_PEDIDOITENS i
        JOIN VE_PEDIDO p ON p.CODIGO = i.PEDIDO
        WHERE p.DTPEDIDO BETWEEN '{ano_ini}-01-01' AND '{ano_fim}-12-31'
        {FILTROS_PECAS_HIST}
        {filtro_cliente_extra}
        GROUP BY YEAR(p.DTPEDIDO), MONTH(p.DTPEDIDO)
        ORDER BY Ano, Mes
    """
    df = pd.read_sql(sql, engine)
    df["Periodo"] = pd.to_datetime(dict(year=df["Ano"], month=df["Mes"], day=1))
    return df


def get_faturamento_mes_atual(engine, filtro_cliente_extra: str = "") -> dict:
    mes, ano = REF_MES, REF_ANO
    ultimo = calendar.monthrange(ano, mes)[1]
    sql = f"""
        SELECT
            COUNT(DISTINCT p.CODIGO) AS Pedidos,
            COUNT(i.CODIGO)          AS Itens,
            SUM(i.VLRTOTAL)          AS Faturamento
        FROM VE_PEDIDOITENS i
        JOIN VE_PEDIDO p ON p.CODIGO = i.PEDIDO
        WHERE p.DTPEDIDO BETWEEN '{ano}-{mes:02d}-01' AND '{ano}-{mes:02d}-{ultimo:02d}'
          AND p.STATUS <> 'C'
          AND i.STATUS <> 'C'
          AND i.TPVENDA NOT IN ({TPVENDA_STR})
          AND i.MATERIAL NOT LIKE '8%'
          AND i.FLAGSUB <> 'S'
          AND p.CODIGO NOT IN (
              SELECT p2.CODIGO FROM VE_PEDIDO p2
              JOIN FN_FORNECEDORES f2 ON f2.CODIGO = p2.CLIENTE
              WHERE f2.UF = 'EX')
                    {filtro_cliente_extra}
    """
    r = pd.read_sql(sql, engine).iloc[0]
    return {"pedidos": int(r["Pedidos"]), "itens": int(r["Itens"]),
            "faturamento": float(r["Faturamento"] or 0)}


def get_faturamento_mtd_atual(engine, filtro_cliente_extra: str = "") -> dict:
    ano = HOJE.year
    mes = HOJE.month
    dia = HOJE.day
    ultimo = calendar.monthrange(ano, mes)[1]
    dt_ini = f"{ano}-{mes:02d}-01"
    dt_hoje = f"{ano}-{mes:02d}-{dia:02d}"

    sql = f"""
        SELECT
            COUNT(DISTINCT p.CODIGO) AS Pedidos_MTD,
            COUNT(i.CODIGO)          AS Itens_MTD,
            SUM(i.VLRTOTAL)          AS Faturamento_MTD,
            COUNT(DISTINCT CASE WHEN CAST(p.DTPEDIDO AS date) = '{dt_hoje}' THEN p.CODIGO END) AS Pedidos_Hoje,
            SUM(CASE WHEN CAST(p.DTPEDIDO AS date) = '{dt_hoje}' THEN i.VLRTOTAL ELSE 0 END) AS Faturamento_Hoje
        FROM VE_PEDIDOITENS i
        JOIN VE_PEDIDO p ON p.CODIGO = i.PEDIDO
        WHERE p.DTPEDIDO BETWEEN '{dt_ini}' AND '{dt_hoje}'
          AND p.STATUS <> 'C'
          AND i.STATUS <> 'C'
          AND i.TPVENDA NOT IN ({TPVENDA_STR})
          AND i.MATERIAL NOT LIKE '8%'
          AND i.FLAGSUB <> 'S'
          AND p.CODIGO NOT IN (
              SELECT p2.CODIGO FROM VE_PEDIDO p2
              JOIN FN_FORNECEDORES f2 ON f2.CODIGO = p2.CLIENTE
              WHERE f2.UF = 'EX')
          {filtro_cliente_extra}
    """

    r = pd.read_sql(sql, engine).iloc[0]
    fat_mtd = float(r["Faturamento_MTD"] or 0)
    fat_hoje = float(r["Faturamento_Hoje"] or 0)
    meta_mes = float(META_MENSAL.get(ano, 0) or 0)
    run_rate = (fat_mtd / dia) if dia > 0 else 0.0
    projecao = run_rate * ultimo
    gap_meta = meta_mes - projecao
    pct_meta = (projecao / meta_mes * 100.0) if meta_mes else 0.0

    return {
        "ano": ano,
        "mes": mes,
        "dia": dia,
        "dias_mes": ultimo,
        "pedidos_mtd": int(r["Pedidos_MTD"] or 0),
        "itens_mtd": int(r["Itens_MTD"] or 0),
        "faturamento_mtd": fat_mtd,
        "pedidos_hoje": int(r["Pedidos_Hoje"] or 0),
        "faturamento_hoje": fat_hoje,
        "run_rate_dia": run_rate,
        "projecao_mes": projecao,
        "meta_mes": meta_mes,
        "gap_meta": gap_meta,
        "pct_meta_proj": pct_meta,
    }


def get_crescimento_real_yoy(engine, filtro_cliente_extra: str = "") -> dict:
    mes, ano = REF_MES, REF_ANO
    ano_ant = ano - 1
    sql = f"""
        SELECT
            SUM(CASE WHEN YEAR(p.DTPEDIDO) = {ano_ant} AND MONTH(p.DTPEDIDO) = {mes}
                THEN i.VLRTOTAL ELSE 0 END) AS Fat_Anterior,
            SUM(CASE WHEN YEAR(p.DTPEDIDO) = {ano} AND MONTH(p.DTPEDIDO) = {mes}
                THEN i.VLRTOTAL ELSE 0 END) AS Fat_Atual
        FROM VE_PEDIDOITENS i
        JOIN VE_PEDIDO p ON p.CODIGO = i.PEDIDO
        WHERE YEAR(p.DTPEDIDO) IN ({ano_ant}, {ano})
          AND MONTH(p.DTPEDIDO) = {mes}
          AND p.STATUS <> 'C'
          AND i.STATUS <> 'C'
          AND i.TPVENDA NOT IN ({TPVENDA_STR})
          AND i.MATERIAL NOT LIKE '8%'
          AND i.FLAGSUB <> 'S'
          AND p.CODIGO NOT IN (
              SELECT p2.CODIGO FROM VE_PEDIDO p2
              JOIN FN_FORNECEDORES f2 ON f2.CODIGO = p2.CLIENTE
              WHERE f2.UF = 'EX')
          {filtro_cliente_extra}
    """
    r = pd.read_sql(sql, engine).iloc[0]
    anterior = float(r["Fat_Anterior"] or 0)
    atual = float(r["Fat_Atual"] or 0)
    variacao_abs = atual - anterior
    variacao_pct = (variacao_abs / anterior * 100.0) if anterior else 0.0
    return {
        "anterior": anterior,
        "atual": atual,
        "var_abs": variacao_abs,
        "var_pct": variacao_pct,
    }


def get_yoy_vendedores_mes(engine, filtro_cliente_extra: str = "") -> pd.DataFrame:
    mes, ano = REF_MES, REF_ANO
    ano_ant = ano - 1
    sql = f"""
        SELECT
            p.CODIGO AS Pedido,
            p.DTPEDIDO AS Dt_Pedido,
            YEAR(p.DTPEDIDO) AS Ano,
            f.UF AS Cliente_UF,
            f.CIDADE AS Cliente_Cidade,
            SUM(i.VLRTOTAL) AS Faturamento
        FROM VE_PEDIDOITENS i
        JOIN VE_PEDIDO p ON p.CODIGO = i.PEDIDO
        JOIN FN_FORNECEDORES f ON f.CODIGO = p.CLIENTE
        WHERE YEAR(p.DTPEDIDO) IN ({ano_ant}, {ano})
          AND MONTH(p.DTPEDIDO) = {mes}
          AND p.STATUS <> 'C'
          AND i.STATUS <> 'C'
          AND i.TPVENDA NOT IN ({TPVENDA_STR})
          AND i.MATERIAL NOT LIKE '8%'
          AND i.FLAGSUB <> 'S'
          AND p.CODIGO NOT IN (
              SELECT p2.CODIGO FROM VE_PEDIDO p2
              JOIN FN_FORNECEDORES f2 ON f2.CODIGO = p2.CLIENTE
              WHERE f2.UF = 'EX')
          {filtro_cliente_extra}
        GROUP BY p.CODIGO, p.DTPEDIDO, YEAR(p.DTPEDIDO), f.UF, f.CIDADE
    """
    df = pd.read_sql(sql, engine)
    if df.empty:
        return pd.DataFrame(columns=["Vendedor", "Tipo", "Fat_Anterior", "Fat_Atual", "Var_Abs", "Var_Pct"])

    df["Vendedor"] = df.apply(
        lambda r: mapear_representante_externo(r.get("Cliente_UF"), r.get("Cliente_Cidade"), r.get("Dt_Pedido")),
        axis=1,
    )
    df["Tipo"] = "E"

    df = (
        df.groupby(["Vendedor", "Tipo", "Ano"], dropna=False)
          .agg(Faturamento=("Faturamento", "sum"))
          .reset_index()
    )

    pivot = (df.pivot_table(index=["Vendedor", "Tipo"], columns="Ano", values="Faturamento",
                            aggfunc="sum", fill_value=0)
               .reset_index())

    pivot["Fat_Anterior"] = pivot[ano_ant] if ano_ant in pivot.columns else 0.0
    pivot["Fat_Atual"] = pivot[ano] if ano in pivot.columns else 0.0
    pivot["Var_Abs"] = pivot["Fat_Atual"] - pivot["Fat_Anterior"]
    pivot["Var_Pct"] = pivot.apply(
        lambda r: (r["Var_Abs"] / r["Fat_Anterior"] * 100.0) if r["Fat_Anterior"] > 0 else None,
        axis=1,
    )

    cols = ["Vendedor", "Tipo", "Fat_Anterior", "Fat_Atual", "Var_Abs", "Var_Pct"]
    return pivot[cols].sort_values("Fat_Atual", ascending=False)


def get_cotacoes_mes(engine, filtro_orc_extra: str = "") -> pd.DataFrame:
    mes, ano = REF_MES, REF_ANO
    ultimo = calendar.monthrange(ano, mes)[1]
    sql = f"""
        SELECT
            o.CODIGO AS Cod_Orcamento,
            o.NUMERO AS Num_Orcamento,
            o.DTCADASTRO AS Dt_Cotacao,
            o.DTVALIDADE AS Dt_Validade,
            o.STATUS AS Status_Orc,
            ISNULL(o.VLRORCADO, 0) AS Vlr_Orcado,
            ISNULL(o.VLREFETIVO, 0) AS Vlr_Convertido,
            f.RAZAO AS Cliente,
            f.UF AS Cliente_UF,
            f.CIDADE AS Cliente_Cidade,
            v.RAZAO AS Vendedor
        FROM VE_ORCAMENTOS o
        LEFT JOIN FN_FORNECEDORES f ON f.CODIGO = o.CODCLIENTE
        LEFT JOIN FN_VENDEDORES v ON v.CODIGO = o.VENDEDOR
        WHERE o.DTCADASTRO BETWEEN '{ano}-{mes:02d}-01' AND '{ano}-{mes:02d}-{ultimo:02d}'
                    AND o.FILIAL IN ({FILIAIS_ORC_STR})
          {filtro_orc_extra}
        ORDER BY o.DTCADASTRO DESC
    """
    df = pd.read_sql(sql, engine)
    if df.empty:
        return df

    df["Vendedor_Territorial"] = df.apply(
        lambda r: mapear_representante_externo(r.get("Cliente_UF"), r.get("Cliente_Cidade"), r.get("Dt_Cotacao")),
        axis=1,
    )
    return df


def fig_funil_cotacoes(df_cot: pd.DataFrame) -> go.Figure:
    if df_cot.empty:
        fig = go.Figure()
        fig.update_layout(title="Funil de Cotações", height=280,
                          annotations=[dict(text="Sem dados", x=0.5, y=0.5, showarrow=False, font_size=14)])
        return fig

    status_norm = df_cot["Status_Orc"].map(classificar_status_orc)
    total = int(len(df_cot))
    encerradas = int((status_norm == "ENCERRADA").sum())
    perdidas = int((status_norm == "PERDIDA").sum())
    abertas = int((status_norm == "ABERTA").sum())

    labels = ["Criadas", "Em aberto", "Encerradas (pedido)", "Perdidas"]
    values = [total, abertas, encerradas, perdidas]
    cores = [COR_AZUL_CLARO, COR_AMARELO, COR_VERDE, COR_VERMELHO]

    fig = go.Figure(go.Funnel(
        y=labels,
        x=values,
        marker=dict(color=cores),
        textinfo="value+percent initial",
    ))
    fig.update_layout(
        title=f"Funil de Cotações — {REF_MES:02d}/{REF_ANO}",
        plot_bgcolor="white", paper_bgcolor="white",
        height=320,
        margin=dict(t=50, b=20, l=40, r=20),
    )
    return fig


def get_carteira(engine, filtro_cliente_extra: str = "") -> pd.DataFrame:
    sql = f"""
        SELECT
            i.CODIGO, p.CODIGO AS Pedido, p.NUMINTERNO,
            f.RAZAO AS Cliente, v.RAZAO AS Vendedor,
            i.MATERIAL, i.DESCRICAO,
            i.STATUS AS Status_Item,
            CASE i.STATUS
                WHEN 'L' THEN i.VLRTOTAL
                WHEN 'V' THEN i.VLRTOTAL
                WHEN 'A' THEN i.VLRTOTAL
                ELSE (ISNULL(i.QTDE,0)-ISNULL(i.QTDEFAT,0))*ISNULL(i.VLRUNITARIO,0)
            END AS Vlr_Saldo,
            i.DTPRAZO,
            DATEDIFF(DAY, GETDATE(), i.DTPRAZO) AS Dias_p_Prazo,
            i.DTALTERAFAT,
            p.DTPEDIDO AS DtPedido
        FROM VE_PEDIDOITENS i
        JOIN VE_PEDIDO p       ON p.CODIGO = i.PEDIDO
        JOIN FN_FORNECEDORES f ON f.CODIGO = p.CLIENTE
        JOIN FN_VENDEDORES v   ON v.CODIGO = p.VENDEDOR
        WHERE 1=1
        {FILTROS_PECAS}
        {filtro_cliente_extra}
        ORDER BY i.DTPRAZO ASC
    """
    df = pd.read_sql(sql, engine)

    def sem(d):
        if pd.isna(d): return "Sem prazo"
        dias = int(d)
        if dias < 0:  return "Atrasado"
        if dias <= 7: return "Urgente"
        return "No prazo"

    df["Semaforo"] = df["Dias_p_Prazo"].apply(sem)

    # Dias aguardando NF para status L
    # Regra: considerar apenas desde liberação para faturamento (DTALTERAFAT)
    def calc_dias_nf(r):
        if r["Status_Item"] != "L":
            return None
        ref = r["DTALTERAFAT"]
        if pd.isna(ref):
            return None
        return (HOJE - pd.Timestamp(ref).date()).days

    df["Dias_Aguard_NF"] = df.apply(calc_dias_nf, axis=1)
    return df


def get_sla_operacao_hoje(engine, filtro_cliente_extra: str = "") -> dict:
    hoje_sql = HOJE.strftime("%Y-%m-%d")

    sql_entradas = f"""
        SELECT
            COUNT(i.CODIGO) AS Itens_Entrada,
            SUM(i.VLRTOTAL) AS Vlr_Entrada
        FROM VE_PEDIDOITENS i
        JOIN VE_PEDIDO p       ON p.CODIGO = i.PEDIDO
        JOIN FN_FORNECEDORES f ON f.CODIGO = p.CLIENTE
        WHERE i.STATUS = 'L'
                    AND i.DTALTERAFAT IS NOT NULL
                    AND CAST(i.DTALTERAFAT AS date) = '{hoje_sql}'
          AND p.STATUS <> 'C'
          AND i.STATUS NOT IN ('C','F')
          AND i.TPVENDA NOT IN ({TPVENDA_STR})
          AND i.MATERIAL NOT LIKE '8%'
          AND i.FLAGSUB <> 'S'
          AND p.CODIGO NOT IN (
              SELECT p2.CODIGO FROM VE_PEDIDO p2
              JOIN FN_FORNECEDORES f2 ON f2.CODIGO = p2.CLIENTE
              WHERE f2.UF = 'EX'
          )
          {filtro_cliente_extra}
    """

    sql_resolvidas = f"""
        SELECT
            COUNT(DISTINCT i.CODIGO) AS Itens_Resolvidas,
            SUM(i.VLRTOTAL) AS Vlr_Resolvidas
        FROM VE_PEDIDOITENS i
        JOIN VE_PEDIDO p         ON p.CODIGO = i.PEDIDO
        JOIN FN_FORNECEDORES f   ON f.CODIGO = p.CLIENTE
        JOIN FN_NFEITENS ni      ON ni.CODIGO = i.NFEITEM
        JOIN FN_NFS nf           ON nf.CODIGO = ni.NFE
        WHERE CAST(nf.DTEMISSAO AS date) = '{hoje_sql}'
          AND ISNULL(nf.STATUSNF, '') <> 'C'
          AND p.STATUS <> 'C'
          AND i.STATUS <> 'C'
          AND i.TPVENDA NOT IN ({TPVENDA_STR})
          AND i.MATERIAL NOT LIKE '8%'
          AND i.FLAGSUB <> 'S'
          AND p.CODIGO NOT IN (
              SELECT p2.CODIGO FROM VE_PEDIDO p2
              JOIN FN_FORNECEDORES f2 ON f2.CODIGO = p2.CLIENTE
              WHERE f2.UF = 'EX'
          )
          {filtro_cliente_extra}
    """

    sql_backlog = f"""
        SELECT
            COUNT(i.CODIGO) AS Itens_Backlog,
            SUM(i.VLRTOTAL) AS Vlr_Backlog
        FROM VE_PEDIDOITENS i
        JOIN VE_PEDIDO p       ON p.CODIGO = i.PEDIDO
        JOIN FN_FORNECEDORES f ON f.CODIGO = p.CLIENTE
        WHERE i.STATUS = 'L'
          AND p.STATUS <> 'C'
          AND i.STATUS NOT IN ('C','F')
          AND i.TPVENDA NOT IN ({TPVENDA_STR})
          AND i.MATERIAL NOT LIKE '8%'
          AND i.FLAGSUB <> 'S'
          AND p.CODIGO NOT IN (
              SELECT p2.CODIGO FROM VE_PEDIDO p2
              JOIN FN_FORNECEDORES f2 ON f2.CODIGO = p2.CLIENTE
              WHERE f2.UF = 'EX'
          )
          {filtro_cliente_extra}
    """

    ent = pd.read_sql(sql_entradas, engine).iloc[0]
    res = pd.read_sql(sql_resolvidas, engine).iloc[0]
    bkl = pd.read_sql(sql_backlog, engine).iloc[0]

    entradas = int(ent["Itens_Entrada"] or 0)
    resolvidas = int(res["Itens_Resolvidas"] or 0)
    backlog = int(bkl["Itens_Backlog"] or 0)

    return {
        "dt_ref": HOJE.strftime("%d/%m/%Y"),
        "entradas_itens": entradas,
        "entradas_valor": float(ent["Vlr_Entrada"] or 0),
        "resolvidas_itens": resolvidas,
        "resolvidas_valor": float(res["Vlr_Resolvidas"] or 0),
        "backlog_itens": backlog,
        "backlog_valor": float(bkl["Vlr_Backlog"] or 0),
        "saldo_itens": entradas - resolvidas,
        "taxa_saida": (resolvidas / entradas * 100.0) if entradas else 0.0,
    }


def construir_prioridades_operacionais(df_cart: pd.DataFrame) -> pd.DataFrame:
    if df_cart.empty:
        return pd.DataFrame()

    df = df_cart.copy()
    df["Dias_Aguard_NF"] = pd.to_numeric(df["Dias_Aguard_NF"], errors="coerce")
    df["Dias_p_Prazo"] = pd.to_numeric(df["Dias_p_Prazo"], errors="coerce")
    df["Vlr_Saldo"] = pd.to_numeric(df["Vlr_Saldo"], errors="coerce").fillna(0.0)

    df_acao = df[
        (df["Status_Item"] == "L")
        | (df["Semaforo"].isin(["Atrasado", "Urgente"]))
    ].copy()

    if df_acao.empty:
        return df_acao

    def definir_prioridade(r):
        if r["Status_Item"] == "L" and pd.notna(r["Dias_Aguard_NF"]) and r["Dias_Aguard_NF"] >= 10:
            return "NF CRITICA", 1, "Cobrar emissão imediata"
        if r["Semaforo"] == "Atrasado":
            return "PRAZO ATRASADO", 2, "Replanejar entrega e alinhar cliente"
        if r["Status_Item"] == "L":
            return "NF ATENCAO", 3, "Priorizar faturamento"
        return "PRAZO URGENTE", 4, "Confirmar expedição da semana"

    tmp = df_acao.apply(definir_prioridade, axis=1, result_type="expand")
    df_acao["Prioridade"] = tmp[0]
    df_acao["Ordem"] = tmp[1]
    df_acao["Acao_Sugerida"] = tmp[2]

    return df_acao.sort_values(["Ordem", "Vlr_Saldo"], ascending=[True, False]).head(20)


def get_ranking_vendedores(engine, filtro_cliente_extra: str = "") -> pd.DataFrame:
    ano = REF_ANO
    sql = f"""
        SELECT
            p.CODIGO AS Pedido,
            p.DTPEDIDO AS Dt_Pedido,
            YEAR(p.DTPEDIDO) AS Ano,
            f.UF AS Cliente_UF,
            f.CIDADE AS Cliente_Cidade,
            SUM(i.VLRTOTAL) AS Faturamento,
            ROUND(AVG(CASE WHEN i.VLRCUSTO>0 AND i.VLRUNITARIO>0
                THEN (i.VLRUNITARIO-i.VLRCUSTO)/i.VLRUNITARIO*100 ELSE NULL END),1) AS Margem_pct
        FROM VE_PEDIDOITENS i
        JOIN VE_PEDIDO p      ON p.CODIGO = i.PEDIDO
        JOIN FN_FORNECEDORES f ON f.CODIGO = p.CLIENTE
        WHERE YEAR(p.DTPEDIDO) IN ({ano-1},{ano})
          AND p.STATUS <> 'C' AND i.STATUS <> 'C'
          AND i.FLAGSUB <> 'S' AND i.MATERIAL NOT LIKE '8%'
          AND i.TPVENDA NOT IN ({TPVENDA_STR})
          AND p.CODIGO NOT IN (SELECT p2.CODIGO FROM VE_PEDIDO p2
              JOIN FN_FORNECEDORES f2 ON f2.CODIGO=p2.CLIENTE WHERE f2.UF='EX')
          {filtro_cliente_extra}
        GROUP BY p.CODIGO, p.DTPEDIDO, YEAR(p.DTPEDIDO), f.UF, f.CIDADE
    """
    df = pd.read_sql(sql, engine)
    if df.empty:
        return df

    df["Vendedor"] = df.apply(
        lambda r: mapear_representante_externo(r.get("Cliente_UF"), r.get("Cliente_Cidade"), r.get("Dt_Pedido")),
        axis=1,
    )
    df["Tipo"] = "E"

    return (
        df.groupby(["Vendedor", "Tipo", "Ano"], dropna=False)
          .agg(
              Faturamento=("Faturamento", "sum"),
              Pedidos=("Pedido", "nunique"),
              Margem_pct=("Margem_pct", "mean"),
          )
          .reset_index()
          .sort_values("Faturamento", ascending=False)
    )


def get_abc_clientes(engine, filtro_cliente_extra: str = "") -> pd.DataFrame:
    ano = REF_ANO
    sql = f"""
        SELECT
            f.RAZAO AS Cliente,
            SUM(i.VLRTOTAL) AS Faturamento
        FROM VE_PEDIDOITENS i
        JOIN VE_PEDIDO p      ON p.CODIGO = i.PEDIDO
        JOIN FN_FORNECEDORES f ON f.CODIGO = p.CLIENTE
        WHERE YEAR(p.DTPEDIDO) = {ano}
          AND p.STATUS <> 'C' AND i.STATUS <> 'C'
          AND i.FLAGSUB <> 'S' AND i.MATERIAL NOT LIKE '8%'
          AND i.TPVENDA NOT IN ({TPVENDA_STR})
          AND p.CODIGO NOT IN (SELECT p2.CODIGO FROM VE_PEDIDO p2
              JOIN FN_FORNECEDORES f2 ON f2.CODIGO=p2.CLIENTE WHERE f2.UF='EX')
          {filtro_cliente_extra}
        GROUP BY f.RAZAO
        ORDER BY Faturamento DESC
    """
    df = pd.read_sql(sql, engine)
    total = df["Faturamento"].sum()
    acum  = 0.0
    curva = []
    for v in df["Faturamento"]:
        acum += v
        pct   = acum / total * 100
        curva.append("A" if pct <= 80 else ("B" if pct <= 95 else "C"))
    df["Curva"] = curva
    return df


def get_margem_alertas(engine, filtro_cliente_extra: str = "") -> pd.DataFrame:
    mes, ano = REF_MES, REF_ANO
    ultimo = calendar.monthrange(ano, mes)[1]
    sql = f"""
        SELECT TOP 20
            f.RAZAO AS Cliente,
            v.RAZAO AS Vendedor,
            i.MATERIAL, i.DESCRICAO,
            i.VLRUNITARIO AS Vlr_Venda,
            i.VLRCUSTO    AS Vlr_Custo,
            ROUND((i.VLRUNITARIO-i.VLRCUSTO)/i.VLRUNITARIO*100,1) AS Margem_pct,
            i.VLRTOTAL    AS Vlr_Total
        FROM VE_PEDIDOITENS i
        JOIN VE_PEDIDO p       ON p.CODIGO = i.PEDIDO
        JOIN FN_FORNECEDORES f ON f.CODIGO = p.CLIENTE
        JOIN FN_VENDEDORES v   ON v.CODIGO = p.VENDEDOR
        WHERE p.DTPEDIDO BETWEEN '{ano}-{mes:02d}-01' AND '{ano}-{mes:02d}-{ultimo:02d}'
          AND i.VLRCUSTO > 0 AND i.VLRUNITARIO > 0
          AND (i.VLRUNITARIO-i.VLRCUSTO)/i.VLRUNITARIO*100 < 20
          AND p.STATUS <> 'C' AND i.STATUS <> 'C'
          AND i.FLAGSUB <> 'S' AND i.MATERIAL NOT LIKE '8%'
          AND i.TPVENDA NOT IN ({TPVENDA_STR})
          AND p.CODIGO NOT IN (SELECT p2.CODIGO FROM VE_PEDIDO p2
              JOIN FN_FORNECEDORES f2 ON f2.CODIGO=p2.CLIENTE WHERE f2.UF='EX')
          {filtro_cliente_extra}
        ORDER BY Margem_pct ASC
    """
    return pd.read_sql(sql, engine)


# ── Formatação ────────────────────────────────────────────────────────────────
def fmt_brl(v: float) -> str:
    return f"R$ {v:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_pct(v: float) -> str:
    return f"{v:.1f}%"


# ── Construção dos gráficos ───────────────────────────────────────────────────

def fig_evolucao(df_ev: pd.DataFrame) -> go.Figure:
    ano_atual  = REF_ANO
    ano_ant    = ano_atual - 1
    d_atual    = df_ev[df_ev["Ano"] == ano_atual].copy()
    d_anterior = df_ev[df_ev["Ano"] == ano_ant].copy()

    meses_nome = ["Jan","Fev","Mar","Abr","Mai","Jun",
                  "Jul","Ago","Set","Out","Nov","Dez"]
    d_atual["MesNome"]    = d_atual["Mes"].apply(lambda m: meses_nome[m-1])
    d_anterior["MesNome"] = d_anterior["Mes"].apply(lambda m: meses_nome[m-1])

    # Eixo X fixo com todos os 12 meses para as linhas de meta ficarem cheias
    todos_meses = meses_nome

    meta_ant    = META_MENSAL.get(ano_ant)
    meta_atual  = META_MENSAL.get(ano_atual)

    fig = go.Figure()

    # Linha de meta do ano anterior
    if meta_ant:
        fig.add_trace(go.Scatter(
            x=todos_meses, y=[meta_ant] * 12,
            mode="lines", name=f"Meta {ano_ant}",
            line=dict(color=COR_CINZA, width=1.5, dash="dot"),
            hovertemplate=f"Meta {ano_ant}: {fmt_brl(meta_ant)}<extra></extra>",
        ))

    # Linha de meta do ano atual
    if meta_atual:
        fig.add_trace(go.Scatter(
            x=todos_meses, y=[meta_atual] * 12,
            mode="lines", name=f"Meta {ano_atual}",
            line=dict(color=COR_LARANJA, width=2, dash="dash"),
            hovertemplate=f"Meta {ano_atual}: {fmt_brl(meta_atual)}<extra></extra>",
        ))

    fig.add_trace(go.Bar(
        x=d_anterior["MesNome"], y=d_anterior["Faturamento"],
        name=str(ano_ant), marker_color=COR_CINZA,
        opacity=0.7,
        text=[fmt_brl(v) for v in d_anterior["Faturamento"]],
        textposition="outside", textfont_size=10,
    ))
    fig.add_trace(go.Bar(
        x=d_atual["MesNome"], y=d_atual["Faturamento"],
        name=str(ano_atual), marker_color=COR_AZUL_CLARO,
        text=[fmt_brl(v) for v in d_atual["Faturamento"]],
        textposition="outside", textfont_size=10,
    ))

    # Determina teto do eixo Y
    max_val = max(
        d_anterior["Faturamento"].max() if not d_anterior.empty else 0,
        d_atual["Faturamento"].max() if not d_atual.empty else 0,
        meta_ant or 0,
        meta_atual or 0,
    )

    fig.update_layout(
        title=f"Evolução de Vendas Mensal — Peças ({ano_ant} vs {ano_atual})",
        barmode="group",
        xaxis_title="", yaxis_title="R$",
        legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="left", x=0),
        plot_bgcolor="white", paper_bgcolor="white",
        height=420,
        margin=dict(t=80, b=30, l=80, r=20),
        yaxis=dict(tickformat=",.0f", range=[0, max_val * 1.18]),
    )
    return fig


def fig_semaforo_carteira(df_cart: pd.DataFrame) -> go.Figure:
    contagem = df_cart["Semaforo"].value_counts()
    ordem    = ["Atrasado", "Urgente", "No prazo", "Sem prazo"]
    cores    = [COR_VERMELHO, COR_AMARELO, COR_VERDE, COR_CINZA]
    labels, values, pull_vals = [], [], []
    for s, c in zip(ordem, cores):
        if s in contagem.index:
            labels.append(s)
            values.append(int(contagem[s]))
            pull_vals.append(0.05 if s == "Atrasado" else 0)

    fig = go.Figure(go.Pie(
        labels=labels, values=values,
        hole=0.55,
        pull=pull_vals,
        marker_colors=[COR_VERMELHO, COR_AMARELO, COR_VERDE, COR_CINZA][:len(labels)],
        textinfo="label+value",
        textfont_size=12,
    ))
    total = df_cart["Vlr_Saldo"].sum()
    fig.update_layout(
        title="Carteira em Aberto",
        annotations=[dict(
            text=f"<b>{fmt_brl(total)}</b>",
            x=0.5, y=0.5, font_size=13, showarrow=False
        )],
        plot_bgcolor="white", paper_bgcolor="white",
        height=320,
        margin=dict(t=50, b=10, l=10, r=10),
        showlegend=False,
    )
    return fig


def _ranking_base(df_vend: pd.DataFrame, tipo: str, cor: str, titulo: str) -> go.Figure:
    """Gera ranking de vendedores filtrado por TIPO ('E'=externo, 'I'=interno)."""
    ano_atual = REF_ANO
    ano_ant   = ano_atual - 1

    mask_tipo = df_vend["Tipo"] == tipo
    d_atual   = (df_vend[mask_tipo & (df_vend["Ano"] == ano_atual)]
                 .set_index("Vendedor")["Faturamento"].sort_values())
    d_ant     = (df_vend[mask_tipo & (df_vend["Ano"] == ano_ant)]
                 .set_index("Vendedor")["Faturamento"])

    if d_atual.empty:
        fig = go.Figure()
        fig.update_layout(title=titulo, height=220,
                          annotations=[dict(text="Sem dados", x=0.5, y=0.5,
                                           showarrow=False, font_size=14)])
        return fig

    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=d_atual.index, x=d_ant.reindex(d_atual.index).fillna(0),
        name=str(ano_ant), orientation="h",
        marker_color=COR_CINZA, opacity=0.65,
    ))
    fig.add_trace(go.Bar(
        y=d_atual.index, x=d_atual.values,
        name=str(ano_atual), orientation="h",
        marker_color=cor,
        text=[fmt_brl(v) for v in d_atual.values],
        textposition="outside", textfont_size=10,
    ))
    fig.update_layout(
        title=titulo,
        barmode="overlay",
        xaxis_title="R$", yaxis_title="",
        xaxis=dict(tickformat=",.0f"),
        plot_bgcolor="white", paper_bgcolor="white",
        height=max(260, 52 * len(d_atual) + 80),
        margin=dict(t=50, b=30, l=170, r=90),
        legend=dict(orientation="h", y=1.08),
    )
    return fig


def fig_ranking_externos(df_vend: pd.DataFrame) -> go.Figure:
    return _ranking_base(df_vend, "E", COR_AZUL,
                         f"Ranking Repr. Externos — {REF_ANO} vs {REF_ANO-1}")


def fig_ranking_internos(df_vend: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    fig.update_layout(
        title="Ranking Internos (desconsiderado)",
        height=220,
        annotations=[dict(
            text="Internos fora do ranking comercial nesta fase",
            x=0.5, y=0.5, showarrow=False, font_size=13
        )],
        plot_bgcolor="white", paper_bgcolor="white",
    )
    return fig


def fig_abc_treemap(df_abc: pd.DataFrame) -> go.Figure:
    cores_curva = {"A": COR_VERDE, "B": COR_AMARELO, "C": COR_LARANJA}
    df_plot = df_abc.head(30).copy()
    df_plot["Cor"] = df_plot["Curva"].map(cores_curva)
    df_plot["Label"] = df_plot.apply(
        lambda r: f"{r['Cliente'][:22]}<br>{fmt_brl(r['Faturamento'])}", axis=1
    )

    fig = px.treemap(
        df_plot,
        path=["Curva", "Cliente"],
        values="Faturamento",
        color="Curva",
        color_discrete_map=cores_curva,
        title=f"Curva ABC Clientes — {REF_ANO} (Top 30)",
    )
    fig.update_traces(
        texttemplate="%{label}<br>%{value:,.0f}",
        textfont_size=11,
    )
    fig.update_layout(
        plot_bgcolor="white", paper_bgcolor="white",
        height=380,
        margin=dict(t=50, b=10, l=10, r=10),
    )
    return fig


def fig_performance_externos(df_vend: pd.DataFrame) -> go.Figure:
    """
    Gráfico duplo eixo para representantes externos no ano atual:
    - Barras: Vendas de peças (eixo esquerdo)
      - Linha + pontos: Margem % média (eixo direito)
    Margem = (Preço Venda − Custo Cadastrado) / Preço Venda × 100
    Calculada sobre itens que têm VLRCUSTO > 0 no período.
    """
    ano = REF_ANO
    d = (df_vend[(df_vend["Tipo"] == "E") & (df_vend["Ano"] == ano)]
         .sort_values("Faturamento", ascending=False))

    if d.empty:
        fig = go.Figure()
        fig.update_layout(title="Performance Externos (Vendas de Peças)", height=280)
        return fig

    nomes = d["Vendedor"].tolist()
    fat   = d["Faturamento"].tolist()
    marg  = d["Margem_pct"].tolist()
    cores_b = [COR_VERMELHO if (m or 0) < 20 else
               (COR_AMARELO if (m or 0) < 40 else COR_AZUL_CLARO)
               for m in marg]

    fig = make_subplots(specs=[[{"secondary_y": True}]])

    fig.add_trace(go.Bar(
        x=nomes, y=fat,
        name="Vendas de Peças",
        marker_color=cores_b,
        text=[fmt_brl(v) for v in fat],
        textposition="outside", textfont_size=10,
    ), secondary_y=False)

    fig.add_trace(go.Scatter(
        x=nomes, y=marg,
        name="Margem %",
        mode="lines+markers+text",
        line=dict(color=COR_LARANJA, width=2),
        marker=dict(size=8, color=COR_LARANJA),
        text=[f"{v:.0f}%" if v is not None else "—" for v in marg],
        textposition="top center", textfont=dict(color=COR_LARANJA, size=10),
    ), secondary_y=True)

    max_fat = max(fat) if fat else 1
    fig.update_yaxes(title_text="R$", secondary_y=False,
                     tickformat=",.0f", range=[0, max_fat * 1.2])
    fig.update_yaxes(title_text="Margem %", secondary_y=True,
                     range=[0, 100], ticksuffix="%")

    fig.update_layout(
        title=f"Performance Externos {ano} — Vendas de Peças e Margem Média",
        plot_bgcolor="white", paper_bgcolor="white",
        height=380,
        margin=dict(t=60, b=40, l=80, r=80),
        legend=dict(orientation="h", y=1.1),
        xaxis=dict(tickangle=-20),
    )
    return fig


# ── HTML completo ─────────────────────────────────────────────────────────────

def gerar_html(dados: dict) -> str:
    # Serializa gráficos para HTML (sem JS embutido — usa CDN)
    config = {"displayModeBar": False, "responsive": True}

    html_ev    = dados["fig_ev"].to_html(full_html=False, include_plotlyjs=False, config=config, div_id="ev")
    html_sem   = dados["fig_sem"].to_html(full_html=False, include_plotlyjs=False, config=config, div_id="sem")
    html_ext   = dados["fig_ext"].to_html(full_html=False, include_plotlyjs=False, config=config, div_id="ext")
    html_abc   = dados["fig_abc"].to_html(full_html=False, include_plotlyjs=False, config=config, div_id="abc")
    html_perf  = dados["fig_perf"].to_html(full_html=False, include_plotlyjs=False, config=config, div_id="perf")
    html_funil = dados["fig_cot_funil"].to_html(full_html=False, include_plotlyjs=False, config=config, div_id="cot_funil")

    fat_mes  = dados["fat_mes"]
    cart     = dados["df_cart"]
    yoy_mes  = dados["yoy_mes"]
    df_yoy_vendedores = dados["df_vend_yoy"]
    df_nf    = cart[cart["Status_Item"] == "L"].sort_values("Dias_Aguard_NF", ascending=False).head(10)
    df_alert = dados["df_alertas"]
    df_cot = dados["df_cotacoes"]
    sla = dados["sla_hoje"]
    mtd = dados["mtd_atual"]

    status_norm = df_cot["Status_Orc"].map(classificar_status_orc) if not df_cot.empty else pd.Series(dtype=str)
    total_cot = int(len(df_cot))
    cot_enc = int((status_norm == "ENCERRADA").sum()) if not df_cot.empty else 0
    cot_perd = int((status_norm == "PERDIDA").sum()) if not df_cot.empty else 0
    cot_aberto = int((status_norm == "ABERTA").sum()) if not df_cot.empty else 0
    cot_vlr = fmt_brl(float(df_cot["Vlr_Orcado"].sum())) if not df_cot.empty else fmt_brl(0)
    cot_conv_qtd = (cot_enc / total_cot * 100.0) if total_cot else 0.0
    cot_finalizadas = cot_enc + cot_perd
    cot_conv_qtd_final = (cot_enc / cot_finalizadas * 100.0) if cot_finalizadas else 0.0
    vlr_orcado = float(df_cot["Vlr_Orcado"].sum()) if not df_cot.empty else 0.0
    vlr_convertido = float(df_cot.loc[status_norm == "ENCERRADA", "Vlr_Orcado"].sum()) if not df_cot.empty else 0.0
    vlr_finalizadas = float(df_cot.loc[status_norm.isin(["ENCERRADA", "PERDIDA"]), "Vlr_Orcado"].sum()) if not df_cot.empty else 0.0
    cot_conv_vlr = (vlr_convertido / vlr_orcado * 100.0) if vlr_orcado else 0.0
    cot_conv_vlr_final = (vlr_convertido / vlr_finalizadas * 100.0) if vlr_finalizadas else 0.0

    mtd_meta_class = "verde" if mtd["gap_meta"] <= 0 else "vermelho"
    mtd_gap_sinal = "+" if mtd["gap_meta"] >= 0 else ""

    df_prioridades = construir_prioridades_operacionais(cart)

    kpi_fat  = fmt_brl(fat_mes["faturamento"])
    kpi_cart = fmt_brl(cart["Vlr_Saldo"].sum())
    n_atr    = int((cart["Semaforo"] == "Atrasado").sum())
    n_urg    = int((cart["Semaforo"] == "Urgente").sum())
    n_nf     = int((cart["Status_Item"] == "L").sum())
    vlr_nf   = fmt_brl(cart[cart["Status_Item"] == "L"]["Vlr_Saldo"].sum())
    yoy_pct = float(yoy_mes["var_pct"])
    yoy_class = "verde" if yoy_pct >= 0 else "vermelho"
    yoy_signal = "+" if yoy_pct >= 0 else ""
    yoy_label = f"{yoy_signal}{yoy_pct:.1f}%"
    yoy_sub = f"{fmt_brl(yoy_mes['atual'])} vs {fmt_brl(yoy_mes['anterior'])} (mesmo filtro ativo)"
    badge_sem_contrato = "Sem contrato" in dados["modo_filtro"]
    modo_badge = "SEM CONTRATO" if badge_sem_contrato else "BASE"
    modo_badge_class = "modo-sem" if badge_sem_contrato else "modo-base"

    def resumo_yoy_vendedores(df: pd.DataFrame) -> dict:
        if df.empty:
            return {
                "maior_alta": "Sem dados",
                "maior_queda": "Sem dados",
            }

        df_valid = df.copy()
        df_valid["Var_Abs"] = pd.to_numeric(df_valid["Var_Abs"], errors="coerce").fillna(0)

        alta = df_valid.sort_values("Var_Abs", ascending=False).iloc[0]
        queda = df_valid.sort_values("Var_Abs", ascending=True).iloc[0]

        def _linha(row, titulo: str) -> str:
            nome = str(row["Vendedor"])[:26]
            var_abs = float(row["Var_Abs"] or 0)
            var_pct = row["Var_Pct"]
            pct_txt = "—" if pd.isna(var_pct) else f" ({var_pct:+.1f}%)"
            return f"{titulo}: {nome} · {fmt_brl(var_abs)}{pct_txt}"

        return {
            "maior_alta": _linha(alta, "Maior alta"),
            "maior_queda": _linha(queda, "Maior queda"),
        }

    resumo_vend = resumo_yoy_vendedores(df_yoy_vendedores)
    impacto_yoy_abs = fmt_brl(float(yoy_mes["var_abs"]))

    mes_nome = ["","Janeiro","Fevereiro","Março","Abril","Maio","Junho",
                "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"][REF_MES]

    def tabela_nf(df) -> str:
        if df.empty:
            return "<p style='color:#888;padding:12px'>Nenhum item aguardando NF.</p>"
        rows = ""
        for _, r in df.iterrows():
            dias = int(r["Dias_Aguard_NF"]) if pd.notna(r["Dias_Aguard_NF"]) else None
            if dias is None:
                cor = "#EDEDED"
                dias_txt = "-"
            else:
                cor = "#FFC7CE" if dias >= 15 else ("#FFEB9C" if dias >= 7 else "#C6EFCE")
                dias_txt = f"{dias}d"
            ref_label = "sem data de lib." if pd.isna(r["DTALTERAFAT"]) else "desde lib."
            rows += f"""<tr style="background:{cor}">
                <td>{int(r['Pedido'])}</td>
                <td>{r['Cliente'][:30]}</td>
                <td>{r['Vendedor'].split()[0]}</td>
                <td>{r['MATERIAL']}</td>
                <td>{r['DESCRICAO'][:28]}</td>
                <td style="text-align:right">{fmt_brl(float(r['Vlr_Saldo']))}</td>
                <td style="text-align:center"><b>{dias_txt}</b><br><span style='font-size:0.7em;color:#666'>{ref_label}</span></td>
            </tr>"""
        return f"""<table class="tabela">
            <thead><tr>
                <th>Pedido</th><th>Cliente</th><th>Vendedor</th>
                <th>Material</th><th>Descrição</th><th>Valor</th><th>Dias</th>
            </tr></thead><tbody>{rows}</tbody></table>
            <p style='font-size:0.72rem;color:#999;padding:6px 10px'>
                            📌 Dias calculados exclusivamente a partir da data de liberação para faturamento (DTALTERAFAT).
                            Itens sem essa data ficam sinalizados como "sem data de lib.".
            </p>"""

    def tabela_alertas(df) -> str:
        if df.empty:
            return "<p style='color:#888;padding:12px'>✅ Nenhum alerta crítico no mês.</p>"
        rows = ""
        for _, r in df.iterrows():
            m    = float(r["Margem_pct"])
            cor  = "#FFC7CE" if m < 0 else "#FFB347"
            rows += f"""<tr style="background:{cor}">
                <td>{r['Cliente'][:28]}</td>
                <td>{r['Vendedor'].split()[0]}</td>
                <td>{r['MATERIAL']}</td>
                <td>{r['DESCRICAO'][:28]}</td>
                <td style="text-align:right">{fmt_brl(float(r['Vlr_Venda']))}</td>
                <td style="text-align:right">{fmt_brl(float(r['Vlr_Custo']))}</td>
                <td style="text-align:center"><b>{m:.1f}%</b></td>
            </tr>"""
        return f"""<table class="tabela">
            <thead><tr>
                <th>Cliente</th><th>Vendedor</th><th>Material</th>
                <th>Descrição</th><th>Vlr Venda</th><th>Vlr Custo</th><th>Margem</th>
            </tr></thead><tbody>{rows}</tbody></table>
            <p style='font-size:0.72rem;color:#666;padding:6px 10px'>
              📐 <b>Como é calculada a margem:</b>
              Margem % = (Preço de Venda − VLRCUSTO) ÷ Preço de Venda × 100.
              Representa o percentual de lucro bruto estimado sobre o item vendido.
              Exibidos apenas itens com custo cadastrado (VLRCUSTO &gt; 0).
              Valores negativos = vendido abaixo do custo. Lista sempre mostra todos abaixo de 20%.
            </p>"""

    def tabela_prioridades(df) -> str:
        if df.empty:
            return "<p style='color:#888;padding:12px'>✅ Sem prioridades críticas no momento.</p>"

        rows = ""
        for _, r in df.iterrows():
            prioridade = str(r["Prioridade"])
            if prioridade == "NF CRITICA":
                cor = "#FFC7CE"
            elif prioridade == "PRAZO ATRASADO":
                cor = "#FFE699"
            elif prioridade == "NF ATENCAO":
                cor = "#FCE4D6"
            else:
                cor = "#E2F0D9"

            dias_nf = "-"
            if pd.notna(r.get("Dias_Aguard_NF")):
                dias_nf = f"{int(r['Dias_Aguard_NF'])}d"

            dias_prazo = "-"
            if pd.notna(r.get("Dias_p_Prazo")):
                dias_prazo = f"{int(r['Dias_p_Prazo'])}d"

            rows += f"""<tr style="background:{cor}">
                <td><b>{prioridade}</b></td>
                <td>{r['Vendedor'][:24]}</td>
                <td>{r['Cliente'][:30]}</td>
                <td style="text-align:center">{int(r['Pedido'])}</td>
                <td>{r['MATERIAL']}</td>
                <td style="text-align:right">{fmt_brl(float(r['Vlr_Saldo']))}</td>
                <td style="text-align:center">{dias_prazo}</td>
                <td style="text-align:center">{dias_nf}</td>
                <td>{r['Acao_Sugerida']}</td>
            </tr>"""

        return f"""<table class="tabela">
            <thead><tr>
                <th>Prioridade</th><th>Vendedor</th><th>Cliente</th><th>Pedido</th>
                <th>Material</th><th>Valor</th><th>Dias Prazo</th><th>Dias NF</th><th>Ação sugerida</th>
            </tr></thead><tbody>{rows}</tbody></table>
            <p style='font-size:0.72rem;color:#666;padding:6px 10px'>
              Ordenação por criticidade e maior impacto financeiro para gestão diária.
            </p>"""

    def tabela_cotacoes(df) -> str:
        if df.empty:
            return "<p style='color:#888;padding:12px'>Sem cotações no período de referência.</p>"

        status_map = df["Status_Orc"].map(classificar_status_orc)
        aberto = df[status_map == "ABERTA"].copy()
        if aberto.empty:
            aberto = df.copy()

        aberto = aberto.sort_values(["Dt_Cotacao", "Vlr_Orcado"], ascending=[True, False]).head(15)

        def status_txt(s):
            classe = classificar_status_orc(s)
            if classe == "ENCERRADA":
                return "ENCERRADA"
            if classe == "PERDIDA":
                return "PERDIDO"
            return "EM ABERTO"

        rows = ""
        for _, r in aberto.iterrows():
            st = status_txt(r.get("Status_Orc"))
            cor = "#C6EFCE" if st == "APROVADO" else ("#FFC7CE" if st == "PERDIDO" else "#FFEB9C")
            dt_cot = pd.to_datetime(r.get("Dt_Cotacao"), errors="coerce")
            dt_txt = dt_cot.strftime("%d/%m/%Y") if pd.notna(dt_cot) else "-"
            rows += f"""<tr style="background:{cor}">
                <td>{str(r.get('Num_Orcamento')) if pd.notna(r.get('Num_Orcamento')) else '-'}</td>
                <td>{r['Vendedor_Territorial'][:24] if pd.notna(r.get('Vendedor_Territorial')) else '-'}</td>
                <td>{r['Cliente'][:30] if pd.notna(r.get('Cliente')) else '-'}</td>
                <td>{r['Cliente_UF'] if pd.notna(r.get('Cliente_UF')) else '-'}</td>
                <td style="text-align:center">{dt_txt}</td>
                <td style="text-align:right">{fmt_brl(float(r['Vlr_Orcado'] or 0))}</td>
                <td style="text-align:center"><b>{st}</b></td>
            </tr>"""

        return f"""<table class="tabela">
            <thead><tr>
                <th>Cotação</th><th>Vendedor</th><th>Cliente</th><th>UF</th><th>Data</th><th>Valor Orçado</th><th>Status</th>
            </tr></thead><tbody>{rows}</tbody></table>
            <p style='font-size:0.72rem;color:#666;padding:6px 10px'>
              Exibe cotações em aberto (prioridade) e, na ausência, as mais recentes do mês fechado.
            </p>"""

    def tabela_conversao_uf(df) -> str:
        if df.empty:
            return "<p style='color:#888;padding:12px'>Sem cotações no período de referência.</p>"

        df_aux = df.copy()
        df_aux["Status_Class"] = df_aux["Status_Orc"].map(classificar_status_orc)
        df_aux["UF_Base"] = df_aux["Cliente_UF"].fillna("-").map(lambda x: normalizar_texto(x) or "-")

        g = (
            df_aux.groupby(["UF_Base"], dropna=False)
              .agg(
                  Cotacoes=("Cod_Orcamento", "nunique"),
                  Encerradas=("Status_Class", lambda s: int((s == "ENCERRADA").sum())),
                  Perdidas=("Status_Class", lambda s: int((s == "PERDIDA").sum())),
                  Em_Aberto=("Status_Class", lambda s: int((s == "ABERTA").sum())),
                  Vlr_Orcado=("Vlr_Orcado", "sum"),
              )
              .reset_index()
        )
        g["Tx_Conv_Qtd"] = g.apply(lambda r: (r["Encerradas"] / r["Cotacoes"] * 100.0) if r["Cotacoes"] else 0.0, axis=1)
        g["Finalizadas"] = g["Encerradas"] + g["Perdidas"]
        g["Tx_Conv_Final"] = g.apply(lambda r: (r["Encerradas"] / r["Finalizadas"] * 100.0) if r["Finalizadas"] else 0.0, axis=1)
        g = g.sort_values(["Vlr_Orcado", "Cotacoes"], ascending=[False, False]).head(12)

        rows = ""
        for _, r in g.iterrows():
            cor = "#E2F0D9" if r["Tx_Conv_Qtd"] >= 40 else ("#FFEB9C" if r["Tx_Conv_Qtd"] >= 20 else "#FCE4D6")
            rows += f"""<tr style="background:{cor}">
                <td>{r['UF_Base']}</td>
                <td style="text-align:center">{int(r['Cotacoes'])}</td>
                <td style="text-align:center">{int(r['Encerradas'])}</td>
                <td style="text-align:center">{int(r['Em_Aberto'])}</td>
                <td style="text-align:center">{int(r['Perdidas'])}</td>
                <td style="text-align:right">{fmt_brl(float(r['Vlr_Orcado']))}</td>
                <td style="text-align:center"><b>{r['Tx_Conv_Qtd']:.1f}%</b></td>
                <td style="text-align:center">{int(r['Finalizadas'])}</td>
                <td style="text-align:center"><b>{r['Tx_Conv_Final']:.1f}%</b></td>
            </tr>"""

        return f"""<table class="tabela">
            <thead><tr>
                                <th>UF</th><th>Cotações</th><th>Encerradas</th><th>Abertas</th><th>Perdidas</th><th>Valor Orçado</th><th>Tx Conv. Geral</th><th>Finalizadas</th><th>Tx Conv. Final</th>
            </tr></thead><tbody>{rows}</tbody></table>
            <p style='font-size:0.72rem;color:#666;padding:6px 10px'>
                            Conversão por UF (estado do cliente). Conversão geral = encerradas/criadas. Conversão final = encerradas/(encerradas+perdidas).
            </p>"""

    def card_sla_operacao(sla_dia: dict) -> str:
        cor_saldo = "#C6EFCE" if sla_dia["saldo_itens"] <= 0 else "#FFE699"
        return f"""
            <table class="tabela">
                <thead><tr>
                    <th>Data</th><th>Entradas (L)</th><th>Resolvidas (NF emitida)</th><th>Saldo do Dia</th><th>Backlog Atual</th><th>Taxa de Saída</th>
                </tr></thead>
                <tbody>
                    <tr style="background:{cor_saldo}">
                        <td style="text-align:center"><b>{sla_dia['dt_ref']}</b></td>
                        <td style="text-align:center"><b>{sla_dia['entradas_itens']}</b><br><span style='font-size:0.72em;color:#666'>{fmt_brl(sla_dia['entradas_valor'])}</span></td>
                        <td style="text-align:center"><b>{sla_dia['resolvidas_itens']}</b><br><span style='font-size:0.72em;color:#666'>{fmt_brl(sla_dia['resolvidas_valor'])}</span></td>
                        <td style="text-align:center"><b>{sla_dia['saldo_itens']:+d}</b></td>
                        <td style="text-align:center"><b>{sla_dia['backlog_itens']}</b><br><span style='font-size:0.72em;color:#666'>{fmt_brl(sla_dia['backlog_valor'])}</span></td>
                        <td style="text-align:center"><b>{sla_dia['taxa_saida']:.1f}%</b></td>
                    </tr>
                </tbody>
            </table>
            <p style='font-size:0.72rem;color:#666;padding:6px 10px'>
                Entradas = itens em <b>status L</b> liberados para faturamento no dia (DTALTERAFAT).
                Resolvidas = itens com NF emitida no dia. Mostra o ritmo diário da fila de faturamento.
            </p>
        """

    def tabela_yoy_vendedores(df) -> str:
        if df.empty:
            return "<p style='color:#888;padding:12px'>Sem dados de vendedores para o comparativo YoY.</p>"

        ano = REF_ANO
        ano_ant = ano - 1
        rows = ""
        for _, r in df.head(15).iterrows():
            tipo_txt = "Externo" if r["Tipo"] == "E" else "Interno"
            var_abs = float(r["Var_Abs"] or 0)
            var_pct = r["Var_Pct"]
            cor = "#C6EFCE" if var_abs >= 0 else "#FFC7CE"
            var_pct_txt = "—" if pd.isna(var_pct) else f"{var_pct:+.1f}%"
            rows += f"""<tr style="background:{cor}">
                <td>{r['Vendedor'][:34]}</td>
                <td>{tipo_txt}</td>
                <td style="text-align:right">{fmt_brl(float(r['Fat_Anterior']))}</td>
                <td style="text-align:right">{fmt_brl(float(r['Fat_Atual']))}</td>
                <td style="text-align:right">{fmt_brl(var_abs)}</td>
                <td style="text-align:center"><b>{var_pct_txt}</b></td>
            </tr>"""
        return f"""<table class="tabela">
            <thead><tr>
                <th>Vendedor</th><th>Tipo</th><th>{mes_nome}/{ano_ant}</th>
                <th>{mes_nome}/{ano}</th><th>Variação R$</th><th>Variação %</th>
            </tr></thead><tbody>{rows}</tbody></table>
            <p style='font-size:0.72rem;color:#666;padding:6px 10px'>
              Ranking por vendas do mês atual. Verde = crescimento YoY, vermelho = queda YoY.
            </p>"""

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Dashboard Comercial — Helibombas</title>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: "Segoe UI", Arial, sans-serif; background: {COR_BG}; color: #222; }}

  /* ── Header ── */
  .header {{
    background: {COR_AZUL}; color: white;
    padding: 16px 32px; display: flex; align-items: center;
    justify-content: space-between;
  }}
  .header h1 {{ font-size: 1.4rem; font-weight: 700; letter-spacing: 1px; }}
  .header .sub {{ font-size: 0.85rem; opacity: 0.8; margin-top: 2px; }}
  .header .atualizado {{ font-size: 0.8rem; opacity: 0.7; text-align: right; }}
    .modo-badge {{
        display: inline-block;
        margin-top: 8px;
        padding: 6px 10px;
        border-radius: 999px;
        font-size: 0.72rem;
        font-weight: 700;
        letter-spacing: 0.8px;
        border: 1px solid rgba(255,255,255,0.55);
    }}
    .modo-base {{ background: rgba(112,173,71,0.28); color: #E9FFE0; }}
    .modo-sem {{ background: rgba(237,125,49,0.32); color: #FFF3E8; }}

    /* ── Navegação rápida ── */
    .quick-nav {{
        background: #FFFFFF;
        border-bottom: 1px solid #D9D9D9;
        padding: 10px 28px;
        display: flex;
        gap: 10px;
        position: sticky;
        top: 0;
        z-index: 50;
    }}
    .quick-link {{
        text-decoration: none;
        color: {COR_AZUL};
        font-size: 0.78rem;
        font-weight: 700;
        padding: 6px 10px;
        border: 1px solid #BFD3E6;
        border-radius: 999px;
        background: #F7FBFF;
    }}
    .quick-link:hover {{ filter: brightness(0.96); }}

    /* ── Abas de visão ── */
    .tabs-row {{
        background: #FFFFFF;
        padding: 8px 28px 0;
        display: flex;
        gap: 8px;
        border-bottom: 1px solid #E5E5E5;
    }}
    .tab-btn {{
        border: 1px solid #D0D7DE;
        border-bottom: none;
        background: #F7F9FB;
        color: {COR_AZUL};
        font-size: 0.78rem;
        font-weight: 700;
        padding: 8px 12px;
        border-top-left-radius: 8px;
        border-top-right-radius: 8px;
        cursor: pointer;
    }}
    .tab-btn.active {{
        background: #FFFFFF;
        color: #1D3B5A;
    }}
    .tab-pane {{ display: none; }}
    .tab-pane.active {{ display: block; }}

  /* ── KPI Cards ── */
  .kpi-row {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 14px; padding: 20px 28px;
  }}
  .kpi-card {{
    background: white; border-radius: 8px;
    padding: 16px 18px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.08);
    border-left: 5px solid {COR_AZUL_CLARO};
  }}
  .kpi-card.vermelho {{ border-left-color: {COR_VERMELHO}; }}
  .kpi-card.amarelo  {{ border-left-color: {COR_AMARELO}; }}
  .kpi-card.verde    {{ border-left-color: {COR_VERDE}; }}
  .kpi-card.laranja  {{ border-left-color: {COR_LARANJA}; }}
  .kpi-label {{ font-size: 0.72rem; color: #888; text-transform: uppercase; letter-spacing: 0.5px; }}
  .kpi-valor {{ font-size: 1.45rem; font-weight: 700; margin-top: 4px; color: {COR_AZUL}; }}
  .kpi-sub   {{ font-size: 0.75rem; color: #aaa; margin-top: 2px; }}

    /* ── Modo Reunião ── */
    .meeting-row {{
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
        gap: 14px; padding: 2px 28px 16px;
    }}
    .meeting-card {{
        background: white; border-radius: 8px;
        padding: 14px 16px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.08);
        border-left: 5px solid {COR_AZUL};
    }}
    .meeting-card.green {{ border-left-color: {COR_VERDE}; }}
    .meeting-card.red {{ border-left-color: {COR_VERMELHO}; }}
    .meeting-title {{ font-size: 0.72rem; color: #777; text-transform: uppercase; letter-spacing: 0.5px; }}
    .meeting-value {{ font-size: 1.10rem; font-weight: 700; color: {COR_AZUL}; margin-top: 4px; }}
    .meeting-sub {{ font-size: 0.78rem; color: #666; margin-top: 4px; }}

  /* ── Grid de gráficos ── */
  .grid-2 {{
    display: grid; grid-template-columns: 1fr 1fr;
    gap: 14px; padding: 0 28px 14px;
  }}
  .grid-wide {{ padding: 0 28px 14px; }}
  .card {{
    background: white; border-radius: 8px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.08);
    padding: 6px;
  }}

  /* ── Tabelas ── */
  .tabela {{ width: 100%; border-collapse: collapse; font-size: 0.80rem; }}
  .tabela th {{
    background: {COR_AZUL}; color: white;
    padding: 7px 10px; text-align: left; font-weight: 600;
  }}
  .tabela td {{ padding: 6px 10px; border-bottom: 1px solid #ddd; }}
  .tabela tr:hover {{ filter: brightness(0.96); }}

  .section-title {{
    font-size: 0.9rem; font-weight: 700; color: {COR_AZUL};
    padding: 6px 10px 4px; text-transform: uppercase; letter-spacing: 0.5px;
  }}

  /* ── Footer ── */
  .footer {{
    text-align: center; font-size: 0.72rem; color: #bbb;
    padding: 18px; border-top: 1px solid #ddd; margin-top: 10px;
  }}
</style>
</head>
<body>

<div class="header">
  <div>
    <div class="h1" style="font-size:1.5rem;font-weight:700">🔩 HELIBOMBAS</div>
    <div class="sub">Dashboard Comercial — Peças e Componentes</div>
        <div class="sub">Filtro ativo: <b>{dados['modo_filtro']}</b></div>
        <div class="sub">Período de referência: <b>{REF_MES:02d}/{REF_ANO}</b> (mês fechado)</div>
        <div class="modo-badge {modo_badge_class}">{modo_badge}</div>
  </div>
  <div class="atualizado">
    Atualizado em<br><b>{HOJE.strftime('%d/%m/%Y')}</b>
  </div>
</div>

<div class="quick-nav">
    <a class="quick-link" href="#sec-operacao">⚡ Operação</a>
    <a class="quick-link" href="#sec-cotacoes">🧾 Cotações</a>
    <a class="quick-link" href="#sec-estrategico">📈 Estratégico</a>
</div>

<div class="tabs-row">
    <button class="tab-btn active" data-tab="aba-fechado" onclick="abrirAba('aba-fechado')">📅 Mês Fechado</button>
    <button class="tab-btn" data-tab="aba-mtd" onclick="abrirAba('aba-mtd')">⚡ MTD Hoje</button>
</div>

<div id="aba-fechado" class="tab-pane active">

<!-- Modo Reunião: Highlights automáticos -->
<div class="meeting-row">
    <div class="meeting-card green">
        <div class="meeting-title">🎯 Crescimento Real YoY</div>
        <div class="meeting-value">{yoy_label} · {impacto_yoy_abs}</div>
        <div class="meeting-sub">{mes_nome}/{REF_ANO} vs {mes_nome}/{REF_ANO-1} ({dados['modo_filtro']})</div>
    </div>
    <div class="meeting-card green">
        <div class="meeting-title">🚀 Destaque Positivo</div>
        <div class="meeting-value">{resumo_vend['maior_alta']}</div>
        <div class="meeting-sub">Baseado na variação do mesmo mês do ano anterior</div>
    </div>
    <div class="meeting-card red">
        <div class="meeting-title">⚠️ Ponto de Atenção</div>
        <div class="meeting-value">{resumo_vend['maior_queda']}</div>
        <div class="meeting-sub">Indica prioridade de ação comercial no curto prazo</div>
    </div>
</div>

<!-- KPI Cards -->
<div class="kpi-row">
  <div class="kpi-card">
    <div class="kpi-label">Vendas {mes_nome}/{REF_ANO}</div>
    <div class="kpi-valor">{kpi_fat}</div>
    <div class="kpi-sub">{fat_mes['pedidos']} pedidos · {fat_mes['itens']} itens</div>
  </div>
    <div class="kpi-card {yoy_class}">
        <div class="kpi-label">📈 Crescimento YoY ({mes_nome})</div>
        <div class="kpi-valor">{yoy_label}</div>
        <div class="kpi-sub">{yoy_sub}</div>
    </div>
  <div class="kpi-card">
    <div class="kpi-label">Carteira em Aberto</div>
    <div class="kpi-valor">{kpi_cart}</div>
    <div class="kpi-sub">{len(cart)} itens em aberto</div>
  </div>
  <div class="kpi-card vermelho">
    <div class="kpi-label">🔴 Atrasados</div>
    <div class="kpi-valor">{n_atr}</div>
    <div class="kpi-sub">itens com prazo vencido</div>
  </div>
  <div class="kpi-card amarelo">
    <div class="kpi-label">🟡 Urgentes (≤ 7 dias)</div>
    <div class="kpi-valor">{n_urg}</div>
    <div class="kpi-sub">itens a vencer essa semana</div>
  </div>
  <div class="kpi-card laranja">
    <div class="kpi-label">📄 Aguardando NF</div>
    <div class="kpi-valor">{n_nf}</div>
    <div class="kpi-sub">{vlr_nf} parado</div>
  </div>
    <div class="kpi-card">
        <div class="kpi-label">🧾 Cotações {mes_nome}/{REF_ANO}</div>
        <div class="kpi-valor">{total_cot}</div>
        <div class="kpi-sub">Enc.:{cot_enc} · Perd.:{cot_perd} · Aberto:{cot_aberto} · {cot_vlr}</div>
        <div class="kpi-sub">Conv. geral (enc./criadas) qtd/valor: {cot_conv_qtd:.1f}% · {cot_conv_vlr:.1f}%</div>
        <div class="kpi-sub">Conv. finalizadas (enc./finaliz.) qtd/valor: {cot_conv_qtd_final:.1f}% · {cot_conv_vlr_final:.1f}%</div>
    </div>
  <div class="kpi-card verde">
    <div class="kpi-label">📊 Margem Média ({mes_nome})</div>
    <div class="kpi-valor">{fmt_pct(dados['margem_media'])}</div>
    <div class="kpi-sub">sobre itens com custo cadastrado</div>
  </div>
</div>

<!-- Modo Operação: foco em ação diária -->
<div class="grid-wide" id="sec-operacao">
    <div class="card">
        <div class="section-title">⚡ Modo Operação — Foco do Dia</div>
        <p style='font-size:0.78rem;color:#666;padding:6px 10px'>
          Priorize na ordem: <b>NF crítica</b> → <b>prazo atrasado</b> → <b>cotações em aberto</b>.
        </p>
    </div>
</div>

<!-- SLA diário operacional -->
<div class="grid-wide">
    <div class="card">
        <div class="section-title">⏱️ SLA Diário da Fila de Faturamento (Status L)</div>
        {card_sla_operacao(sla)}
    </div>
</div>

<!-- Tabela: Prioridades operacionais -->
<div class="grid-wide">
    <div class="card">
        <div class="section-title">🎯 Prioridades do Dia — Ação Comercial</div>
        {tabela_prioridades(df_prioridades)}
    </div>
</div>

<!-- Tabela: Cotações -->
<div class="grid-wide" id="sec-cotacoes">
    <div class="card">
        <div class="section-title">🧾 Cotações Geradas — Controle Comercial</div>
        {tabela_cotacoes(df_cot)}
    </div>
</div>

<!-- Linha: Funil e Conversão por território -->
<div class="grid-2">
    <div class="card">{html_funil}</div>
    <div class="card">
        <div class="section-title">🎯 Conversão por UF (Estado do Cliente)</div>
        {tabela_conversao_uf(df_cot)}
    </div>
</div>

<!-- Linha 1: Evolução + Semáforo -->
<div class="grid-2" id="sec-estrategico">
  <div class="card">{html_ev}</div>
  <div class="card">{html_sem}</div>
</div>

<!-- Linha 2: Ranking Externos -->
<div class="grid-wide">
    <div class="card">{html_ext}</div>
</div>

<!-- Linha 3: ABC Treemap + Performance Externos -->
<div class="grid-2">
  <div class="card">{html_abc}</div>
  <div class="card">{html_perf}</div>
</div>

<!-- Tabela: YoY por vendedor -->
<div class="grid-wide">
    <div class="card">
        <div class="section-title">📈 Crescimento YoY por Vendedor — {mes_nome}/{REF_ANO} vs {mes_nome}/{REF_ANO-1}</div>
        {tabela_yoy_vendedores(df_yoy_vendedores)}
    </div>
</div>

<!-- Tabela: Aguardando NF -->
<div class="grid-wide">
  <div class="card">
    <div class="section-title">📄 Top 10 Itens Aguardando Emissão de NF (mais antigos primeiro)</div>
    {tabela_nf(df_nf)}
  </div>
</div>

<!-- Tabela: Alertas de Margem -->
<div class="grid-wide">
  <div class="card">
    <div class="section-title">⚠️ Alertas de Margem Crítica — {mes_nome}/{REF_ANO} (abaixo de 20%)</div>
    {tabela_alertas(df_alert)}
  </div>
</div>

<div class="footer">
  Gerado automaticamente por fase4_dashboard.py · {datetime.now().strftime('%d/%m/%Y %H:%M')} ·
  Para atualizar: <code>python fase4_dashboard.py</code>
</div>

</div>

<div id="aba-mtd" class="tab-pane">
    <div class="kpi-row">
        <div class="kpi-card">
            <div class="kpi-label">Hoje ({HOJE.strftime('%d/%m')})</div>
            <div class="kpi-valor">{fmt_brl(mtd['faturamento_hoje'])}</div>
            <div class="kpi-sub">{mtd['pedidos_hoje']} pedidos no dia</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">MTD ({mtd['mes']:02d}/{mtd['ano']})</div>
            <div class="kpi-valor">{fmt_brl(mtd['faturamento_mtd'])}</div>
            <div class="kpi-sub">{mtd['pedidos_mtd']} pedidos · {mtd['itens_mtd']} itens</div>
        </div>
        <div class="kpi-card amarelo">
            <div class="kpi-label">Run-rate diário</div>
            <div class="kpi-valor">{fmt_brl(mtd['run_rate_dia'])}</div>
            <div class="kpi-sub">média por dia até hoje ({mtd['dia']}/{mtd['dias_mes']})</div>
        </div>
        <div class="kpi-card {mtd_meta_class}">
            <div class="kpi-label">Projeção de Fechamento</div>
            <div class="kpi-valor">{fmt_brl(mtd['projecao_mes'])}</div>
            <div class="kpi-sub">Meta: {fmt_brl(mtd['meta_mes'])} · Atingimento proj.: {mtd['pct_meta_proj']:.1f}%</div>
            <div class="kpi-sub">Gap projetado: {mtd_gap_sinal}{fmt_brl(mtd['gap_meta'])}</div>
        </div>
    </div>

    <div class="grid-wide">
        <div class="card">
            <div class="section-title">⚡ Leitura MTD em tempo real</div>
            <p style='font-size:0.78rem;color:#666;padding:6px 10px'>
              Esta aba mostra o mês em andamento até o momento da atualização do arquivo.
              Para operação diária, atualizar em ciclos curtos (ex.: a cada 1 hora) mantém a visão próxima do real.
            </p>
        </div>
    </div>
</div>

<script>
function abrirAba(tabId) {{
    document.querySelectorAll('.tab-pane').forEach(function(el) {{ el.classList.remove('active'); }});
    document.querySelectorAll('.tab-btn').forEach(function(el) {{ el.classList.remove('active'); }});
    const alvo = document.getElementById(tabId);
    if (alvo) alvo.classList.add('active');
    const botao = document.querySelector('.tab-btn[data-tab="' + tabId + '"]');
    if (botao) botao.classList.add('active');
}}
</script>

</body>
</html>"""


# ── Main ──────────────────────────────────────────────────────────────────────

def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument("--no-open", action="store_true",
                   help="Não abre o navegador automaticamente")
    p.add_argument("--sem-contrato", action="store_true",
                   help="Exclui clientes de contrato por nome (PETROBRAS/PETROLEO)")
    return p.parse_args()


def coletar_dados_dashboard(engine, sem_contrato: bool) -> tuple[dict, str]:
    padroes_contrato = ["PETROBRAS", "PETROLEO"] if sem_contrato else []
    filtro_cliente_extra = build_filtro_clientes_contrato(padroes_contrato)
    filtro_orc_extra = build_filtro_clientes_contrato_orc("f", padroes_contrato)
    modo_filtro = (
        "Sem contrato (PETROBRAS/PETROLEO)"
        if sem_contrato
        else "Base padrão (todos os clientes válidos)"
    )

    df_ev = get_evolucao_mensal(engine, filtro_cliente_extra)
    fat_mes = get_faturamento_mes_atual(engine, filtro_cliente_extra)
    df_cart = get_carteira(engine, filtro_cliente_extra)
    df_vend = get_ranking_vendedores(engine, filtro_cliente_extra)
    df_abc = get_abc_clientes(engine, filtro_cliente_extra)
    df_alert = get_margem_alertas(engine, filtro_cliente_extra)
    yoy_mes = get_crescimento_real_yoy(engine, filtro_cliente_extra)
    df_vend_yoy = get_yoy_vendedores_mes(engine, filtro_cliente_extra)
    df_cot = get_cotacoes_mes(engine, filtro_orc_extra)
    sla_hoje = get_sla_operacao_hoje(engine, filtro_cliente_extra)
    mtd_atual = get_faturamento_mtd_atual(engine, filtro_cliente_extra)

    ano = REF_ANO
    margem_media = float(
        df_vend[df_vend["Ano"] == ano]["Margem_pct"].mean() or 0
    )

    dados = {
        "fig_ev": fig_evolucao(df_ev),
        "fig_sem": fig_semaforo_carteira(df_cart),
        "fig_ext": fig_ranking_externos(df_vend),
        "fig_abc": fig_abc_treemap(df_abc),
        "fig_perf": fig_performance_externos(df_vend),
        "fig_cot_funil": fig_funil_cotacoes(df_cot),
        "fat_mes": fat_mes,
        "yoy_mes": yoy_mes,
        "df_vend_yoy": df_vend_yoy,
        "df_cart": df_cart,
        "df_alertas": df_alert,
        "df_cotacoes": df_cot,
        "sla_hoje": sla_hoje,
        "mtd_atual": mtd_atual,
        "margem_media": margem_media,
        "modo_filtro": modo_filtro,
    }
    return dados, modo_filtro


def main():
    args = parse_args()
    engine = get_engine()

    modo_filtro = (
        "Sem contrato (PETROBRAS/PETROLEO)"
        if args.sem_contrato
        else "Base padrão (todos os clientes válidos)"
    )

    print(f"Dashboard Comercial — {HOJE.strftime('%d/%m/%Y')}")
    print(f"  Modo de filtro: {modo_filtro}")
    print(f"  Período de referência: {REF_MES:02d}/{REF_ANO} (mês fechado)")
    print("  Coletando dados...")

    print("  Gerando gráficos...")
    dados_base, _ = coletar_dados_dashboard(engine, sem_contrato=False)
    dados_sem, _ = coletar_dados_dashboard(engine, sem_contrato=True)

    arquivo_base = OUTPUT / "dashboard_base.html"
    arquivo_sem = OUTPUT / "dashboard_sem_contrato.html"
    arquivo_base.write_text(gerar_html(dados_base), encoding="utf-8")
    arquivo_sem.write_text(gerar_html(dados_sem), encoding="utf-8")

    dados_principal = dados_sem if args.sem_contrato else dados_base
    arquivo = OUTPUT / "dashboard.html"
    arquivo.write_text(gerar_html(dados_principal), encoding="utf-8")

    df_prioridades = construir_prioridades_operacionais(dados_principal["df_cart"])
    if not df_prioridades.empty:
        cols_export = [
            "Prioridade", "Acao_Sugerida", "Vendedor", "Cliente", "Pedido",
            "MATERIAL", "DESCRICAO", "Vlr_Saldo", "Dias_p_Prazo", "Dias_Aguard_NF", "Status_Item", "Semaforo"
        ]
        df_exp = df_prioridades[[c for c in cols_export if c in df_prioridades.columns]].copy()
        arquivo_fila = OUTPUT / "fila_acao_diaria.xlsx"
        arquivo_fila_ts = OUTPUT / f"fila_acao_diaria_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        df_exp.to_excel(arquivo_fila, index=False)
        df_exp.to_excel(arquivo_fila_ts, index=False)

    arquivo_reuniao = OUTPUT / "dashboard_reuniao.html"
    arquivo_reuniao.write_text(gerar_html(dados_sem), encoding="utf-8")

    print(f"\n✅ Dashboard principal: {arquivo.resolve()}")
    print(f"✅ Versão base: {arquivo_base.resolve()}")
    print(f"✅ Versão sem contrato: {arquivo_sem.resolve()}")
    print(f"✅ Versão reunião: {arquivo_reuniao.resolve()}")
    if not df_prioridades.empty:
        print(f"✅ Fila de ação diária: {(OUTPUT / 'fila_acao_diaria.xlsx').resolve()}")
    print(f"\nResumo rápido {REF_MES:02d}/{REF_ANO}:")
    print(f"  Base: {fmt_brl(dados_base['fat_mes']['faturamento'])}")
    print(f"  Sem contrato: {fmt_brl(dados_sem['fat_mes']['faturamento'])}")

    if not args.no_open:
        import webbrowser
        webbrowser.open(arquivo.resolve().as_uri())
        print("   Abrindo no navegador...")


if __name__ == "__main__":
    main()
