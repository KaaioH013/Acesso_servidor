"""
relatorio_comissoes.py — Relatório de Comissões por Competência (Financeiro)

Regra aplicada:
- Comissão entra na competência do mês seguinte ao pagamento.
- Considera apenas NFs com 100% das parcelas quitadas no mês de pagamento-base.
- Data gatilho = MAX(DTPAGAMENTO) da FN_RECEBER (última parcela paga).
- Separação de comissão por tipo de item:
    * BOMBA  -> MATERIAL LIKE '8%'
    * PEÇA   -> demais materiais

Uso:
    python relatorio_comissoes.py
    python relatorio_comissoes.py --mes 3 --ano 2026
    python relatorio_comissoes.py --vendedor "RAFAEL"
"""

import sys
import argparse
from datetime import date, datetime
from pathlib import Path
import re
import unicodedata

import pandas as pd

sys.path.insert(0, "src")
from conexao import get_engine

OUTPUT = Path("exports")
OUTPUT.mkdir(exist_ok=True)


def default_competencia() -> tuple[int, int]:
    hoje = date.today()
    return hoje.month, hoje.year


def mes_anterior(mes: int, ano: int) -> tuple[int, int]:
    if mes == 1:
        return 12, ano - 1
    return mes - 1, ano


def parse_args():
    mes_default, ano_default = default_competencia()
    p = argparse.ArgumentParser()
    p.add_argument("--mes", type=int, default=mes_default,
                   help="Mês da competência de comissão (1-12)")
    p.add_argument("--ano", type=int, default=ano_default,
                   help="Ano da competência de comissão")
    p.add_argument("--vendedor", type=str, default="",
                   help="Filtro opcional por nome do vendedor (ex.: RAFAEL)")
    p.add_argument("--incluir-exterior", action="store_true",
                   help="Inclui clientes com UF='EX' (padrão: exclui exterior)")
    p.add_argument("--arquivo-regras", type=str, default="config/regras_comissao_vendedores.csv",
                   help="CSV com regras de comissão por vendedor")
    return p.parse_args()


def primeiro_dia_mes(mes: int, ano: int) -> str:
    return f"{ano:04d}-{mes:02d}-01"


def primeiro_dia_mes_seguinte(mes: int, ano: int) -> str:
    if mes == 12:
        return f"{ano + 1:04d}-01-01"
    return f"{ano:04d}-{mes + 1:02d}-01"


def query_comissao_competencia(
    engine,
    ini_pagamento: str,
    fim_pagamento: str,
    vendedor_filtro: str,
    incluir_exterior: bool,
) -> pd.DataFrame:
    filtro_vendedor = ""
    if vendedor_filtro.strip():
        nome = vendedor_filtro.upper().replace("'", "''")
        filtro_vendedor = f"AND UPPER(v.RAZAO) LIKE '%{nome}%'"

    filtro_exterior = ""
    if not incluir_exterior:
        filtro_exterior = "AND ISNULL(c.UF, '') <> 'EX'"

    sql = f"""
WITH NfsQuitadas AS (
    SELECT
        r.NFS,
        COUNT(*) AS Parcelas_Total,
        SUM(CASE WHEN r.DTPAGAMENTO IS NOT NULL THEN 1 ELSE 0 END) AS Parcelas_Pagas,
        MAX(r.DTPAGAMENTO) AS Dt_Ultima_Parcela,
        SUM(ISNULL(r.VLREFETIVO, 0)) AS Vlr_Recebido_Parcelas,
        SUM(ISNULL(r.VLRBASECOMISSAO, 0)) AS Vlr_Base_Comissao_Receber
    FROM FN_RECEBER r
    WHERE r.NFS IS NOT NULL
      AND r.STATUS = 'A'
      AND r.DTCANCELAMENTO IS NULL
    GROUP BY r.NFS
    HAVING
        COUNT(*) = SUM(CASE WHEN r.DTPAGAMENTO IS NOT NULL THEN 1 ELSE 0 END)
        AND MAX(r.DTPAGAMENTO) >= '{ini_pagamento}'
        AND MAX(r.DTPAGAMENTO) <  '{fim_pagamento}'
),
ItensNf AS (
    SELECT
        ni.NFE AS NFS,
        SUM(CASE WHEN ni.MATERIAL LIKE '8%' THEN ISNULL(ni.VLRTOTAL, 0) ELSE 0 END) AS Vlr_Bombas,
        SUM(CASE WHEN ni.MATERIAL NOT LIKE '8%' THEN ISNULL(ni.VLRTOTAL, 0) ELSE 0 END) AS Vlr_Pecas,
        SUM(ISNULL(ni.VLRTOTAL, 0)) AS Vlr_Total_Itens,
        COUNT(*) AS Itens
    FROM FN_NFEITENS ni
    GROUP BY ni.NFE
)
SELECT
    q.NFS,
    n.NRONOTA,
    n.DTEMISSAO,
    n.DTSAIDA,
    q.Dt_Ultima_Parcela,
    q.Parcelas_Total,
    q.Parcelas_Pagas,
    v.CODIGO AS Cod_Vendedor,
    v.RAZAO AS Vendedor,
    v.TIPO AS Tipo_Vendedor,
    v.UF AS Vendedor_UF,
    v.COBUF AS Vendedor_CobUF,
    c.CODIGO AS Cod_Cliente,
    c.RAZAO AS Cliente,
    c.UF AS Cliente_UF,
    c.CIDADE AS Cliente_Cidade,
    q.Vlr_Recebido_Parcelas,
    q.Vlr_Base_Comissao_Receber,
    ISNULL(i.Vlr_Total_Itens, 0) AS Vlr_Total_Itens,
    ISNULL(i.Vlr_Pecas, 0) AS Vlr_Pecas,
    ISNULL(i.Vlr_Bombas, 0) AS Vlr_Bombas,
    ISNULL(i.Itens, 0) AS Itens
FROM NfsQuitadas q
JOIN FN_NFS n ON n.CODIGO = q.NFS
LEFT JOIN ItensNf i ON i.NFS = q.NFS
LEFT JOIN FN_VENDEDORES v ON v.CODIGO = n.VENDEDOR
LEFT JOIN FN_FORNECEDORES c ON c.CODIGO = n.CLIENTE
WHERE n.STATUSNF <> 'C'
    {filtro_exterior}
  {filtro_vendedor}
ORDER BY q.Dt_Ultima_Parcela, v.RAZAO, n.NRONOTA
"""

    return pd.read_sql(sql, engine)


def normalizar_texto(valor: str) -> str:
    if valor is None:
        return ""
    txt = str(valor).strip().upper()
    txt = unicodedata.normalize("NFKD", txt).encode("ASCII", "ignore").decode("ASCII")
    txt = re.sub(r"\s+", " ", txt)
    return txt


def parse_regiao_ufs(regiao: str) -> set[str]:
    r = normalizar_texto(regiao)
    if not r:
        return set()
    if "TODAS" in r:
        return {"*"}

    ufs_brasil = {
        "AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS",
        "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RS", "RO", "RR", "SC",
        "SP", "SE", "TO"
    }
    nordeste = {"AL", "BA", "CE", "MA", "PB", "PE", "PI", "RN", "SE"}

    if "NORDESTE" in r:
        return nordeste

    encontrados = set(re.findall(r"\b[A-Z]{2}\b", r))
    return {uf for uf in encontrados if uf in ufs_brasil}


def carregar_regras_comissao(arquivo_regras: str) -> pd.DataFrame:
    path = Path(arquivo_regras)
    if not path.exists():
        return pd.DataFrame()

    regras = pd.read_csv(path)
    for col in ["vendedor", "tipo", "regiao", "observacao"]:
        if col not in regras.columns:
            regras[col] = ""

    for col in ["perc_bombas", "perc_pecas"]:
        regras[col] = pd.to_numeric(regras[col], errors="coerce").fillna(0.0)

    for col in ["ano_inicio", "ano_fim"]:
        if col not in regras.columns:
            regras[col] = pd.NA

    regras["ano_inicio"] = pd.to_numeric(regras["ano_inicio"], errors="coerce")
    regras["ano_fim"] = pd.to_numeric(regras["ano_fim"], errors="coerce")
    regras["vend_norm"] = regras["vendedor"].map(normalizar_texto).str.replace("*", "", regex=False).str.strip()
    regras["tipo_norm"] = regras["tipo"].map(normalizar_texto)
    regras["regiao_norm"] = regras["regiao"].map(normalizar_texto)
    regras["ufs_regra"] = regras["regiao"].map(parse_regiao_ufs)
    return regras


def selecionar_regra_vendedor(regras_vendedor: pd.DataFrame, cliente_uf: str, ano_comp: int):
    if regras_vendedor.empty:
        return None

    uf = normalizar_texto(cliente_uf)
    candidatas = regras_vendedor.copy()

    cond_vigencia = (
        (candidatas["ano_inicio"].isna() | (candidatas["ano_inicio"] <= ano_comp))
        & (candidatas["ano_fim"].isna() | (candidatas["ano_fim"] >= ano_comp))
    )
    candidatas = candidatas[cond_vigencia]
    if candidatas.empty:
        candidatas = regras_vendedor.copy()

    especificas = candidatas[candidatas["ufs_regra"].map(lambda s: uf in s if isinstance(s, set) else False)]
    if not especificas.empty:
        especificas = especificas.assign(_ordem=especificas["ufs_regra"].map(len)).sort_values("_ordem")
        return especificas.iloc[0]

    genericas = candidatas[candidatas["ufs_regra"].map(lambda s: "*" in s if isinstance(s, set) else False)]
    if not genericas.empty:
        return genericas.iloc[0]

    return candidatas.iloc[0]


def selecionar_regra_representante(
    regras_rep: pd.DataFrame,
    cliente_uf: str,
    ano_comp: int,
    vendedor_nf_norm: str,
):
    if regras_rep.empty:
        return None, "SEM_TABELA"

    candidatas = regras_rep.copy()
    cond_vigencia = (
        (candidatas["ano_inicio"].isna() | (candidatas["ano_inicio"] <= ano_comp))
        & (candidatas["ano_fim"].isna() | (candidatas["ano_fim"] >= ano_comp))
    )
    candidatas = candidatas[cond_vigencia]
    if candidatas.empty:
        return None, "SEM_VIGENCIA"

    uf = normalizar_texto(cliente_uf)
    por_uf = candidatas[candidatas["ufs_regra"].map(lambda s: uf in s if isinstance(s, set) else False)]
    if por_uf.empty:
        return None, "SEM_REGIAO"

    match_nome = por_uf[por_uf["vend_norm"] == vendedor_nf_norm]
    if not match_nome.empty:
        return match_nome.iloc[0], "REGIAO_NOME"

    if len(por_uf) == 1:
        return por_uf.iloc[0], "REGIAO_UNICA"

    return None, "REGIAO_AMBIGUA"


def aplicar_regras_comissao(
    df: pd.DataFrame,
    regras: pd.DataFrame,
    ano_competencia: int,
) -> pd.DataFrame:
    if df.empty:
        return df

    df2 = df.copy()
    df2["Vend_Norm"] = df2["Vendedor"].map(normalizar_texto)

    regras_rep = regras[regras["tipo_norm"].isin(["REP", "EXTERNO"])].copy() if not regras.empty else pd.DataFrame()

    comissionado = []
    tipo_comissionado = []
    perc_peca = []
    perc_bomba = []
    regra_origem = []
    regra_regiao = []
    regra_obs = []

    for _, row in df2.iterrows():
        vend_nf = row["Vend_Norm"]
        cliente_uf = row.get("Cliente_UF", "")

        regra, motivo = selecionar_regra_representante(
            regras_rep=regras_rep,
            cliente_uf=cliente_uf,
            ano_comp=ano_competencia,
            vendedor_nf_norm=vend_nf,
        )

        if regra is None:
            comissionado.append("HELIBOMBAS")
            tipo_comissionado.append("EMPRESA")
            perc_peca.append(0.0)
            perc_bomba.append(0.0)
            regra_origem.append("HELIBOMBAS")
            regra_regiao.append(normalizar_texto(cliente_uf))
            regra_obs.append(f"Sem regra aplicável ({motivo})")
            continue

        comissionado.append(str(regra.get("vendedor", "")))
        tipo_comissionado.append("REP")
        perc_peca.append(float(regra["perc_pecas"]))
        perc_bomba.append(float(regra["perc_bombas"]))
        regra_origem.append("TABELA")
        regra_regiao.append(str(regra.get("regiao", "")))
        regra_obs.append(str(regra.get("observacao", "")))

    df2["Comissionado"] = comissionado
    df2["Tipo_Comissionado"] = tipo_comissionado
    df2["Perc_Peca"] = perc_peca
    df2["Perc_Bomba"] = perc_bomba
    df2["Regra_Origem"] = regra_origem
    df2["Regra_Regiao"] = regra_regiao
    df2["Regra_Observacao"] = regra_obs

    df2["Vlr_Comissao_Pecas"] = (df2["Vlr_Pecas"] * (df2["Perc_Peca"] / 100.0)).round(2)
    df2["Vlr_Comissao_Bombas"] = (df2["Vlr_Bombas"] * (df2["Perc_Bomba"] / 100.0)).round(2)
    df2["Vlr_Comissao_Calculada"] = (df2["Vlr_Comissao_Pecas"] + df2["Vlr_Comissao_Bombas"]).round(2)

    return df2


def gerar_excel(df: pd.DataFrame, competencia_mes: int, competencia_ano: int) -> Path:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo = OUTPUT / f"comissoes_{competencia_mes:02d}-{competencia_ano}_{ts}.xlsx"

    if df.empty:
        with pd.ExcelWriter(arquivo, engine="openpyxl") as writer:
            pd.DataFrame([{
                "Mensagem": "Nenhuma NF totalmente quitada no mês-base para esta competência."
            }]).to_excel(writer, sheet_name="Resumo", index=False)
        return arquivo

    resumo_vendedor = (
        df.groupby(["Comissionado", "Tipo_Comissionado", "Regra_Regiao", "Regra_Origem"], dropna=False)
          .agg(
              NFs=("NFS", "nunique"),
              Clientes=("Cod_Cliente", "nunique"),
              Vlr_Pecas=("Vlr_Pecas", "sum"),
              Vlr_Bombas=("Vlr_Bombas", "sum"),
              Vlr_Base_Comissao=("Vlr_Base_Comissao_Receber", "sum"),
              Vlr_Comissao=("Vlr_Comissao_Calculada", "sum"),
          )
          .reset_index()
          .sort_values("Vlr_Comissao", ascending=False)
    )

    resumo_uf = (
        df.groupby(["Cliente_UF"], dropna=False)
          .agg(
              NFs=("NFS", "nunique"),
                            Comissionados=("Comissionado", "nunique"),
              Vlr_Pecas=("Vlr_Pecas", "sum"),
              Vlr_Bombas=("Vlr_Bombas", "sum"),
              Vlr_Comissao=("Vlr_Comissao_Calculada", "sum"),
          )
          .reset_index()
          .sort_values("Vlr_Comissao", ascending=False)
    )

    resumo_cidade = (
        df.groupby(["Cliente_UF", "Cliente_Cidade"], dropna=False)
          .agg(
              NFs=("NFS", "nunique"),
                            Comissionados=("Comissionado", "nunique"),
              Vlr_Pecas=("Vlr_Pecas", "sum"),
              Vlr_Bombas=("Vlr_Bombas", "sum"),
              Vlr_Comissao=("Vlr_Comissao_Calculada", "sum"),
          )
          .reset_index()
          .sort_values(["Cliente_UF", "Vlr_Comissao"], ascending=[True, False])
    )

    resumo_tipo = pd.DataFrame([
        {
            "Tipo": "Peças",
            "Percentual_Medio": round((df["Vlr_Comissao_Pecas"].sum() / df["Vlr_Pecas"].sum() * 100.0), 4) if df["Vlr_Pecas"].sum() else 0.0,
            "Base": float(df["Vlr_Pecas"].sum()),
            "Comissao": float(df["Vlr_Comissao_Pecas"].sum()),
        },
        {
            "Tipo": "Bombas",
            "Percentual_Medio": round((df["Vlr_Comissao_Bombas"].sum() / df["Vlr_Bombas"].sum() * 100.0), 4) if df["Vlr_Bombas"].sum() else 0.0,
            "Base": float(df["Vlr_Bombas"].sum()),
            "Comissao": float(df["Vlr_Comissao_Bombas"].sum()),
        },
    ])

    kpis = pd.DataFrame([
        {
            "Competencia": f"{competencia_mes:02d}/{competencia_ano}",
            "NFs_aptas": int(df["NFS"].nunique()),
            "Comissionados": int(df["Comissionado"].nunique()),
            "Clientes": int(df["Cod_Cliente"].nunique()),
            "Total_Recebido_Parcelas": float(df["Vlr_Recebido_Parcelas"].sum()),
            "Base_Comissao_Receber": float(df["Vlr_Base_Comissao_Receber"].sum()),
            "Vlr_Pecas": float(df["Vlr_Pecas"].sum()),
            "Vlr_Bombas": float(df["Vlr_Bombas"].sum()),
            "Comissao_Total_Calculada": float(df["Vlr_Comissao_Calculada"].sum()),
            "Comissao_Origem_Tabela": int((df["Regra_Origem"] == "TABELA").sum()),
            "Comissao_Helibombas": int((df["Regra_Origem"] == "HELIBOMBAS").sum()),
        }
    ])

    regras_aplicadas = (
        df.groupby(["Comissionado", "Tipo_Comissionado", "Regra_Origem", "Regra_Regiao", "Perc_Peca", "Perc_Bomba"], dropna=False)
          .agg(
              NFs=("NFS", "nunique"),
              Vlr_Comissao=("Vlr_Comissao_Calculada", "sum"),
          )
          .reset_index()
          .sort_values("Vlr_Comissao", ascending=False)
    )

    sem_regra = (
        df[df["Regra_Origem"] == "HELIBOMBAS"]
        .groupby(["Comissionado", "Tipo_Comissionado", "Cliente_UF"], dropna=False)
        .agg(
            NFs=("NFS", "nunique"),
            Clientes=("Cod_Cliente", "nunique"),
            Vlr_Pecas=("Vlr_Pecas", "sum"),
            Vlr_Bombas=("Vlr_Bombas", "sum"),
            Vlr_Comissao=("Vlr_Comissao_Calculada", "sum"),
        )
        .reset_index()
        .sort_values("Vlr_Comissao", ascending=False)
    )

    with pd.ExcelWriter(arquivo, engine="openpyxl") as writer:
        kpis.to_excel(writer, sheet_name="KPIs", index=False)
        resumo_vendedor.to_excel(writer, sheet_name="Resumo_Vendedor", index=False)
        resumo_uf.to_excel(writer, sheet_name="Resumo_UF", index=False)
        resumo_cidade.to_excel(writer, sheet_name="Resumo_Cidade", index=False)
        resumo_tipo.to_excel(writer, sheet_name="Pecas_x_Bombas", index=False)
        regras_aplicadas.to_excel(writer, sheet_name="Regras_Aplicadas", index=False)
        sem_regra.to_excel(writer, sheet_name="Sem_Regra", index=False)

        cols_detalhe = [
            "NFS", "NRONOTA", "DTEMISSAO", "Dt_Ultima_Parcela",
            "Parcelas_Total", "Parcelas_Pagas",
            "Cod_Vendedor", "Vendedor", "Tipo_Vendedor", "Vendedor_UF", "Vendedor_CobUF",
            "Comissionado", "Tipo_Comissionado",
            "Cod_Cliente", "Cliente", "Cliente_UF", "Cliente_Cidade",
            "Vlr_Recebido_Parcelas", "Vlr_Base_Comissao_Receber",
            "Vlr_Total_Itens", "Vlr_Pecas", "Vlr_Bombas",
            "Perc_Peca", "Perc_Bomba",
            "Vlr_Comissao_Pecas", "Vlr_Comissao_Bombas", "Vlr_Comissao_Calculada",
            "Regra_Origem", "Regra_Regiao", "Regra_Observacao",
        ]
        detalhe = df[cols_detalhe].copy()
        detalhe.to_excel(writer, sheet_name="Detalhado", index=False)

    return arquivo


def main():
    args = parse_args()

    if args.mes < 1 or args.mes > 12:
        raise ValueError("--mes deve estar entre 1 e 12")

    mes_pag, ano_pag = mes_anterior(args.mes, args.ano)
    ini_pag = primeiro_dia_mes(mes_pag, ano_pag)
    fim_pag = primeiro_dia_mes_seguinte(mes_pag, ano_pag)

    print(f"Relatório de Comissões — Competência {args.mes:02d}/{args.ano}")
    print(f"  Mês-base de pagamento: {mes_pag:02d}/{ano_pag} ({ini_pag} até {fim_pag})")
    print("  Regra: comissão somente pela tabela de representantes externos")
    print(f"  Arquivo de regras: {args.arquivo_regras}")
    print("  Internos: não comissionam (apenas lançam pedido)")
    print("  Escopo geográfico: Brasil (sem EX)" if not args.incluir_exterior else "  Escopo geográfico: Brasil + exterior")
    if args.vendedor.strip():
        print(f"  Filtro vendedor: {args.vendedor}")

    engine = get_engine()
    df = query_comissao_competencia(
        engine,
        ini_pagamento=ini_pag,
        fim_pagamento=fim_pag,
        vendedor_filtro=args.vendedor,
        incluir_exterior=args.incluir_exterior,
    )

    regras = carregar_regras_comissao(args.arquivo_regras)
    df = aplicar_regras_comissao(
        df,
        regras,
        ano_competencia=args.ano,
    )

    arquivo = gerar_excel(df, args.mes, args.ano)

    print(f"\n✅ Arquivo gerado: {arquivo.resolve()}")
    if not df.empty:
        print(f"   NFs aptas: {df['NFS'].nunique()}")
        print(f"   Base peças: R$ {df['Vlr_Pecas'].sum():,.2f}")
        print(f"   Base bombas: R$ {df['Vlr_Bombas'].sum():,.2f}")
        print(f"   Comissão calculada total: R$ {df['Vlr_Comissao_Calculada'].sum():,.2f}")
        print(f"   Regras por tabela: {(df['Regra_Origem'] == 'TABELA').sum()} | Helibombas (sem regra/região): {(df['Regra_Origem'] == 'HELIBOMBAS').sum()}")
    else:
        print("   Sem dados para a competência informada.")


if __name__ == "__main__":
    main()
