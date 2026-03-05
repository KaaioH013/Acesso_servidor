import argparse
from datetime import datetime, date
from pathlib import Path

import pandas as pd

import fase4_dashboard as f4
from src.conexao import get_engine

OUTPUT = Path("exports")
OUTPUT.mkdir(exist_ok=True)


def parse_args():
    parser = argparse.ArgumentParser(
        description="Relatório automático de cobrança por vendedor com UF e cidade"
    )
    parser.add_argument("--uf", type=str, default=None, help="Filtra uma UF específica (ex: MG)")
    parser.add_argument(
        "--inicio",
        type=str,
        default=None,
        help="Data inicial de faturamento (YYYY-MM-DD). Padrão: início do ano atual",
    )
    parser.add_argument(
        "--fim",
        type=str,
        default=None,
        help="Data final de faturamento (YYYY-MM-DD). Padrão: hoje",
    )
    parser.add_argument(
        "--vencimento-ate",
        type=str,
        default=None,
        help="Filtra títulos com vencimento até a data YYYY-MM-DD",
    )
    parser.add_argument(
        "--somente-vencidos",
        action="store_true",
        help="Mostra apenas títulos vencidos até hoje",
    )
    return parser.parse_args()


def montar_cotacao(numero, seq):
    if pd.isna(numero):
        return ""
    try:
        numero_i = int(float(numero))
        if pd.isna(seq):
            return str(numero_i)
        seq_i = int(float(seq))
        return f"{numero_i}.{seq_i:05d}"
    except Exception:
        return str(numero)


def montar_cotacao_final(numero_interno, orc_numero, orc_seq):
    if pd.notna(numero_interno) and str(numero_interno).strip() not in {"", "None", "nan"}:
        return str(numero_interno).strip()
    return montar_cotacao(orc_numero, orc_seq)


def default_periodo():
    hoje = date.today()
    inicio = date(hoje.year, 1, 1)
    fim = hoje
    return inicio.isoformat(), fim.isoformat()


def query_base(engine, inicio, fim, uf=None, vencimento_ate=None, somente_vencidos=False):
    filtro_uf = ""
    if uf:
        uf_limpa = str(uf).strip().upper().replace("'", "''")
        filtro_uf = f" AND ISNULL(cli.UF,'') = '{uf_limpa}' "

    filtro_venc = ""
    if vencimento_ate:
        filtro_venc = f" AND rec.Vencimento <= '{vencimento_ate}' "

    filtro_vencidos = ""
    if somente_vencidos:
        filtro_vencidos = " AND rec.Vencimento < CAST(GETDATE() AS date) "

    sql = f"""
    WITH ReceberNF AS (
        SELECT
            r.NFS,
            MAX(r.DTVENCIMENTO) AS Vencimento,
            SUM(CASE WHEN r.DTCANCELAMENTO IS NULL THEN ISNULL(r.VLRDEVIDO, 0) ELSE 0 END) AS Valor_Parcelado,
            SUM(CASE WHEN r.DTCANCELAMENTO IS NULL AND r.DTPAGAMENTO IS NULL THEN ISNULL(r.VLRDEVIDO, 0) ELSE 0 END) AS Valor_Em_Aberto
        FROM FN_RECEBER r
        WHERE r.NFS IS NOT NULL
        GROUP BY r.NFS
    )
    SELECT
        cli.RAZAO AS Cliente,
        n.NRONOTA AS NF,
        p.PEDIDOCLI AS Pedido_Cliente,
        n.VEPEDIDO AS PV,
        p.NUMINTERNO AS Numero_Interno,
        p.PEDORIGEM AS Orcamento_Codigo,
        o.NUMERO AS Orcamento_Numero,
        o.SEQ AS Orcamento_Seq,
        n.DTEMISSAO AS Data_Faturamento,
        rec.Vencimento,
        ISNULL(n.VLRLIQUIDO, ISNULL(n.VLRTOTAL, 0)) AS Valor,
        ISNULL(rec.Valor_Parcelado, 0) AS Valor_Parcelado,
        ISNULL(rec.Valor_Em_Aberto, 0) AS Valor_Em_Aberto,
        ISNULL(cli.UF, '') AS UF,
        ISNULL(cli.CIDADE, '') AS Cidade,
        p.CODIGO AS Pedido,
        n.CODIGO AS NFS
    FROM FN_NFS n
    LEFT JOIN ReceberNF rec ON rec.NFS = n.CODIGO
    LEFT JOIN VE_PEDIDO p ON p.CODIGO = n.VEPEDIDO
    LEFT JOIN VE_ORCAMENTOS o ON o.CODIGO = p.PEDORIGEM
    LEFT JOIN FN_FORNECEDORES cli ON cli.CODIGO = n.CLIENTE
    WHERE n.STATUSNF <> 'C'
      AND n.DTEMISSAO >= '{inicio}'
      AND n.DTEMISSAO < DATEADD(DAY, 1, '{fim}')
      AND ISNULL(cli.UF, '') <> 'EX'
      {filtro_uf}
      {filtro_venc}
      {filtro_vencidos}
      AND EXISTS (
          SELECT 1
          FROM VE_PEDIDOITENS i
          WHERE i.PEDIDO = p.CODIGO
            AND i.STATUS <> 'C'
            AND i.FLAGSUB <> 'S'
            AND i.MATERIAL NOT LIKE '8%'
            AND i.TPVENDA NOT IN ({f4.TPVENDA_STR})
      )
    ORDER BY UF, Cidade, Vencimento, Cliente
    """
    return pd.read_sql(sql, engine)


def aplicar_regras_vendedor(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    df = df.copy()
    df["Responsavel"] = df.apply(
        lambda r: f4.mapear_representante_externo(r["UF"], r["Cidade"], r["Data_Faturamento"]),
        axis=1,
    )

    cidades_excluidas = f4.carregar_cidades_excluidas_mg_alexandre()
    df["Cidade_Norm"] = df["Cidade"].map(f4.normalizar_texto)
    df["Cidade_Excluida_Alexandre"] = df["Cidade_Norm"].isin(cidades_excluidas).map(
        lambda x: "SIM" if x else "NAO"
    )

    df["Cotacao"] = df.apply(
        lambda r: montar_cotacao_final(r["Numero_Interno"], r["Orcamento_Numero"], r["Orcamento_Seq"]),
        axis=1,
    )

    return df


def exportar_excel(df: pd.DataFrame, uf=None):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    sufixo_uf = f"_{str(uf).upper()}" if uf else ""

    arquivo_master = OUTPUT / f"relatorio_cobranca_vendedores{sufixo_uf}_{ts}.xlsx"

    if df.empty:
        with pd.ExcelWriter(arquivo_master, engine="openpyxl") as writer:
            pd.DataFrame([{"Mensagem": "Sem dados para os filtros informados."}]).to_excel(
                writer, sheet_name="Resumo", index=False
            )
        return arquivo_master, []

    colunas_layout = [
        "Cliente",
        "NF",
        "Pedido_Cliente",
        "PV",
        "Cotacao",
        "Data_Faturamento",
        "Vencimento",
        "Valor",
        "UF",
        "Cidade",
        "Responsavel",
        "Cidade_Excluida_Alexandre",
    ]

    df_layout = df[colunas_layout].copy()

    resumo_resp = (
        df_layout.groupby(["Responsavel"], dropna=False)
        .agg(
            Titulos=("NF", "count"),
            Clientes=("Cliente", "nunique"),
            Valor_Total=("Valor", "sum"),
        )
        .reset_index()
        .sort_values("Valor_Total", ascending=False)
    )

    resumo_resp_uf = (
        df_layout.groupby(["Responsavel", "UF"], dropna=False)
        .agg(
            Titulos=("NF", "count"),
            Clientes=("Cliente", "nunique"),
            Valor_Total=("Valor", "sum"),
        )
        .reset_index()
        .sort_values(["Responsavel", "Valor_Total"], ascending=[True, False])
    )

    with pd.ExcelWriter(arquivo_master, engine="openpyxl") as writer:
        resumo_resp.to_excel(writer, sheet_name="Resumo_Vendedor", index=False)
        resumo_resp_uf.to_excel(writer, sheet_name="Resumo_Vendedor_UF", index=False)
        df_layout.sort_values(["Responsavel", "UF", "Cidade", "Vencimento", "Cliente"]).to_excel(
            writer, sheet_name="Detalhado", index=False
        )

    arquivos_vendedor = []
    for vendedor, bloco in df_layout.groupby("Responsavel", dropna=False):
        nome = str(vendedor).strip() if pd.notna(vendedor) else "SEM_RESPONSAVEL"
        nome_safe = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in nome.upper()).strip("_")
        nome_safe = nome_safe or "SEM_RESPONSAVEL"

        arq = OUTPUT / f"relatorio_cobranca_{nome_safe}{sufixo_uf}_{ts}.xlsx"
        resumo_uf = (
            bloco.groupby(["UF", "Cidade"], dropna=False)
            .agg(Titulos=("NF", "count"), Clientes=("Cliente", "nunique"), Valor_Total=("Valor", "sum"))
            .reset_index()
            .sort_values(["UF", "Valor_Total"], ascending=[True, False])
        )

        with pd.ExcelWriter(arq, engine="openpyxl") as writer:
            resumo_uf.to_excel(writer, sheet_name="Resumo_UF_Cidade", index=False)
            bloco.sort_values(["UF", "Cidade", "Vencimento", "Cliente"]).to_excel(
                writer, sheet_name="Detalhado", index=False
            )

        arquivos_vendedor.append(arq)

    return arquivo_master, arquivos_vendedor


def main():
    args = parse_args()
    engine = get_engine()

    inicio_default, fim_default = default_periodo()
    inicio = args.inicio or inicio_default
    fim = args.fim or fim_default

    df = query_base(
        engine,
        inicio=inicio,
        fim=fim,
        uf=args.uf,
        vencimento_ate=args.vencimento_ate,
        somente_vencidos=args.somente_vencidos,
    )
    df = aplicar_regras_vendedor(df)

    arquivo_master, arquivos_vendedor = exportar_excel(df, uf=args.uf)

    print("Relatório automático de cobrança por vendedor")
    print(f"Período faturamento: {inicio} até {fim}")
    print(f"Arquivo master: {arquivo_master}")
    print(f"Linhas: {len(df)}")
    if not df.empty:
        print(f"Vendedores: {df['Responsavel'].nunique()}")
        print(f"Valor total: R$ {df['Valor'].sum():,.2f}")
    print(f"Arquivos por vendedor: {len(arquivos_vendedor)}")


if __name__ == "__main__":
    main()
