import argparse
from datetime import date, datetime
from pathlib import Path

import pandas as pd

from src.conexao import get_engine

OUTPUT = Path("exports")
OUTPUT.mkdir(exist_ok=True)


def parse_args():
    hoje = date.today()
    dt_ini_default = f"{hoje.year}-{hoje.month:02d}-01"
    dt_fim_default = hoje.strftime("%Y-%m-%d")

    p = argparse.ArgumentParser(description="Replica resumo 528 (Saidas x Devolucoes)")
    p.add_argument("--dt-ini", type=str, default=dt_ini_default, help="Data inicial emissao (YYYY-MM-DD)")
    p.add_argument("--dt-fim", type=str, default=dt_fim_default, help="Data final emissao (YYYY-MM-DD)")
    p.add_argument("--usuario", type=int, default=124, help="Codigo do usuario no GR_USUARIOS")
    p.add_argument(
        "--formato",
        type=str,
        default="xlsx",
        choices=["xlsx", "csv", "ambos"],
        help="Formato de saida: xlsx, csv ou ambos",
    )
    return p.parse_args()


def query_detalhe(engine, dt_ini: str, dt_fim: str, usuario: int) -> pd.DataFrame:
    sql = f"""
WITH
FS AS (
    SELECT VALOR FROM GR_FILTROITENSUSUARIOCONTIDO WHERE FILTROITEM = 6958 AND USUARIO = {usuario}
),
FE AS (
    SELECT VALOR FROM GR_FILTROITENSUSUARIOCONTIDO WHERE FILTROITEM = 6966 AND USUARIO = {usuario}
),
FT AS (
    SELECT VALOR FROM GR_FILTROITENSUSUARIOCONTIDO WHERE FILTROITEM = 6959 AND USUARIO = {usuario}
),
FU AS (
    SELECT VALOR FROM GR_FILTROITENSUSUARIOCONTIDO WHERE FILTROITEM = 6962 AND USUARIO = {usuario}
),
FM AS (
    SELECT VALOR FROM GR_FILTROITENSUSUARIOCONTIDO WHERE FILTROITEM = 6963 AND USUARIO = {usuario}
),
FST AS (
    SELECT VALOR FROM GR_FILTROITENSUSUARIOCONTIDO WHERE FILTROITEM = 6965 AND USUARIO = {usuario}
),
VendedoresPedido AS (
    SELECT
        pv.PEDIDO,
        MAX(CASE WHEN fv.TIPO = 'E' AND pv.CKPRINCIPAL = 'S' THEN fv.RAZAO END) AS Vend_Ext_Principal,
        MAX(CASE WHEN fv.TIPO = 'E' THEN fv.RAZAO END) AS Vend_Ext,
        MAX(CASE WHEN fv.TIPO = 'I' AND pv.CKPRINCIPAL = 'S' THEN fv.RAZAO END) AS Vend_Int_Principal,
        MAX(CASE WHEN fv.TIPO = 'I' THEN fv.RAZAO END) AS Vend_Int
    FROM VE_PEDIDOVENDEDOR pv
    JOIN FN_VENDEDORES fv ON fv.CODIGO = pv.VENDEDOR
    GROUP BY pv.PEDIDO
),
PedidosNFS AS (
    SELECT DISTINCT n.CODIGO AS NFS, n.VEPEDIDO AS PEDIDO
    FROM FN_NFS n
    WHERE n.VEPEDIDO IS NOT NULL

    UNION

    SELECT DISTINCT x.NFS, x.VEPEDIDO AS PEDIDO
    FROM FN_NFSADTPEDIDO x
    WHERE x.NFS IS NOT NULL AND x.VEPEDIDO IS NOT NULL

    UNION

    SELECT DISTINCT i.NFS, pi.PEDIDO
    FROM FN_NFSITENS i
    JOIN VE_PEDIDOITENS pi ON pi.CODIGO = i.PEDIDOITEM
    WHERE i.NFS IS NOT NULL AND i.PEDIDOITEM IS NOT NULL AND pi.PEDIDO IS NOT NULL
),
VendedoresNFS AS (
    SELECT
        pn.NFS,
        MAX(vp.Vend_Ext_Principal) AS Vend_Ext_Principal,
        MAX(vp.Vend_Ext) AS Vend_Ext,
        MAX(vp.Vend_Int_Principal) AS Vend_Int_Principal,
        MAX(vp.Vend_Int) AS Vend_Int
    FROM PedidosNFS pn
    JOIN VendedoresPedido vp ON vp.PEDIDO = pn.PEDIDO
    GROUP BY pn.NFS
),
Saidas AS (
    SELECT
        c.RAZAO AS Cliente,
        ISNULL(c.UF, '') AS UF,
        i.CFOP,
        n.NRONOTA AS Nro_Nota,
        n.DTEMISSAO AS Dt_Emissao,
        n.STATUSNF AS Status,
        ISNULL(i.VLRICMS, 0) AS Vlr_ICMS,
        ISNULL(i.VLRIPI, 0) AS Vlr_IPI,
        ISNULL(i.VLRIPIDEVOLVIDO, 0) AS Vlr_IPI_Devolvido,
        ISNULL(i.VLRICMSSUB, 0) AS Vlr_ICMS_Subst,
        ISNULL(i.VLRPIS, 0) AS Vlr_PIS,
        ISNULL(i.VLRCOFINS, 0) AS Vlr_COFINS,
        ISNULL(i.VLRTOTAL, 0) AS Vlr_Total_Nota,
        r.DTVENCIMENTO AS Dt_Vencimento,
        COALESCE(vnfs.Vend_Int_Principal, vnfs.Vend_Int) AS Vendedor_Interno,
        COALESCE(vnfs.Vend_Ext_Principal, vnfs.Vend_Ext) AS Vendedor_Externo,
        'S' AS Tipo
    FROM FN_NFS n
    JOIN FN_NFSITENS i ON i.NFS = n.CODIGO
    LEFT JOIN FN_FORNECEDORES c ON c.CODIGO = n.CLIENTE
    OUTER APPLY (
        SELECT MAX(rr.DTVENCIMENTO) AS DTVENCIMENTO
        FROM FN_RECEBER rr
        WHERE rr.NFS = n.CODIGO AND rr.DTCANCELAMENTO IS NULL
    ) r
    LEFT JOIN VendedoresNFS vnfs ON vnfs.NFS = n.CODIGO
    JOIN FT ON FT.VALOR = n.TPDOCUMENTO
    JOIN FST ON FST.VALOR = n.STATUSNF
    JOIN FU ON FU.VALOR = ISNULL(c.UF, '')
    JOIN FM ON FM.VALOR = i.MATERIAL
    JOIN FS ON FS.VALOR = i.CFOP
    WHERE n.DTEMISSAO >= '{dt_ini}'
      AND n.DTEMISSAO < DATEADD(DAY, 1, '{dt_fim}')
),
Devolucoes AS (
    SELECT
        c.RAZAO AS Cliente,
        ISNULL(c.UF, '') AS UF,
        i.CFOP,
        n.NRODOCUMENTO AS Nro_Nota,
        n.DTLANCAMENTO AS Dt_Emissao,
        n.STATUSNF AS Status,
        ISNULL(i.VLRICMS, 0) AS Vlr_ICMS,
        ISNULL(i.VLRIPI, 0) AS Vlr_IPI,
        ISNULL(i.VLRIPIDEVOLVIDO, 0) AS Vlr_IPI_Devolvido,
        ISNULL(i.VLRICMSSUB, 0) AS Vlr_ICMS_Subst,
        ISNULL(i.VLRPIS, 0) AS Vlr_PIS,
        ISNULL(i.VLRCOFINS, 0) AS Vlr_COFINS,
        ISNULL(i.VLRTOTAL, 0) AS Vlr_Total_Nota,
        CAST(NULL AS DATETIME) AS Dt_Vencimento,
        CAST(NULL AS VARCHAR(120)) AS Vendedor_Interno,
        CAST(NULL AS VARCHAR(120)) AS Vendedor_Externo,
        'D' AS Tipo
    FROM FN_NFE n
    JOIN FN_NFEITENS i ON i.NFE = n.CODIGO
    LEFT JOIN FN_FORNECEDORES c ON c.CODIGO = n.FORNECEDOR
    JOIN FT ON FT.VALOR = n.TPDOCUMENTO
    JOIN FST ON FST.VALOR = n.STATUSNF
    JOIN FU ON FU.VALOR = ISNULL(c.UF, '')
    JOIN FM ON FM.VALOR = i.MATERIAL
    JOIN FE ON FE.VALOR = i.CFOP
        WHERE n.DTENTRADA >= '{dt_ini}'
            AND n.DTENTRADA < DATEADD(DAY, 1, '{dt_fim}')
)
SELECT * FROM Saidas
UNION ALL
SELECT * FROM Devolucoes
ORDER BY Dt_Emissao, Nro_Nota, CFOP
"""
    return pd.read_sql(sql, engine)


def montar_resumo(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame([
            {"Tipo": "S", "Vlr_Total_Nota": 0.0},
            {"Tipo": "D", "Vlr_Total_Nota": 0.0},
            {"Tipo": "TOTAL", "Vlr_Total_Nota": 0.0},
        ])

    resumo = (
        df.groupby("Tipo", dropna=False, as_index=False)
        .agg(Vlr_Total_Nota=("Vlr_Total_Nota", "sum"))
    )

    for t in ["S", "D"]:
        if t not in resumo["Tipo"].astype(str).tolist():
            resumo = pd.concat([resumo, pd.DataFrame([{"Tipo": t, "Vlr_Total_Nota": 0.0}])], ignore_index=True)

    resumo = resumo[resumo["Tipo"].isin(["S", "D"])].copy()
    saidas = float(resumo.loc[resumo["Tipo"] == "S", "Vlr_Total_Nota"].sum())
    devol = float(resumo.loc[resumo["Tipo"] == "D", "Vlr_Total_Nota"].sum())
    total = saidas - devol

    resumo = pd.concat(
        [
            resumo,
            pd.DataFrame(
                [
                    {"Tipo": "TOTAL_LIQ", "Vlr_Total_Nota": total},
                ]
            ),
        ],
        ignore_index=True,
    )
    return resumo


def salvar_excel(df: pd.DataFrame, resumo: pd.DataFrame, dt_ini: str, dt_fim: str) -> Path:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = OUTPUT / f"relatorio_528_replicado_{dt_ini}_{dt_fim}_{ts}.xlsx"

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Detalhe_528", index=False)
        resumo.to_excel(writer, sheet_name="Resumo_528", index=False)

    return out


def _formatar_layout_528(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out = out.rename(
        columns={
            "Nro_Nota": "Nro. Nota",
            "Dt_Emissao": "Dt. Emissao",
            "Vlr_ICMS": "Vlr. ICMS",
            "Vlr_IPI": "Vlr. IPI",
            "Vlr_IPI_Devolvido": "Vlr. IPI Devolvido",
            "Vlr_ICMS_Subst": "Vlr. ICMS Subst.",
            "Vlr_PIS": "Vlr. PIS",
            "Vlr_COFINS": "Vlr. COFINS",
            "Vlr_Total_Nota": "Vlr. Total Nota",
            "Dt_Vencimento": "Dt. Vencimento",
            "Vendedor_Interno": "Vendedor Interno",
            "Vendedor_Externo": "Vendedor Externo",
        }
    )

    cols = [
        "Cliente",
        "UF",
        "CFOP",
        "Nro. Nota",
        "Dt. Emissao",
        "Status",
        "Vlr. ICMS",
        "Vlr. IPI",
        "Vlr. IPI Devolvido",
        "Vlr. ICMS Subst.",
        "Vlr. PIS",
        "Vlr. COFINS",
        "Vlr. Total Nota",
        "Dt. Vencimento",
        "Vendedor Interno",
        "Vendedor Externo",
        "Tipo",
    ]
    return out[cols]


def salvar_csv(df: pd.DataFrame, dt_ini: str, dt_fim: str) -> Path:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = OUTPUT / f"relatorio_528_replicado_{dt_ini}_{dt_fim}_{ts}.csv"
    layout = _formatar_layout_528(df)
    layout.to_csv(out, sep=";", index=False, encoding="utf-8-sig")
    return out


def main():
    args = parse_args()
    engine = get_engine()

    df = query_detalhe(engine, args.dt_ini, args.dt_fim, args.usuario)
    resumo = montar_resumo(df)
    out_xlsx = None
    out_csv = None
    if args.formato in ("xlsx", "ambos"):
        out_xlsx = salvar_excel(df, resumo, args.dt_ini, args.dt_fim)
    if args.formato in ("csv", "ambos"):
        out_csv = salvar_csv(df, args.dt_ini, args.dt_fim)

    saidas = float(resumo.loc[resumo["Tipo"] == "S", "Vlr_Total_Nota"].sum())
    devol = float(resumo.loc[resumo["Tipo"] == "D", "Vlr_Total_Nota"].sum())
    total = float(resumo.loc[resumo["Tipo"] == "TOTAL_LIQ", "Vlr_Total_Nota"].sum())

    print("Relatorio 528 replicado gerado")
    if out_xlsx:
        print(f"Arquivo XLSX: {out_xlsx}")
    if out_csv:
        print(f"Arquivo CSV: {out_csv}")
    print(f"Saidas: R$ {saidas:,.2f}")
    print(f"Devolucoes: R$ {devol:,.2f}")
    print(f"Total liquido: R$ {total:,.2f}")
    print(f"Linhas: {len(df)}")


if __name__ == "__main__":
    main()
