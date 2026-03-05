"""
relatorio_pagas_estado.py — Pagamentos por Estado (Peça x Bomba)

Objetivo:
- Mostrar todas as parcelas pagas no período, por estado do cliente (UF),
  com valor classificado entre peça e bomba.

Classificação:
- BOMBA: MATERIAL LIKE '8%'
- PEÇA: MATERIAL NOT LIKE '8%'

Critério de valor:
- Usa o valor efetivamente pago da parcela (FN_RECEBER.VLREFETIVO).
- Rateia o valor pago da parcela conforme composição da NF em peças/bombas,
  com base em FN_NFEITENS.VLRTOTAL.

Uso:
    python relatorio_pagas_estado.py
    python relatorio_pagas_estado.py --mes 2 --ano 2026
    python relatorio_pagas_estado.py --incluir-exterior
"""

import argparse
from datetime import date, datetime
from pathlib import Path
import sys

import pandas as pd

sys.path.insert(0, "src")
from conexao import get_engine

OUTPUT = Path("exports")
OUTPUT.mkdir(exist_ok=True)


def default_mes_ano_pagamento() -> tuple[int, int]:
    hoje = date.today()
    if hoje.month == 1:
        return 12, hoje.year - 1
    return hoje.month - 1, hoje.year


def parse_args():
    mes_default, ano_default = default_mes_ano_pagamento()
    p = argparse.ArgumentParser()
    p.add_argument("--mes", type=int, default=mes_default, help="Mês de pagamento")
    p.add_argument("--ano", type=int, default=ano_default, help="Ano de pagamento")
    p.add_argument("--incluir-exterior", action="store_true", help="Inclui UF='EX'")
    return p.parse_args()


def primeiro_dia(mes: int, ano: int) -> str:
    return f"{ano:04d}-{mes:02d}-01"


def primeiro_dia_mes_seguinte(mes: int, ano: int) -> str:
    if mes == 12:
        return f"{ano + 1:04d}-01-01"
    return f"{ano:04d}-{mes + 1:02d}-01"


def query_pagas_estado_tipo(engine, ini: str, fim: str, incluir_exterior: bool) -> pd.DataFrame:
    filtro_ex = "" if incluir_exterior else "AND ISNULL(c.UF,'') <> 'EX'"

    sql = f"""
WITH NfMix AS (
    SELECT
        ni.NFE AS NFS,
        SUM(CASE WHEN ni.MATERIAL LIKE '8%' THEN ISNULL(ni.VLRTOTAL,0) ELSE 0 END) AS Vlr_Bombas_NF,
        SUM(CASE WHEN ni.MATERIAL NOT LIKE '8%' THEN ISNULL(ni.VLRTOTAL,0) ELSE 0 END) AS Vlr_Pecas_NF,
        SUM(ISNULL(ni.VLRTOTAL,0)) AS Vlr_Total_NF
    FROM FN_NFEITENS ni
    GROUP BY ni.NFE
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
Pagas AS (
    SELECT
        r.CODIGO AS Receber,
        r.NFS,
        r.NOTASEQ,
        r.DTPAGAMENTO,
        CASE WHEN r.DTPAGAMENTO IS NULL THEN 'NAO PAGA' ELSE 'PAGA' END AS Status_Pagamento,
        ISNULL(r.VLREFETIVO,0) AS Vlr_Pago_Parcela,
        r.NRODOCUMENTO,
        n.NRONOTA,
        n.DTEMISSAO,
        n.VEPEDIDO,
        n.CLIENTE AS Cod_Cliente,
        c.RAZAO AS Cliente,
        c.UF AS Cliente_UF,
        c.CIDADE AS Cliente_Cidade,
        n.VENDEDOR AS Cod_Vendedor,
        v.RAZAO AS Vendedor_NF,
        COALESCE(vp.Vend_Ext_Principal, vp.Vend_Ext,
                 CASE WHEN v.TIPO = 'E' THEN v.RAZAO END) AS Vendedor_Externo,
        COALESCE(vp.Vend_Int_Principal, vp.Vend_Int,
                 CASE WHEN v.TIPO = 'I' THEN v.RAZAO END) AS Vendedor_Interno
    FROM FN_RECEBER r
    JOIN FN_NFS n ON n.CODIGO = r.NFS
    LEFT JOIN FN_FORNECEDORES c ON c.CODIGO = n.CLIENTE
    LEFT JOIN FN_VENDEDORES v ON v.CODIGO = n.VENDEDOR
    LEFT JOIN VendedoresPedido vp ON vp.PEDIDO = n.VEPEDIDO
    WHERE r.NFS IS NOT NULL
      AND r.DTPAGAMENTO >= '{ini}'
      AND r.DTPAGAMENTO <  '{fim}'
      AND r.DTCANCELAMENTO IS NULL
      AND n.STATUSNF <> 'C'
      {filtro_ex}
)
SELECT
    p.Receber,
    p.NFS,
    p.NRONOTA,
    p.NOTASEQ,
    p.DTPAGAMENTO,
    p.Status_Pagamento,
    p.DTEMISSAO,
    p.Cliente,
    p.Cliente_UF,
    p.Cliente_Cidade,
    p.Vendedor_Externo,
    p.Vendedor_Interno,
    p.Vendedor_NF,
    p.Vlr_Pago_Parcela,
    x.Tipo,
    x.Valor_Tipo,
    nm.Vlr_Pecas_NF,
    nm.Vlr_Bombas_NF,
    nm.Vlr_Total_NF
FROM Pagas p
LEFT JOIN NfMix nm ON nm.NFS = p.NFS
CROSS APPLY (
    VALUES
    ('PECA',
        CASE
            WHEN ISNULL(nm.Vlr_Total_NF,0) > 0
                THEN p.Vlr_Pago_Parcela * (ISNULL(nm.Vlr_Pecas_NF,0) / nm.Vlr_Total_NF)
            ELSE 0
        END
    ),
    ('BOMBA',
        CASE
            WHEN ISNULL(nm.Vlr_Total_NF,0) > 0
                THEN p.Vlr_Pago_Parcela * (ISNULL(nm.Vlr_Bombas_NF,0) / nm.Vlr_Total_NF)
            ELSE 0
        END
    )
) x(Tipo, Valor_Tipo)
WHERE x.Valor_Tipo > 0
ORDER BY p.DTPAGAMENTO, p.Cliente_UF, p.NRONOTA, p.NOTASEQ, x.Tipo
"""

    return pd.read_sql(sql, engine)


def gerar_excel(df: pd.DataFrame, mes: int, ano: int) -> Path:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo = OUTPUT / f"pagas_estado_{mes:02d}-{ano}_{ts}.xlsx"

    if df.empty:
        with pd.ExcelWriter(arquivo, engine="openpyxl") as writer:
            pd.DataFrame([{"Mensagem": "Sem pagamentos no período informado."}]).to_excel(
                writer, sheet_name="Resumo", index=False
            )
        return arquivo

    resumo_estado = (
        df.groupby(["Cliente_UF"], dropna=False)
          .agg(
              Parcelas=("Receber", "nunique"),
              NFs=("NFS", "nunique"),
              Clientes=("Cliente", "nunique"),
              Valor_Pago=("Valor_Tipo", "sum"),
          )
          .reset_index()
          .sort_values("Valor_Pago", ascending=False)
    )

    resumo_estado_tipo = (
        df.groupby(["Cliente_UF", "Tipo"], dropna=False)
          .agg(
              Parcelas=("Receber", "nunique"),
              NFs=("NFS", "nunique"),
              Valor_Pago=("Valor_Tipo", "sum"),
          )
          .reset_index()
          .sort_values(["Cliente_UF", "Valor_Pago"], ascending=[True, False])
    )

    resumo_cidade = (
        df.groupby(["Cliente_UF", "Cliente_Cidade", "Tipo"], dropna=False)
          .agg(
              Parcelas=("Receber", "nunique"),
              NFs=("NFS", "nunique"),
              Valor_Pago=("Valor_Tipo", "sum"),
          )
          .reset_index()
          .sort_values(["Cliente_UF", "Valor_Pago"], ascending=[True, False])
    )

    resumo_vendedor_externo = (
        df.groupby(["Vendedor_Externo", "Tipo"], dropna=False)
          .agg(
              Parcelas=("Receber", "nunique"),
              NFs=("NFS", "nunique"),
              Valor_Pago=("Valor_Tipo", "sum"),
          )
          .reset_index()
          .sort_values("Valor_Pago", ascending=False)
    )

    resumo_vendedor_interno = (
        df.groupby(["Vendedor_Interno", "Tipo"], dropna=False)
          .agg(
              Parcelas=("Receber", "nunique"),
              NFs=("NFS", "nunique"),
              Valor_Pago=("Valor_Tipo", "sum"),
          )
          .reset_index()
          .sort_values("Valor_Pago", ascending=False)
    )

    with pd.ExcelWriter(arquivo, engine="openpyxl") as writer:
        resumo_estado.to_excel(writer, sheet_name="Resumo_Estado", index=False)
        resumo_estado_tipo.to_excel(writer, sheet_name="Resumo_Estado_Tipo", index=False)
        resumo_cidade.to_excel(writer, sheet_name="Resumo_Cidade_Tipo", index=False)
        resumo_vendedor_externo.to_excel(writer, sheet_name="Resumo_Vend_Externo", index=False)
        resumo_vendedor_interno.to_excel(writer, sheet_name="Resumo_Vend_Interno", index=False)
        df.to_excel(writer, sheet_name="Detalhado", index=False)

    return arquivo


def main():
    args = parse_args()
    if args.mes < 1 or args.mes > 12:
        raise ValueError("--mes deve estar entre 1 e 12")

    ini = primeiro_dia(args.mes, args.ano)
    fim = primeiro_dia_mes_seguinte(args.mes, args.ano)

    print(f"Relatório de Pagas por Estado — {args.mes:02d}/{args.ano}")
    print(f"  Período pagamento: {ini} até {fim}")
    print("  Escopo geográfico: Brasil (sem EX)" if not args.incluir_exterior else "  Escopo geográfico: Brasil + exterior")

    engine = get_engine()
    df = query_pagas_estado_tipo(engine, ini, fim, args.incluir_exterior)
    arquivo = gerar_excel(df, args.mes, args.ano)

    print(f"\n✅ Arquivo gerado: {arquivo.resolve()}")
    if not df.empty:
        print(f"   Parcelas pagas: {df['Receber'].nunique()}")
        print(f"   NFs: {df['NFS'].nunique()}")
        print(f"   Valor pago total (rateado): R$ {df['Valor_Tipo'].sum():,.2f}")
    else:
        print("   Sem dados no período informado.")


if __name__ == "__main__":
    main()
