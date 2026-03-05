import sys
from pathlib import Path
from datetime import datetime

import pandas as pd

sys.path.insert(0, "src")
from conexao import get_engine
import fase4_dashboard as f4


def main():
    engine = get_engine()
    ano = f4.REF_ANO

    sql = f"""
    SELECT
        p.CODIGO AS Pedido,
        p.DTPEDIDO AS Dt_Pedido,
        f.CIDADE AS Cidade_MG,
        f.UF AS UF,
        SUM(i.VLRTOTAL) AS Faturamento
    FROM VE_PEDIDOITENS i
    JOIN VE_PEDIDO p ON p.CODIGO = i.PEDIDO
    JOIN FN_FORNECEDORES f ON f.CODIGO = p.CLIENTE
    WHERE YEAR(p.DTPEDIDO) = {ano}
      AND f.UF = 'MG'
      AND p.STATUS <> 'C'
      AND i.STATUS <> 'C'
      AND i.FLAGSUB <> 'S'
      AND i.MATERIAL NOT LIKE '8%'
      AND i.TPVENDA NOT IN ({f4.TPVENDA_STR})
      AND p.CODIGO NOT IN (
          SELECT p2.CODIGO FROM VE_PEDIDO p2
          JOIN FN_FORNECEDORES f2 ON f2.CODIGO = p2.CLIENTE
          WHERE f2.UF = 'EX'
      )
    GROUP BY p.CODIGO, p.DTPEDIDO, f.CIDADE, f.UF
    """

    df = pd.read_sql(sql, engine)

    if df.empty:
        print("Sem dados de MG no período.")
        return

    cidades_excluidas = f4.carregar_cidades_excluidas_mg_alexandre()

    df["Cidade_MG_Norm"] = df["Cidade_MG"].map(f4.normalizar_texto)
    df["Responsavel_Aplicado"] = df.apply(
        lambda r: f4.mapear_representante_externo("MG", r["Cidade_MG"], r["Dt_Pedido"]),
        axis=1,
    )
    df["Cidade_Excluida_Alexandre"] = df["Cidade_MG_Norm"].isin(cidades_excluidas).map(lambda x: "SIM" if x else "NAO")

    resumo_cidade = (
        df.groupby(["Cidade_MG", "Cidade_Excluida_Alexandre", "Responsavel_Aplicado"], dropna=False)
          .agg(
              Pedidos=("Pedido", "nunique"),
              Faturamento=("Faturamento", "sum"),
              Primeiro_Pedido=("Dt_Pedido", "min"),
              Ultimo_Pedido=("Dt_Pedido", "max"),
          )
          .reset_index()
          .sort_values(["Cidade_Excluida_Alexandre", "Faturamento"], ascending=[False, False])
    )

    resumo_resp = (
        df.groupby(["Responsavel_Aplicado"], dropna=False)
          .agg(
              Cidades=("Cidade_MG_Norm", "nunique"),
              Pedidos=("Pedido", "nunique"),
              Faturamento=("Faturamento", "sum"),
          )
          .reset_index()
          .sort_values("Faturamento", ascending=False)
    )

    cidades_planilha = pd.DataFrame({"Cidade_Excluida_MG": sorted(cidades_excluidas)})

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = Path("exports") / f"conferencia_territorio_mg_{ts}.xlsx"

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        resumo_resp.to_excel(writer, sheet_name="Resumo_Responsavel", index=False)
        resumo_cidade.to_excel(writer, sheet_name="Resumo_Cidade", index=False)
        df[["Pedido", "Dt_Pedido", "Cidade_MG", "Cidade_Excluida_Alexandre", "Responsavel_Aplicado", "Faturamento"]].to_excel(
            writer,
            sheet_name="Detalhe_Pedido",
            index=False,
        )
        cidades_planilha.to_excel(writer, sheet_name="Cidades_Excluidas_Arquivo", index=False)

    print(f"Arquivo gerado: {out.resolve()}")
    print("Resumo por responsável:")
    print(resumo_resp.to_string(index=False))


if __name__ == "__main__":
    main()
