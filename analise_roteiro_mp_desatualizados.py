from datetime import datetime
from pathlib import Path
import argparse

import pandas as pd

from src.conexao import get_engine

OUTPUT = Path("exports")
OUTPUT.mkdir(exist_ok=True)


def obter_roteiro_desatualizado(engine, top_n: int) -> pd.DataFrame:
    top_n = int(top_n)
    top_clause = "" if top_n <= 0 else f"TOP {top_n}"
    sql = """
    WITH RoteiroDesenho AS (
        SELECT
            op.CODIGO AS OP,
            MAX(dro.DTPROJETADO) AS Dt_Roteiro_Ref
        FROM PR_OP op
        JOIN PR_DESENHO d ON d.DESENHO = op.DESENHO
        JOIN PR_LISTAENG le ON le.CODDESENHO = d.CODIGO AND le.POSICAO = '00'
        JOIN PR_DESENHOROTEIROOPCAO dro ON dro.LISTAENG = le.CODIGO AND dro.CKPADRAO = 'S'
        WHERE op.OPSTATUS = 2
          AND dro.DTPROJETADO IS NOT NULL
        GROUP BY op.CODIGO
    )
    SELECT {top_clause}
        op.CODIGO AS OP,
        op.NROOP,
        op.DESENHO AS Desenho_OP,
        op.DESCRICAO AS Descricao_OP,
        s.DESCRICAO AS Status_OP,
        op.DTCADASTRO AS Dt_Cadastro_OP,
        op.DTINICIO AS Dt_Inicio_OP,
        rd.Dt_Roteiro_Ref,
        DATEDIFF(DAY, rd.Dt_Roteiro_Ref, GETDATE()) AS Dias_Desde_Roteiro
    FROM PR_OP op
    LEFT JOIN RoteiroDesenho rd ON rd.OP = op.CODIGO
    LEFT JOIN PR_OPSTATUS s ON s.CODIGO = op.OPSTATUS
    WHERE op.OPSTATUS = 2
      AND rd.Dt_Roteiro_Ref IS NOT NULL
      AND rd.Dt_Roteiro_Ref < '2025-01-01'
    ORDER BY rd.Dt_Roteiro_Ref ASC, op.DTCADASTRO ASC
    """.format(top_clause=top_clause)
    return pd.read_sql(sql, engine)


def obter_mp_sem_compra_6m(engine, top_n: int) -> pd.DataFrame:
    top_n = int(top_n)
    top_clause = "" if top_n <= 0 else f"TOP {top_n}"
    sql = """
    WITH MP_EmProducao AS (
        SELECT DISTINCT
            op.CODIGO AS OP,
            op.NROOP,
            op.DESENHO AS Desenho_OP,
            op.DESCRICAO AS Descricao_OP,
                        CAST(le.CODMP AS varchar(30)) AS Materia_Prima,
            le.DESCRICAO AS Desc_MP,
            op.DTCADASTRO,
            op.DTINICIO
        FROM PR_OP op
                JOIN PR_DESENHO d ON d.DESENHO = op.DESENHO
                JOIN PR_LISTAENG le ON le.CODDESENHO = d.CODIGO
        WHERE op.OPSTATUS = 2
                    AND le.POSICAO <> '00'
                    AND le.CODMP IS NOT NULL
    ),
    UltCompra AS (
        SELECT
            CAST(i.MATERIAL AS varchar(30)) AS MATERIAL,
            MAX(p.DTPEDIDO) AS Dt_Ultima_Compra
        FROM CO_PEDIDOITENS i
        JOIN CO_PEDIDOS p ON p.CODIGO = i.PEDIDO
                WHERE p.STATUS = 'E'
                    AND i.STATUS = 'E'
        GROUP BY CAST(i.MATERIAL AS varchar(30))
    )
    SELECT {top_clause}
        p.OP,
        p.NROOP,
        p.Desenho_OP,
        p.Descricao_OP,
        p.Materia_Prima,
        p.Desc_MP,
        uc.Dt_Ultima_Compra,
        CASE
            WHEN uc.Dt_Ultima_Compra IS NULL THEN NULL
            ELSE DATEDIFF(DAY, uc.Dt_Ultima_Compra, GETDATE())
        END AS Dias_Sem_Compra
    FROM MP_EmProducao p
    LEFT JOIN UltCompra uc ON uc.MATERIAL = p.Materia_Prima
    WHERE uc.Dt_Ultima_Compra IS NULL
       OR uc.Dt_Ultima_Compra < DATEADD(MONTH, -6, CAST(GETDATE() AS date))
    ORDER BY
        CASE WHEN uc.Dt_Ultima_Compra IS NULL THEN 0 ELSE 1 END,
        uc.Dt_Ultima_Compra ASC
    """.format(top_clause=top_clause)
    return pd.read_sql(sql, engine)


def main():
    parser = argparse.ArgumentParser(description="Analisa roteiro e matéria-prima sem compra")
    parser.add_argument("--top", type=int, default=5, help="Quantidade máxima de registros por lista (<=0 = sem limite)")
    args = parser.parse_args()

    engine = get_engine()

    df_roteiro = obter_roteiro_desatualizado(engine, args.top)
    df_mp = obter_mp_sem_compra_6m(engine, args.top)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    arq = OUTPUT / f"analise_roteiro_mp_{ts}.xlsx"

    with pd.ExcelWriter(arq, engine="openpyxl") as writer:
        df_roteiro.to_excel(writer, index=False, sheet_name="Roteiro_<2025")
        df_mp.to_excel(writer, index=False, sheet_name="MP_sem_compra_6m")

    print("Análise concluída")
    print(f"Arquivo: {arq}")
    print(f"Roteiro desatualizado (top5): {len(df_roteiro)}")
    print(f"MP sem compra >= 6 meses (top5): {len(df_mp)}")

    if not df_roteiro.empty:
        print("\nTop roteiro:")
        print(df_roteiro.to_string(index=False))

    if not df_mp.empty:
        print("\nTop MP sem compra:")
        print(df_mp.to_string(index=False))


if __name__ == "__main__":
    main()
