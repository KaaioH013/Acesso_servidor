import pandas as pd

from src.conexao import get_engine


def show_query(engine, titulo: str, sql: str):
    print(f"\n=== {titulo} ===")
    try:
        df = pd.read_sql(sql, engine)
        if df.empty:
            print("(sem registros)")
        else:
            print(df.to_string(index=False))
    except Exception as e:
        print(f"ERRO: {e}")


def main():
    engine = get_engine()

    # 1) Estrutura do desenho / lista de engenharia
    show_query(
        engine,
        "PR_LISTAENG (top 10)",
        """
        SELECT TOP 10 *
        FROM PR_LISTAENG
        ORDER BY CODIGO DESC
        """,
    )

    show_query(
        engine,
        "PR_LISTAENG colunas-chave",
        """
        SELECT COLUMN_NAME, DATA_TYPE
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = 'PR_LISTAENG'
          AND COLUMN_NAME IN ('CODIGO','DESENHO','CODMHP','MATERIAL','POSICAO','DESCRICAO','QTDE')
        ORDER BY ORDINAL_POSITION
        """,
    )

    # 2) Data projetada do roteiro por opção do desenho
    show_query(
        engine,
        "PR_DESENHOROTEIROOPCAO (top 10)",
        """
        SELECT TOP 10 *
        FROM PR_DESENHOROTEIROOPCAO
        ORDER BY CODIGO DESC
        """,
    )

    show_query(
        engine,
        "PR_DESENHOROTEIROOPCAO colunas-chave",
        """
        SELECT COLUMN_NAME, DATA_TYPE
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = 'PR_DESENHOROTEIROOPCAO'
          AND COLUMN_NAME IN ('CODIGO','LISTAENG','DTPROJETADO','PROJETADO','REVISAO')
        ORDER BY ORDINAL_POSITION
        """,
    )

    # 3) Tabelas de compra candidatas vistas na tela do material
    show_query(
        engine,
        "Tabelas compra candidatas",
        """
        SELECT TABLE_NAME
        FROM INFORMATION_SCHEMA.TABLES
        WHERE TABLE_TYPE = 'BASE TABLE'
          AND (TABLE_NAME LIKE 'CP%PED%' OR TABLE_NAME LIKE 'FN%PED%' OR TABLE_NAME LIKE 'MT%PED%' OR TABLE_NAME LIKE 'FN%COMP%')
        ORDER BY TABLE_NAME
        """,
    )

    for t in ["FN_PEDIDO", "FN_PEDIDOITEM", "CP_PEDIDO", "CP_PEDIDOITEM", "FN_FORNECEDORES"]:
        show_query(
            engine,
            f"{t} colunas (se existir)",
            f"""
            SELECT COLUMN_NAME, DATA_TYPE
            FROM INFORMATION_SCHEMA.COLUMNS
            WHERE TABLE_NAME = '{t}'
            ORDER BY ORDINAL_POSITION
            """,
        )

    show_query(
        engine,
        "Tabelas com coluna PEDIDO",
        """
        SELECT TABLE_NAME
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE COLUMN_NAME = 'PEDIDO'
        GROUP BY TABLE_NAME
        ORDER BY TABLE_NAME
        """,
    )

    show_query(
        engine,
        "Tabelas com coluna DTPEDIDO",
        """
        SELECT TABLE_NAME
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE COLUMN_NAME = 'DTPEDIDO'
        GROUP BY TABLE_NAME
        ORDER BY TABLE_NAME
        """,
    )

    show_query(
        engine,
        "Tabelas com STATUS + FORNECEDOR",
        """
        SELECT c1.TABLE_NAME
        FROM INFORMATION_SCHEMA.COLUMNS c1
        JOIN INFORMATION_SCHEMA.COLUMNS c2
          ON c2.TABLE_NAME = c1.TABLE_NAME AND c2.COLUMN_NAME = 'STATUS'
        JOIN INFORMATION_SCHEMA.COLUMNS c3
          ON c3.TABLE_NAME = c1.TABLE_NAME AND c3.COLUMN_NAME = 'FORNECEDOR'
        WHERE c1.COLUMN_NAME = 'PEDIDO'
        GROUP BY c1.TABLE_NAME
        ORDER BY c1.TABLE_NAME
        """,
    )

    for t in ["CO_PEDIDOS", "CO_PEDIDOITENS", "IT_PEDIDO", "IT_PEDIDOITENS", "VE_PEDIDO", "VE_PEDIDOITENS"]:
        show_query(
            engine,
            f"{t} colunas (se existir)",
            f"""
            SELECT COLUMN_NAME, DATA_TYPE
            FROM INFORMATION_SCHEMA.COLUMNS
            WHERE TABLE_NAME = '{t}'
            ORDER BY ORDINAL_POSITION
            """,
        )

    show_query(
        engine,
        "Amostra compras por material (CO_*)",
        """
        SELECT TOP 10
            i.PEDIDO,
            p.DTPEDIDO,
            p.FORNECEDOR,
            p.STATUS AS ST_PED,
            i.STATUS AS ST_ITEM,
            i.MATERIAL,
            i.QTDEPED,
            i.UNDSUP,
            i.PRECIPI,
            i.VLRUNIT
        FROM CO_PEDIDOITENS i
        JOIN CO_PEDIDOS p ON p.CODIGO = i.PEDIDO
        ORDER BY p.DTPEDIDO DESC
        """,
    )


if __name__ == "__main__":
    main()
