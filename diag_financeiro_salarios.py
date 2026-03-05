import pandas as pd

from src.conexao import get_engine


def main():
    engine = get_engine()

    q_tabelas = """
    SELECT TABLE_NAME
    FROM INFORMATION_SCHEMA.TABLES
    WHERE TABLE_TYPE = 'BASE TABLE'
      AND (
        TABLE_NAME LIKE 'RH%'
        OR TABLE_NAME LIKE '%FOLHA%'
        OR TABLE_NAME LIKE '%SALAR%'
        OR TABLE_NAME LIKE '%PAGTO%'
        OR TABLE_NAME LIKE '%PAGAMENTO%'
      )
    ORDER BY TABLE_NAME
    """

    q_colunas = """
    SELECT TABLE_NAME, COLUMN_NAME, DATA_TYPE
    FROM INFORMATION_SCHEMA.COLUMNS
    WHERE (
        TABLE_NAME LIKE 'RH%'
        OR TABLE_NAME LIKE '%FOLHA%'
        OR TABLE_NAME LIKE '%SALAR%'
        OR TABLE_NAME LIKE '%PAGTO%'
        OR TABLE_NAME LIKE '%PAGAMENTO%'
    )
    AND (
        COLUMN_NAME LIKE '%SALAR%'
        OR COLUMN_NAME LIKE '%LIQ%'
        OR COLUMN_NAME LIKE '%BRUT%'
        OR COLUMN_NAME LIKE '%PAG%'
        OR COLUMN_NAME LIKE '%VENC%'
        OR COLUMN_NAME LIKE '%COMPET%'
        OR COLUMN_NAME LIKE '%FUNC%'
        OR COLUMN_NAME LIKE '%CPF%'
    )
    ORDER BY TABLE_NAME, COLUMN_NAME
    """

    tbs = pd.read_sql(q_tabelas, engine)
    cols = pd.read_sql(q_colunas, engine)

    print("=== TABELAS CANDIDATAS (RH/FOLHA/SALARIO/PAGAMENTO) ===")
    print(tbs.to_string(index=False) if not tbs.empty else "(nenhuma)")

    print("\n=== COLUNAS CANDIDATAS ===")
    print(cols.to_string(index=False) if not cols.empty else "(nenhuma)")

    if not tbs.empty:
        print("\n=== AMOSTRA DE ESTRUTURA (TOP 8 tabelas) ===")
        for tabela in tbs["TABLE_NAME"].head(8):
            print(f"\n-- {tabela}")
            try:
                c = pd.read_sql(
                    f"""
                    SELECT COUNT(*) AS QTD
                    FROM {tabela}
                    """,
                    engine,
                )
                print(f"Registros: {int(c.iloc[0, 0])}")
            except Exception as e:
                print(f"Sem acesso/erro em count: {e}")

            try:
                s = pd.read_sql(
                    f"""
                    SELECT TOP 3 *
                    FROM {tabela}
                    """,
                    engine,
                )
                print("Colunas:", ", ".join(s.columns.tolist()))
            except Exception as e:
                print(f"Sem acesso/erro em select: {e}")


if __name__ == "__main__":
    main()
