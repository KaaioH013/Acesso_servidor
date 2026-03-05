import pandas as pd

from src.conexao import get_engine


def main():
    engine = get_engine()
    for tabela in ["MT_ESTOQUE", "VE_PEDIDOITENS", "VE_PEDIDO", "MT_MOVIMENTACAO"]:
        print(f"\n=== {tabela} ===")
        df = pd.read_sql(
            f"""
            SELECT COLUMN_NAME, DATA_TYPE
            FROM INFORMATION_SCHEMA.COLUMNS
            WHERE TABLE_NAME = '{tabela}'
            ORDER BY ORDINAL_POSITION
            """,
            engine,
        )
        print(df.to_string(index=False))


if __name__ == "__main__":
    main()
