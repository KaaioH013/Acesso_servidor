import pandas as pd
from src.conexao import get_engine

engine = get_engine()

for t in ["CO_PEDIDOS", "CO_PEDIDOITENS"]:
    print(f"\n=== {t} ===")
    df = pd.read_sql(
        f"""
        SELECT COLUMN_NAME, DATA_TYPE
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = '{t}'
        ORDER BY ORDINAL_POSITION
        """,
        engine,
    )
    print(df.to_string(index=False))

print("\n=== AMOSTRA JOIN ===")
df2 = pd.read_sql(
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
        i.VLRUNIT,
        i.VLRDESCONTO,
        i.VLRST,
        i.VLRTOTAL
    FROM CO_PEDIDOITENS i
    JOIN CO_PEDIDOS p ON p.CODIGO = i.PEDIDO
    ORDER BY p.DTPEDIDO DESC
    """,
    engine,
)
print(df2.to_string(index=False))
