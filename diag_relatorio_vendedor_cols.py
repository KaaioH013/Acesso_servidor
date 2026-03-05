import pandas as pd
from src.conexao import get_engine

engine = get_engine()

alvos = ["FN_RECEBER", "FN_NFS", "VE_PEDIDO", "VE_ORCAMENTOS", "FN_FORNECEDORES"]

for t in alvos:
    print(f"\n=== {t} COLUNAS ===")
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

print("\n=== AMOSTRA FN_RECEBER + FN_NFS + VE_PEDIDO ===")
q = """
SELECT TOP 20
    c.RAZAO AS Cliente,
    n.NRONOTA AS NF,
    p.PEDIDOCLIENTE,
    n.VEPEDIDO AS PV,
    p.ORCAMENTO,
    n.DTEMISSAO AS Data_Faturamento,
    r.DTVENCIMENTO AS Vencimento,
    r.VLRPARCELA,
    c.UF,
    c.CIDADE,
    r.DTPAGAMENTO,
    r.DTCANCELAMENTO
FROM FN_RECEBER r
JOIN FN_NFS n ON n.CODIGO = r.NFS
LEFT JOIN VE_PEDIDO p ON p.CODIGO = n.VEPEDIDO
LEFT JOIN FN_FORNECEDORES c ON c.CODIGO = n.CLIENTE
WHERE r.NFS IS NOT NULL
ORDER BY r.CODIGO DESC
"""
print(pd.read_sql(q, engine).to_string(index=False))
