import pandas as pd
from src.conexao import get_engine

engine = get_engine()

sql = """
SELECT t.name AS tabela, c.name AS coluna
FROM sys.tables t
JOIN sys.columns c ON c.object_id = t.object_id
WHERE t.is_ms_shipped = 0
  AND (
    t.name LIKE '%NFS%'
    OR t.name LIKE '%PEDIDO%'
    OR t.name LIKE '%VINC%'
    OR t.name LIKE '%FAT%'
  )
ORDER BY t.name, c.column_id
"""

cols = pd.read_sql(sql, engine)

# Candidate tables that have both NF-ish and pedido-ish columns.
nf_tokens = ('NFS', 'NF', 'NOTA', 'NRONOTA')
ped_tokens = ('PEDIDO', 'VEPEDIDO', 'ORCAMENTO')

cand = []
for t, g in cols.groupby('tabela'):
    names = [str(x).upper() for x in g['coluna'].tolist()]
    has_nf = any(any(tok in n for tok in nf_tokens) for n in names)
    has_ped = any(any(tok in n for tok in ped_tokens) for n in names)
    if has_nf and has_ped:
        cand.append((t, ', '.join(g['coluna'].astype(str).tolist())))

out = pd.DataFrame(cand, columns=['tabela', 'colunas'])
out.to_csv('exports/diag_nfs_pedido_tables.csv', index=False, encoding='utf-8-sig')
print('exports/diag_nfs_pedido_tables.csv')
