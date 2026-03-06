import pandas as pd
from src.conexao import get_engine

eng = get_engine()

q_itens = """
SELECT
    COUNT(DISTINCT n.CODIGO) AS NFS_COM_PECA,
    SUM(ISNULL(i.VLRTOTAL, 0)) AS FAT_ITENS_PECA
FROM FN_NFS n
JOIN FN_NFSITENS i ON i.NFS = n.CODIGO
LEFT JOIN FN_FORNECEDORES c ON c.CODIGO = n.CLIENTE
WHERE n.DTEMISSAO >= '2026-01-01'
  AND n.DTEMISSAO < '2026-02-01'
  AND n.STATUSNF <> 'C'
  AND ISNULL(c.UF, '') <> 'EX'
  AND i.MATERIAL NOT LIKE '8%'
"""

q_notas = """
WITH T AS (
    SELECT
        n.CODIGO AS NFS,
        SUM(CASE WHEN i.MATERIAL LIKE '8%' THEN 1 ELSE 0 END) AS IT_BOMBA,
        SUM(CASE WHEN i.MATERIAL NOT LIKE '8%' THEN 1 ELSE 0 END) AS IT_PECA,
        ISNULL(n.VLRTOTAL, 0) AS VLRTOTAL
    FROM FN_NFS n
    JOIN FN_NFSITENS i ON i.NFS = n.CODIGO
    LEFT JOIN FN_FORNECEDORES c ON c.CODIGO = n.CLIENTE
    WHERE n.DTEMISSAO >= '2026-01-01'
      AND n.DTEMISSAO < '2026-02-01'
      AND n.STATUSNF <> 'C'
      AND ISNULL(c.UF, '') <> 'EX'
    GROUP BY n.CODIGO, n.VLRTOTAL
)
SELECT
    COUNT(*) AS NFS_SOMENTE_PECA,
    SUM(VLRTOTAL) AS FAT_NFS_SOMENTE_PECA
FROM T
WHERE IT_PECA > 0
  AND IT_BOMBA = 0
"""

a = pd.read_sql(q_itens, eng)
b = pd.read_sql(q_notas, eng)

fat_itens = float(a.loc[0, "FAT_ITENS_PECA"] or 0)
nfs_itens = int(a.loc[0, "NFS_COM_PECA"] or 0)
fat_notas = float(b.loc[0, "FAT_NFS_SOMENTE_PECA"] or 0)
nfs_notas = int(b.loc[0, "NFS_SOMENTE_PECA"] or 0)

print("JANEIRO 2026 - SO PECA (SEM EXPORTACAO)")
print(f"FAT_ITENS_PECA= R$ {fat_itens:,.2f} | NFS_COM_PECA= {nfs_itens}")
print(f"FAT_NFS_SOMENTE_PECA= R$ {fat_notas:,.2f} | NFS_SOMENTE_PECA= {nfs_notas}")
