import pandas as pd
from src.conexao import get_engine

engine = get_engine()

q_tabelas = """
SELECT TABLE_NAME
FROM INFORMATION_SCHEMA.TABLES
WHERE TABLE_TYPE = 'BASE TABLE'
  AND (
    TABLE_NAME LIKE 'PR_%'
    OR TABLE_NAME LIKE '%ROTEIR%'
    OR TABLE_NAME LIKE '%FICHA%'
    OR TABLE_NAME LIKE '%ESTRUT%'
    OR TABLE_NAME LIKE '%MATERIA%'
    OR TABLE_NAME LIKE '%MP%'
  )
ORDER BY TABLE_NAME
"""

q_colunas = """
SELECT TABLE_NAME, COLUMN_NAME, DATA_TYPE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE (
    TABLE_NAME LIKE 'PR_%'
    OR TABLE_NAME LIKE '%ROTEIR%'
    OR TABLE_NAME LIKE '%FICHA%'
    OR TABLE_NAME LIKE '%ESTRUT%'
    OR TABLE_NAME LIKE '%MATERIA%'
    OR TABLE_NAME LIKE '%MP%'
)
AND (
    COLUMN_NAME LIKE '%ROTEIR%'
    OR COLUMN_NAME LIKE 'DT%'
    OR COLUMN_NAME LIKE '%MATERIAL%'
    OR COLUMN_NAME LIKE '%MATERIA%'
    OR COLUMN_NAME LIKE '%COMPRA%'
    OR COLUMN_NAME LIKE '%DATA%'
)
ORDER BY TABLE_NAME, COLUMN_NAME
"""

print('=== TABELAS CANDIDATAS ===')
df_t = pd.read_sql(q_tabelas, engine)
print(df_t.to_string(index=False))

print('\n=== COLUNAS CANDIDATAS ===')
df_c = pd.read_sql(q_colunas, engine)
print(df_c.to_string(index=False))
