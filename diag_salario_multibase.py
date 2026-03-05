import pandas as pd
from sqlalchemy import text
from src.conexao import get_engine

engine = get_engine()

with engine.connect() as conn:
    dbs = pd.read_sql(
        "SELECT name FROM sys.databases WHERE name NOT IN ('master','model','msdb','tempdb') ORDER BY name",
        conn,
    )

bases = dbs['name'].tolist()
print('Bases analisadas:', bases)

achados = []

for db in bases:
    print(f"\n=== VARREDURA {db} ===")
    try:
        q = f"""
        SELECT
            '{db}' AS BANCO,
            c.TABLE_NAME,
            SUM(CASE WHEN c.COLUMN_NAME LIKE '%SALAR%' OR c.COLUMN_NAME LIKE '%LIQUIDO%' OR c.COLUMN_NAME LIKE '%VENCTOS%' THEN 1 ELSE 0 END) AS COLS_SAL,
            SUM(CASE WHEN c.COLUMN_NAME IN ('FUNCIONARIO','CONTRATACAO','CPF','NUMEROFUNC') THEN 1 ELSE 0 END) AS COLS_CHAVE
        FROM [{db}].INFORMATION_SCHEMA.COLUMNS c
        GROUP BY c.TABLE_NAME
        HAVING SUM(CASE WHEN c.COLUMN_NAME LIKE '%SALAR%' OR c.COLUMN_NAME LIKE '%LIQUIDO%' OR c.COLUMN_NAME LIKE '%VENCTOS%' THEN 1 ELSE 0 END) > 0
           AND SUM(CASE WHEN c.COLUMN_NAME IN ('FUNCIONARIO','CONTRATACAO','CPF','NUMEROFUNC') THEN 1 ELSE 0 END) > 0
        ORDER BY c.TABLE_NAME
        """
        df = pd.read_sql(q, engine)
        if df.empty:
            print('(sem objetos com perfil salario+funcionario)')
        else:
            print(df.to_string(index=False))
            achados.append(df)
    except Exception as e:
        print('Erro:', e)

if achados:
    full = pd.concat(achados, ignore_index=True)
    print('\n=== RESUMO GLOBAL ===')
    print(full.to_string(index=False))
else:
    print('\n=== RESUMO GLOBAL ===\nNenhum objeto encontrado fora da lógica já conhecida.')
