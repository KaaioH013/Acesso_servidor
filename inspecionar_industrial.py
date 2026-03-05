"""
inspecionar_industrial.py — Exibe resumo e tabelas do banco INDUSTRIAL.
"""
import os, sys
sys.path.insert(0, ".")
from dotenv import load_dotenv
load_dotenv()
import pyodbc

SERVER   = os.getenv("DB_SERVER")
PORT     = os.getenv("DB_PORT", "1433")
DATABASE = os.getenv("DB_DATABASE")
USERNAME = os.getenv("DB_USERNAME")
PASSWORD = os.getenv("DB_PASSWORD")
DRIVER   = os.getenv("DB_DRIVER", "SQL Server")

pwd_escaped = "{" + PASSWORD.replace("}", "}}") + "}"
conn_str = (
    f"DRIVER={{{DRIVER}}};SERVER={SERVER},{PORT};"
    f"DATABASE={DATABASE};UID={USERNAME};PWD={pwd_escaped};"
    "TrustServerCertificate=Yes;"
)

conn = pyodbc.connect(conn_str, timeout=10)
cursor = conn.cursor()

# ── Resumo de objetos ────────────────────────────────────────
print(f"\n{'='*55}")
print(f"  Banco: {DATABASE}  |  Servidor: {SERVER}")
print(f"{'='*55}")
print("\n[RESUMO DE OBJETOS]\n")
cursor.execute("""
    SELECT type_desc, COUNT(*)
    FROM sys.objects
    WHERE is_ms_shipped = 0
    GROUP BY type_desc
    ORDER BY 2 DESC
""")
for r in cursor.fetchall():
    print(f"  {r[0]:<35} {r[1]}")

# ── Tabelas com contagem de linhas ────────────────────────────
print("\n[TABELAS — ordenadas por qtd de linhas]\n")
cursor.execute("""
    SELECT t.TABLE_NAME, ISNULL(p.rows, 0) AS linhas
    FROM INFORMATION_SCHEMA.TABLES t
    LEFT JOIN sys.partitions p
           ON p.object_id = OBJECT_ID(t.TABLE_SCHEMA+'.'+t.TABLE_NAME)
          AND p.index_id IN (0, 1)
    WHERE t.TABLE_TYPE = 'BASE TABLE'
    ORDER BY linhas DESC, t.TABLE_NAME
""")
rows = cursor.fetchall()
print(f"  {'Tabela':<55} {'Linhas':>10}")
print(f"  {'-'*55} {'-'*10}")
for r in rows:
    print(f"  {r[0]:<55} {r[1]:>10,}")

print(f"\n  Total: {len(rows)} tabelas")
print(f"\n✅ Banco: {DATABASE}\n")

conn.close()
