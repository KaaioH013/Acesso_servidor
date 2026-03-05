"""
inspecionar_vendas.py — Mapeia todas as tabelas do módulo VE_ (Vendas)
com colunas, tipos e contagem de linhas.
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

conn = pyodbc.connect(conn_str, timeout=15)
cursor = conn.cursor()

print(f"\n{'='*65}")
print(f"  MODULO VENDAS (VE_)  |  Banco: {DATABASE}")
print(f"{'='*65}")

# -- 1. Tabelas VE_ com contagem de linhas
print("\n[TABELAS VE_ — ordenadas por qtd de linhas]\n")
cursor.execute("""
    SELECT t.TABLE_NAME, ISNULL(p.rows, 0) AS linhas
    FROM INFORMATION_SCHEMA.TABLES t
    LEFT JOIN sys.partitions p
           ON p.object_id = OBJECT_ID(t.TABLE_SCHEMA+'.'+t.TABLE_NAME)
          AND p.index_id IN (0, 1)
    WHERE t.TABLE_TYPE = 'BASE TABLE'
      AND t.TABLE_NAME LIKE 'VE[_]%'
    ORDER BY linhas DESC, t.TABLE_NAME
""")
tabelas_ve = cursor.fetchall()
print(f"  {'Tabela':<50} {'Linhas':>10}")
print(f"  {'-'*50} {'-'*10}")
for r in tabelas_ve:
    print(f"  {r[0]:<50} {r[1]:>10,}")

print(f"\n  Total: {len(tabelas_ve)} tabelas VE_\n")

# -- 2. Colunas das principais tabelas de vendas
principais = [
    "VE_PEDIDO",
    "VE_PEDIDOITENS",
    "VE_ORCAMENTO",
    "VE_ORCAMENTOITENS",
    "VE_CLIENTE",
    "VE_CLIENTES",
    "VE_PLANILHACUSTO",
    "VE_VENDEDOR",
    "VE_VENDEDORES",
    "VE_REPRESENTANTE",
    "VE_REPRESENTANTES",
    "VE_EXPEDICAO",
    "VE_EXPEDICAOITENS",
    "VE_CONTRATO",
    "VE_CONTRATOITENS",
    "VE_META",
    "VE_METAITENS",
    "VE_TABELAPRECO",
    "VE_TABELAPRECOITENS",
    "VE_COMISSAO",
    "VE_COMISSAOITENS",
]

# Filtra só as que existem no banco
tabelas_existentes = [r[0] for r in tabelas_ve]
# Também verifica com busca parcial
cursor.execute("""
    SELECT DISTINCT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES
    WHERE TABLE_NAME LIKE 'VE[_]%'
    AND TABLE_TYPE = 'BASE TABLE'
""")
todas_ve = [r[0] for r in cursor.fetchall()]

# Prioridade: tabelas com mais linhas primeiro
tabelas_com_linhas = {r[0]: r[1] for r in tabelas_ve}
candidatas = sorted(
    [t for t in todas_ve if tabelas_com_linhas.get(t, 0) > 0],
    key=lambda t: tabelas_com_linhas.get(t, 0),
    reverse=True
)[:20]  # Top 20 com mais dados

print(f"\n{'='*65}")
print("  COLUNAS DAS PRINCIPAIS TABELAS VE_ (com dados)")
print(f"{'='*65}")

for tabela in candidatas:
    cursor.execute("""
        SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, IS_NULLABLE
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = ? AND TABLE_SCHEMA = 'dbo'
        ORDER BY ORDINAL_POSITION
    """, tabela)
    colunas = cursor.fetchall()
    linhas = tabelas_com_linhas.get(tabela, 0)
    print(f"\n  +-- {tabela}  ({linhas:,} linhas) -----")
    for c in colunas:
        tam = f"({c[2]})" if c[2] else ""
        nulo = "NULL" if c[3] == "YES" else "NOT NULL"
        print(f"  |  {c[0]:<40} {c[1]}{tam:<15} {nulo}")
    print(f"  +{'-'*60}")

# -- 3. Views de vendas
print(f"\n{'='*65}")
print("  VIEWS VE_ (uteis para relatorios)")
print(f"{'='*65}\n")
cursor.execute("""
    SELECT TABLE_NAME FROM INFORMATION_SCHEMA.VIEWS
    WHERE TABLE_NAME LIKE 'VE[_]%'
    ORDER BY TABLE_NAME
""")
views = cursor.fetchall()
if views:
    for v in views:
        print(f"  • {v[0]}")
else:
    print("  (nenhuma view VE_ encontrada)")

# -- 4. Relacionamentos saindo de tabelas VE_
print(f"\n{'='*65}")
print("  RELACIONAMENTOS (FK) das tabelas VE_")
print(f"{'='*65}\n")
cursor.execute("""
    SELECT tp.name AS origem, cp.name AS col_origem,
           tr.name AS destino, cr.name AS col_destino
    FROM sys.foreign_keys fk
    JOIN sys.foreign_key_columns fkc ON fk.object_id = fkc.constraint_object_id
    JOIN sys.tables  tp ON tp.object_id = fkc.parent_object_id
    JOIN sys.columns cp ON cp.object_id = fkc.parent_object_id   AND cp.column_id = fkc.parent_column_id
    JOIN sys.tables  tr ON tr.object_id = fkc.referenced_object_id
    JOIN sys.columns cr ON cr.object_id = fkc.referenced_object_id AND cr.column_id = fkc.referenced_column_id
    WHERE tp.name LIKE 'VE[_]%'
    ORDER BY tp.name, tr.name
""")
fks = cursor.fetchall()
print(f"  {'Tabela Origem':<35} {'Coluna':<30} → {'Tabela Destino':<35} {'Coluna'}")
print(f"  {'-'*35} {'-'*30}   {'-'*35} {'-'*20}")
for r in fks:
    print(f"  {r[0]:<35} {r[1]:<30} → {r[2]:<35} {r[3]}")

print(f"\n✅ Mapeamento de Vendas concluído.\n")
conn.close()
