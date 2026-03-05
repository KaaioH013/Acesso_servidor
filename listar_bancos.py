"""
listar_bancos.py — Lista todos os bancos de dados disponíveis no servidor.
Útil para descobrir o DB_DATABASE correto antes de configurar o .env.

Execute: c:/python314/python.exe listar_bancos.py
"""

import os
import pyodbc
from dotenv import load_dotenv

load_dotenv()

SERVER   = os.getenv("DB_SERVER")
PORT     = os.getenv("DB_PORT", "1433")
USERNAME = os.getenv("DB_USERNAME")
PASSWORD = os.getenv("DB_PASSWORD")
DRIVER   = os.getenv("DB_DRIVER", "ODBC Driver 17 for SQL Server")

print(f"\nConectando em: {SERVER}:{PORT}  com usuário: {USERNAME}\n")

# Conecta no banco 'master' (sempre existe) só para listar os outros
# Escapa senha com caracteres especiais (; { }) para não quebrar a connection string
password_escaped = "{" + PASSWORD.replace("}", "}}") + "}"

conn_str = (
    f"DRIVER={{{DRIVER}}};"
    f"SERVER={SERVER},{PORT};"
    f"DATABASE=master;"
    f"UID={USERNAME};PWD={password_escaped};"
    "TrustServerCertificate=Yes;"
)

try:
    conn = pyodbc.connect(conn_str, timeout=10)
    cursor = conn.cursor()
    cursor.execute("""
        SELECT name, create_date, state_desc
        FROM sys.databases
        WHERE state_desc = 'ONLINE'
        ORDER BY name
    """)
    rows = cursor.fetchall()
    conn.close()

    print(f"{'Banco de Dados':<40} {'Criado em':<22} {'Status'}")
    print("-" * 75)
    for row in rows:
        print(f"{row[0]:<40} {str(row[1]):<22} {row[2]}")

    print(f"\nTotal: {len(rows)} banco(s) encontrado(s)")
    print("\n👉 Copie o nome correto para DB_DATABASE no arquivo .env")

except Exception as e:
    print(f"❌ Erro ao conectar: {e}")
    print("\nVerifique:")
    print("  1. DB_SERVER no .env está correto")
    print("  2. DB_USERNAME e DB_PASSWORD estão preenchidos")
    print("  3. O driver ODBC está instalado (rode: python -c \"import pyodbc; print(pyodbc.drivers())\")")
