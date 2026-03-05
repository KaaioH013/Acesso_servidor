"""
conexao.py — Gerencia a conexão com SQL Server via pyodbc / SQLAlchemy.
"""

import os
import pyodbc
from sqlalchemy import create_engine, text
from dotenv import load_dotenv
from urllib.parse import quote_plus

load_dotenv()


def listar_drivers():
    """Retorna os drivers ODBC disponíveis na máquina."""
    return pyodbc.drivers()


def get_connection_string() -> str:
    """Monta a connection string a partir do .env"""
    server   = os.getenv("DB_SERVER", "")
    port     = os.getenv("DB_PORT", "1433")
    database = os.getenv("DB_DATABASE", "")
    username = os.getenv("DB_USERNAME", "")
    password = os.getenv("DB_PASSWORD", "")
    driver   = os.getenv("DB_DRIVER", "SQL Server")

    # Escapa senha com caracteres especiais (; { }) para não quebrar a connection string
    password_escaped = "{" + password.replace("}", "}}") + "}"

    return (
        f"DRIVER={{{driver}}};"
        f"SERVER={server},{port};"
        f"DATABASE={database};"
        f"UID={username};"
        f"PWD={password_escaped};"
        "Readonly=Yes;"
        "TrustServerCertificate=Yes;"
    )


def get_pyodbc_connection():
    """Retorna uma conexão pyodbc (para queries simples)."""
    conn_str = get_connection_string()
    return pyodbc.connect(conn_str, timeout=10)


def get_engine():
    """Retorna um SQLAlchemy engine (para pandas read_sql)."""
    conn_str = get_connection_string()
    params = quote_plus(conn_str)
    engine = create_engine(
        f"mssql+pyodbc:///?odbc_connect={params}",
        connect_args={"timeout": 10},
    )
    return engine


def testar_conexao():
    """Testa a conexão e retorna True se ok."""
    try:
        engine = get_engine()
        with engine.connect() as conn:
            resultado = conn.execute(text("SELECT @@VERSION")).scalar()
        print("✅ Conexão OK!")
        print(f"   {resultado[:80]}...")
        return True
    except Exception as e:
        print(f"❌ Falha na conexão: {e}")
        return False
