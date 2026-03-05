"""
explorador.py — Explora o schema do banco: tabelas, colunas, FKs, contagens.
"""

import pandas as pd
from sqlalchemy import text
from src.conexao import get_engine


def listar_tabelas(schema: str = None) -> pd.DataFrame:
    """Lista todas as tabelas (e views) do banco."""
    engine = get_engine()
    where = f"AND t.TABLE_SCHEMA = '{schema}'" if schema else ""
    query = f"""
        SELECT
            t.TABLE_SCHEMA   AS [schema],
            t.TABLE_NAME     AS tabela,
            t.TABLE_TYPE     AS tipo,
            p.rows           AS qtd_linhas
        FROM INFORMATION_SCHEMA.TABLES t
        LEFT JOIN sys.partitions p
               ON p.object_id = OBJECT_ID(t.TABLE_SCHEMA + '.' + t.TABLE_NAME)
              AND p.index_id IN (0, 1)
        WHERE t.TABLE_TYPE IN ('BASE TABLE', 'VIEW')
        {where}
        ORDER BY t.TABLE_SCHEMA, t.TABLE_TYPE, t.TABLE_NAME
    """
    with engine.connect() as conn:
        return pd.read_sql(text(query), conn)


def descrever_tabela(tabela: str, schema: str = "dbo") -> pd.DataFrame:
    """Retorna as colunas de uma tabela com tipo, tamanho e nulabilidade."""
    engine = get_engine()
    query = """
        SELECT
            COLUMN_NAME        AS coluna,
            DATA_TYPE          AS tipo,
            CHARACTER_MAXIMUM_LENGTH AS tamanho,
            IS_NULLABLE        AS nulo,
            COLUMN_DEFAULT     AS [default],
            ORDINAL_POSITION   AS ordem
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME   = :tabela
          AND TABLE_SCHEMA = :schema
        ORDER BY ORDINAL_POSITION
    """
    with engine.connect() as conn:
        return pd.read_sql(text(query), conn, params={"tabela": tabela, "schema": schema})


def listar_chaves_estrangeiras() -> pd.DataFrame:
    """Lista todos os relacionamentos (FKs) do banco."""
    engine = get_engine()
    query = """
        SELECT
            fk.name                          AS fk_nome,
            tp.name                          AS tabela_origem,
            cp.name                          AS coluna_origem,
            tr.name                          AS tabela_destino,
            cr.name                          AS coluna_destino
        FROM sys.foreign_keys fk
        JOIN sys.foreign_key_columns fkc ON fk.object_id = fkc.constraint_object_id
        JOIN sys.tables tp  ON tp.object_id  = fkc.parent_object_id
        JOIN sys.columns cp ON cp.object_id  = fkc.parent_object_id
                           AND cp.column_id  = fkc.parent_column_id
        JOIN sys.tables tr  ON tr.object_id  = fkc.referenced_object_id
        JOIN sys.columns cr ON cr.object_id  = fkc.referenced_object_id
                           AND cr.column_id  = fkc.referenced_column_id
        ORDER BY tabela_origem, fk_nome
    """
    with engine.connect() as conn:
        return pd.read_sql(text(query), conn)


def listar_indices(tabela: str = None) -> pd.DataFrame:
    """Lista os índices do banco (ou de uma tabela específica)."""
    engine = get_engine()
    where = f"AND t.name = '{tabela}'" if tabela else ""
    query = f"""
        SELECT
            t.name  AS tabela,
            i.name  AS indice,
            i.type_desc AS tipo,
            i.is_unique AS unico,
            i.is_primary_key AS pk,
            STRING_AGG(c.name, ', ') WITHIN GROUP (ORDER BY ic.key_ordinal) AS colunas
        FROM sys.indexes i
        JOIN sys.tables t ON t.object_id = i.object_id
        JOIN sys.index_columns ic ON ic.object_id = i.object_id AND ic.index_id = i.index_id
        JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
        WHERE i.name IS NOT NULL
        {where}
        GROUP BY t.name, i.name, i.type_desc, i.is_unique, i.is_primary_key
        ORDER BY t.name, i.name
    """
    with engine.connect() as conn:
        return pd.read_sql(text(query), conn)


def resumo_banco() -> pd.DataFrame:
    """Retorna contagem de objetos por tipo no banco."""
    engine = get_engine()
    query = """
        SELECT type_desc AS tipo, COUNT(*) AS quantidade
        FROM sys.objects
        WHERE is_ms_shipped = 0
        GROUP BY type_desc
        ORDER BY quantidade DESC
    """
    with engine.connect() as conn:
        return pd.read_sql(text(query), conn)


def buscar_coluna(nome_coluna: str) -> pd.DataFrame:
    """Busca em quais tabelas uma coluna (ou parte do nome) aparece."""
    engine = get_engine()
    query = """
        SELECT TABLE_SCHEMA AS [schema], TABLE_NAME AS tabela, COLUMN_NAME AS coluna, DATA_TYPE AS tipo
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE COLUMN_NAME LIKE :nome
        ORDER BY TABLE_NAME, COLUMN_NAME
    """
    with engine.connect() as conn:
        return pd.read_sql(text(query), conn, params={"nome": f"%{nome_coluna}%"})


def preview_tabela(tabela: str, schema: str = "dbo", linhas: int = 10) -> pd.DataFrame:
    """Retorna as primeiras N linhas de uma tabela."""
    engine = get_engine()
    query = f"SELECT TOP {linhas} * FROM [{schema}].[{tabela}]"
    with engine.connect() as conn:
        return pd.read_sql(text(query), conn)
