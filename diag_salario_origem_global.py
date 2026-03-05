import pandas as pd
from src.conexao import get_engine


def main():
    engine = get_engine()

    q_cols_tables = """
    SELECT 'TABLE' AS OBJ_TIPO, c.TABLE_NAME AS OBJ_NOME, c.COLUMN_NAME, c.DATA_TYPE
    FROM INFORMATION_SCHEMA.COLUMNS c
    JOIN INFORMATION_SCHEMA.TABLES t ON t.TABLE_NAME = c.TABLE_NAME AND t.TABLE_TYPE = 'BASE TABLE'
    WHERE c.COLUMN_NAME LIKE '%SALAR%'
       OR c.COLUMN_NAME LIKE '%LIQ%'
       OR c.COLUMN_NAME LIKE '%VENCT%'
       OR c.COLUMN_NAME LIKE '%REMUN%'
       OR c.COLUMN_NAME LIKE '%PROLABORE%'
    """

    q_cols_views = """
    SELECT 'VIEW' AS OBJ_TIPO, c.TABLE_NAME AS OBJ_NOME, c.COLUMN_NAME, c.DATA_TYPE
    FROM INFORMATION_SCHEMA.COLUMNS c
    JOIN INFORMATION_SCHEMA.VIEWS v ON v.TABLE_NAME = c.TABLE_NAME
    WHERE c.COLUMN_NAME LIKE '%SALAR%'
       OR c.COLUMN_NAME LIKE '%LIQ%'
       OR c.COLUMN_NAME LIKE '%VENCT%'
       OR c.COLUMN_NAME LIKE '%REMUN%'
       OR c.COLUMN_NAME LIKE '%PROLABORE%'
    """

    q_func_cols_tables = """
    SELECT c.TABLE_NAME
    FROM INFORMATION_SCHEMA.COLUMNS c
    JOIN INFORMATION_SCHEMA.TABLES t ON t.TABLE_NAME = c.TABLE_NAME AND t.TABLE_TYPE = 'BASE TABLE'
    WHERE c.COLUMN_NAME IN ('FUNCIONARIO','CONTRATACAO','CPF','NUMEROFUNC')
    GROUP BY c.TABLE_NAME
    ORDER BY c.TABLE_NAME
    """

    q_func_cols_views = """
    SELECT c.TABLE_NAME
    FROM INFORMATION_SCHEMA.COLUMNS c
    JOIN INFORMATION_SCHEMA.VIEWS v ON v.TABLE_NAME = c.TABLE_NAME
    WHERE c.COLUMN_NAME IN ('FUNCIONARIO','CONTRATACAO','CPF','NUMEROFUNC')
    GROUP BY c.TABLE_NAME
    ORDER BY c.TABLE_NAME
    """

    df_t = pd.read_sql(q_cols_tables, engine)
    df_v = pd.read_sql(q_cols_views, engine)
    df_tf = pd.read_sql(q_func_cols_tables, engine)
    df_vf = pd.read_sql(q_func_cols_views, engine)

    print('=== COLUNAS SALARIAIS EM TABELAS ===')
    print(df_t.sort_values(['OBJ_NOME','COLUMN_NAME']).to_string(index=False) if not df_t.empty else '(nenhuma)')

    print('\n=== COLUNAS SALARIAIS EM VIEWS ===')
    print(df_v.sort_values(['OBJ_NOME','COLUMN_NAME']).to_string(index=False) if not df_v.empty else '(nenhuma)')

    print('\n=== OBJETOS COM CHAVE DE FUNCIONARIO (TABELAS) ===')
    print(df_tf.to_string(index=False) if not df_tf.empty else '(nenhuma)')

    print('\n=== OBJETOS COM CHAVE DE FUNCIONARIO (VIEWS) ===')
    print(df_vf.to_string(index=False) if not df_vf.empty else '(nenhuma)')


if __name__ == '__main__':
    main()
