import pandas as pd
from src.conexao import get_engine

engine = get_engine()
func = 147
nome_like = "%CAIO HENRIQUE RODRIGUES DE SANTANA%"

views = [
    "VW_RH_FUNCIONARIO_CONTRATACAO",
    "VW_FUNCIONARIO_APONTAMENTO",
    "VW_FUNCIONARIO_SETOR",
    "VW_RH_APONT_DIASTRAB",
    "VW_RH_APONT_DIASSEMTRAB",
]

for v in views:
    print(f"\n=== {v} COLUNAS ===")
    try:
        cols = pd.read_sql(
            f"""
            SELECT COLUMN_NAME, DATA_TYPE
            FROM INFORMATION_SCHEMA.COLUMNS
            WHERE TABLE_NAME = '{v}'
            ORDER BY ORDINAL_POSITION
            """,
            engine,
        )
        print(cols.to_string(index=False))
    except Exception as e:
        print("ERRO COLS:", e)
        continue

    colnames = cols["COLUMN_NAME"].str.upper().tolist()

    filtros = []
    if "FUNCIONARIO" in colnames:
        filtros.append(f"FUNCIONARIO = {func}")
    if "NOME" in colnames:
        filtros.append(f"UPPER(NOME) LIKE UPPER('{nome_like}')")
    if "RAZAO" in colnames:
        filtros.append(f"UPPER(RAZAO) LIKE UPPER('{nome_like}')")

    if not filtros:
        print("Sem chave direta para filtrar por funcionário/nome")
        try:
            sample = pd.read_sql(f"SELECT TOP 3 * FROM {v}", engine)
            print(sample.to_string(index=False))
        except Exception as e:
            print("ERRO SAMPLE:", e)
        continue

    where = " OR ".join(filtros)
    print(f"\n=== {v} DADOS FILTRADOS ===")
    try:
        df = pd.read_sql(f"SELECT TOP 20 * FROM {v} WHERE {where}", engine)
        print(df.to_string(index=False) if not df.empty else "(sem registros)")
    except Exception as e:
        print("ERRO DADOS:", e)
