import pandas as pd
from src.conexao import get_engine

engine = get_engine()
func = 147
contr = 135

tabelas = [
    "RH_TRANSFERENCIAFUNCAO",
    "RH_PROMOCAO",
    "RH_GRATIFICACAO",
    "RH_LANCAMENTOS",
]

for t in tabelas:
    print(f"\n=== {t} COLUNAS ===")
    cols = pd.read_sql(
        f"""
        SELECT COLUMN_NAME
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = '{t}'
        ORDER BY ORDINAL_POSITION
        """,
        engine,
    )
    print(cols.to_string(index=False))

    colnames = cols["COLUMN_NAME"].str.upper().tolist()
    filtros = []
    if "FUNCIONARIO" in colnames:
        filtros.append(f"FUNCIONARIO = {func}")
    if "CONTRATACAO" in colnames:
        filtros.append(f"CONTRATACAO = {contr}")
    if "CONTRATACAOOLD" in colnames:
        filtros.append(f"CONTRATACAOOLD = {contr}")
    if "CONTRATACAONEW" in colnames:
        filtros.append(f"CONTRATACAONEW = {contr}")

    if not filtros:
        print("Sem coluna direta de vínculo com funcionário/contratação")
        continue

    where = " OR ".join(filtros)
    print(f"\n=== {t} DADOS ===")
    try:
        df = pd.read_sql(f"SELECT TOP 50 * FROM {t} WHERE {where} ORDER BY CODIGO DESC", engine)
        print(df.to_string(index=False) if not df.empty else "(sem registros)")
    except Exception as e:
        print("ERRO:", e)
