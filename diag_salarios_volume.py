import pandas as pd
from src.conexao import get_engine

engine = get_engine()

tabelas = [
    "RH_LANCPAGTO",
    "RH_LANC13PAGTO",
    "RH_LANCFERIAS",
    "RH_LANCRECISAO",
    "RH_LANCAMENTOSMES",
    "RH_FUNCIONARIO",
    "RH_CONTRATACAO",
    "RH_CONTRATACAOSALARIO",
]

for t in tabelas:
    print(f"\n=== {t} ===")
    try:
        c = pd.read_sql(f"SELECT COUNT(*) AS QTD FROM {t}", engine)
        print("Registros:", int(c.iloc[0, 0]))
    except Exception as e:
        print("Erro count:", e)
        continue

    try:
        s = pd.read_sql(f"SELECT TOP 3 * FROM {t}", engine)
        print("Colunas:", ", ".join(s.columns.tolist()))
    except Exception as e:
        print("Erro sample:", e)
