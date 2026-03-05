import pandas as pd
from src.conexao import get_engine

engine = get_engine()
func = 147
contr = 135

queries = {
    "RH_TRANSFERENCIAFUNCAO": f"SELECT TOP 50 * FROM RH_TRANSFERENCIAFUNCAO WHERE CONTRATACAO = {contr} OR FUNCIONARIO = {func} ORDER BY CODIGO DESC",
    "RH_PROMOCAO": f"SELECT TOP 50 * FROM RH_PROMOCAO WHERE CONTRATACAO = {contr} OR FUNCIONARIO = {func} ORDER BY CODIGO DESC",
    "RH_GRATIFICACAO": f"SELECT TOP 50 * FROM RH_GRATIFICACAO WHERE CONTRATACAO = {contr} OR FUNCIONARIO = {func} ORDER BY CODIGO DESC",
    "RH_LANCAMENTOS": f"SELECT TOP 50 * FROM RH_LANCAMENTOS WHERE CONTRATACAO = {contr} OR FUNCIONARIO = {func} ORDER BY CODIGO DESC",
    "RH_LANCAMENTOSMES": f"SELECT TOP 50 * FROM RH_LANCAMENTOSMES WHERE CONTRATACAO = {contr} OR FUNCIONARIO = {func} ORDER BY CODIGO DESC",
}

for nome, sql in queries.items():
    print(f"\n=== {nome} ===")
    try:
        df = pd.read_sql(sql, engine)
        if df.empty:
            print("(sem registros)")
        else:
            print(df.to_string(index=False))
    except Exception as e:
        print("ERRO:", e)
