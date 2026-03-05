import pandas as pd
from sqlalchemy import text
from src.conexao import get_engine

engine = get_engine()

with engine.connect() as conn:
    atual = conn.execute(text("SELECT DB_NAME() AS DB_ATUAL")).fetchone()[0]
    print("DB atual:", atual)

    try:
        dbs = pd.read_sql("SELECT name FROM sys.databases ORDER BY name", conn)
    except Exception as e:
        print("Erro ao listar bases:", e)
        raise

print("\n=== BASES VISIVEIS ===")
print(dbs.to_string(index=False))

candidatas = [n for n in dbs['name'].tolist() if any(x in n.upper() for x in ['RH','FOLHA','PESSOAL','INDUSTRIAL','ERP'])]
print("\n=== BASES CANDIDATAS ===")
print(candidatas)

for db in candidatas:
    print(f"\n=== TESTE BASE {db} ===")
    try:
        q = f"""
        SELECT TOP 5
            f.CODIGO AS FUNCIONARIO,
            f.NOME,
            c.CODIGO AS CONTRATACAO,
            c.VLRSALARIO1,
            c.VLRSALARIO2,
            c.DTADMISSAO,
            c.DTDEMISSAO
        FROM [{db}].dbo.RH_FUNCIONARIO f
        JOIN [{db}].dbo.RH_CONTRATACAO c ON c.FUNCIONARIO = f.CODIGO
        WHERE UPPER(f.NOME) LIKE '%CAIO HENRIQUE RODRIGUES DE SANTANA%'
        ORDER BY c.CODIGO DESC
        """
        df = pd.read_sql(q, engine)
        print(df.to_string(index=False) if not df.empty else "(sem registros)")
    except Exception as e:
        print("Erro:", e)
