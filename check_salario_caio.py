import pandas as pd
from src.conexao import get_engine

engine = get_engine()

q_func = """
SELECT TOP 20
    f.CODIGO AS FUNCIONARIO,
    f.NOME,
    f.CKATIVO,
    f.CPF,
    f.NUMEROFUNC,
    f.FILIAL
FROM RH_FUNCIONARIO f
WHERE UPPER(f.NOME) LIKE '%CAIO%'
   OR UPPER(f.NOME) LIKE '%SANTANA%'
ORDER BY f.NOME
"""

q_contrato = """
SELECT TOP 50
    c.CODIGO AS CONTRATACAO,
    c.FUNCIONARIO,
    f.NOME,
    c.CKATIVO AS ATIVO_CONTRATO,
    c.DTADMISSAO,
    c.DTDEMISSAO,
    c.VLRSALARIO1,
    c.VLRSALARIO2,
    c.SALREF,
    c.CARGO,
    c.FUNCAO,
    c.TPPAGAMENTO,
    c.DEPARTAMENTO,
    c.SETOR
FROM RH_CONTRATACAO c
JOIN RH_FUNCIONARIO f ON f.CODIGO = c.FUNCIONARIO
WHERE UPPER(f.NOME) LIKE '%CAIO%'
   OR UPPER(f.NOME) LIKE '%SANTANA%'
ORDER BY c.CODIGO DESC
"""

q_hist = """
SELECT TOP 200
    hs.CODIGO,
    hs.CONTRATACAO,
    hs.FUNCIONARIO,
    f.NOME,
    hs.STATUS,
    hs.AUTORIZADO,
    hs.DTCADASTRO,
    hs.USERDATE,
    hs.VLRSALARIO1,
    hs.VLRSALARIO2,
    hs.SALREF,
    hs.MOTIVO,
    hs.ORIGEM
FROM RH_CONTRATACAOSALARIO hs
JOIN RH_FUNCIONARIO f ON f.CODIGO = hs.FUNCIONARIO
WHERE UPPER(f.NOME) LIKE '%CAIO%'
   OR UPPER(f.NOME) LIKE '%SANTANA%'
ORDER BY hs.CODIGO DESC
"""

print("\n=== FUNCIONARIO(S) ENCONTRADO(S) ===")
func = pd.read_sql(q_func, engine)
print(func.to_string(index=False) if not func.empty else "(nenhum)")

print("\n=== CONTRATOS ===")
cont = pd.read_sql(q_contrato, engine)
print(cont.to_string(index=False) if not cont.empty else "(nenhum)")

print("\n=== HISTORICO SALARIAL ===")
hist = pd.read_sql(q_hist, engine)
print(hist.to_string(index=False) if not hist.empty else "(nenhum)")
