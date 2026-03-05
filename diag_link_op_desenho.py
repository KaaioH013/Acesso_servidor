import pandas as pd
from src.conexao import get_engine

engine = get_engine()

queries = {
    "OP exemplo": """
        SELECT TOP 20 CODIGO, NROOP, OPSTATUS, MATERIAL, DESENHO, FICHA, PLANEJAMENTO, DESCRICAO
        FROM PR_OP
        WHERE CODIGO = 137529 OR NROOP = 135300
    """,
    "LISTAENG por CODDESENHO=DESENHO": """
        SELECT TOP 20 CODIGO, CODDESENHO, POSICAO, DESENHO, CODMP, DESCRICAO
        FROM PR_LISTAENG
        WHERE CODDESENHO IN (
            SELECT CAST(DESENHO AS numeric(18,0)) FROM PR_OP WHERE CODIGO = 137529
        )
        ORDER BY CODIGO DESC
    """,
    "LISTAENG por CODDESENHO=MATERIAL": """
        SELECT TOP 20 CODIGO, CODDESENHO, POSICAO, DESENHO, CODMP, DESCRICAO
        FROM PR_LISTAENG
        WHERE CODDESENHO IN (
            SELECT CAST(MATERIAL AS numeric(18,0)) FROM PR_OP WHERE CODIGO = 137529
        )
        ORDER BY CODIGO DESC
    """,
    "LISTAENG onde DESENHO=MATERIAL OP": """
        SELECT TOP 20 CODIGO, CODDESENHO, POSICAO, DESENHO, CODMP, DESCRICAO
        FROM PR_LISTAENG
        WHERE DESENHO = (SELECT MATERIAL FROM PR_OP WHERE CODIGO = 137529)
        ORDER BY CODIGO DESC
    """,
    "PR_DESENHO por numero do desenho OP": """
        SELECT TOP 20 CODIGO, DESENHO, DESCRICAO, STATUS, DTCADASTRO
        FROM PR_DESENHO
        WHERE DESENHO = (SELECT DESENHO FROM PR_OP WHERE CODIGO = 137529)
        ORDER BY CODIGO DESC
    """,
    "LISTAENG via PR_DESENHO": """
        SELECT TOP 50 le.CODIGO, le.CODDESENHO, le.POSICAO, le.DESENHO, le.CODMP, le.DESCRICAO
        FROM PR_OP op
        JOIN PR_DESENHO d ON d.DESENHO = op.DESENHO
        JOIN PR_LISTAENG le ON le.CODDESENHO = d.CODIGO
        WHERE op.CODIGO = 137529
        ORDER BY le.POSICAO
    """,
    "ROTEIRO OPCAO via LISTAENG raiz": """
        SELECT TOP 20 o.CODIGO, o.LISTAENG, o.CKPADRAO, o.DTPROJETADO, o.REVISAO
        FROM PR_OP op
        JOIN PR_DESENHO d ON d.DESENHO = op.DESENHO
        JOIN PR_LISTAENG le ON le.CODDESENHO = d.CODIGO AND le.POSICAO = '00'
        JOIN PR_DESENHOROTEIROOPCAO o ON o.LISTAENG = le.CODIGO
        WHERE op.CODIGO = 137529
        ORDER BY o.CODIGO DESC
    """,
    "ROTEIRO OP": """
        SELECT TOP 20 CODIGO, OP, INSERTDATE, DTPROJETADO
        FROM PR_OPROTEIRO
        WHERE OP = 137529
        ORDER BY CODIGO DESC
    """,
    "ROTEIRO ITENS OP": """
        SELECT TOP 20 ri.OPROTEIRO, ri.MATERIAL, ri.QTDE, mm.DESCRICAO
        FROM PR_OPROTEIROITEM ri
        JOIN PR_OPROTEIRO r ON r.CODIGO = ri.OPROTEIRO
        LEFT JOIN MT_MATERIAL mm ON mm.CODIGO = ri.MATERIAL
        WHERE r.OP = 137529
    """,
    "Compras do material 2101029663": """
        SELECT TOP 20 i.PEDIDO, p.DTPEDIDO, p.STATUS ST_PED, i.STATUS ST_ITEM, i.MATERIAL, i.QTDEPED
        FROM CO_PEDIDOITENS i
        JOIN CO_PEDIDOS p ON p.CODIGO = i.PEDIDO
        WHERE i.MATERIAL = '2101029663'
        ORDER BY p.DTPEDIDO DESC
    """,
}

for nome, sql in queries.items():
    print(f"\n=== {nome} ===")
    try:
        df = pd.read_sql(sql, engine)
        print(df.to_string(index=False) if not df.empty else "(sem registros)")
    except Exception as e:
        print(f"ERRO: {e}")
