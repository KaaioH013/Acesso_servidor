import sys
sys.path.insert(0, 'src')
from conexao import get_engine
import pandas as pd
engine = get_engine()

TPVENDA_EXCLUIR = (7, 21, 5, 12, 24, 11, 26, 6, 15, 16, 8, 19, 9, 17, 53, 18, 65, 23)
tv = ','.join(str(x) for x in TPVENDA_EXCLUIR)

df = pd.read_sql(f"""
    SELECT
        p.DTPEDIDO              AS dt_pedido,
        p.CODIGO                AS pedido,
        p.NUMINTERNO            AS num_interno,
        p.STATUS                AS status_pedido,
        c.RAZAO                 AS cliente,
        c.UF                    AS uf_cliente,
        v.RAZAO                 AS vendedor,
        i.SEQ                   AS seq,
        i.MATERIAL              AS cod_material,
        i.DESCRICAO             AS descricao,
        i.UNIDADE               AS unidade,
        t.DESCRICAO             AS tipo_venda,
        i.QTDE                  AS qtde_pedida,
        i.QTDEFAT               AS qtde_faturada,
        i.VLRUNITARIO           AS vlr_unitario,
        i.VLRTOTAL              AS vlr_total,
        i.STATUS                AS status_item
    FROM VE_PEDIDOITENS i
    JOIN VE_PEDIDO p              ON p.CODIGO = i.PEDIDO
    LEFT JOIN FN_FORNECEDORES c   ON c.CODIGO = p.CLIENTE
    LEFT JOIN FN_VENDEDORES v     ON v.CODIGO = p.VENDEDOR
    LEFT JOIN VE_TPVENDA t        ON t.CODIGO = i.TPVENDA
    WHERE p.DTPEDIDO BETWEEN '2025-04-01' AND '2025-04-30'
      AND p.STATUS <> 'C'
      AND i.STATUS <> 'C'
      AND i.TPVENDA NOT IN ({tv})
      AND i.MATERIAL NOT LIKE '8%'
      AND i.FLAGSUB <> 'S'
      AND p.CODIGO NOT IN (
          SELECT DISTINCT i2.PEDIDO FROM VE_PEDIDOITENS i2 WHERE i2.TPVENDA = 23
      )
      AND p.CODIGO NOT IN (
          SELECT DISTINCT p2.CODIGO FROM VE_PEDIDO p2
          JOIN FN_FORNECEDORES f ON f.CODIGO = p2.CLIENTE
          WHERE f.UF = 'EX' AND p2.DTPEDIDO BETWEEN '2025-04-01' AND '2025-04-30'
      )
    ORDER BY p.DTPEDIDO, p.CODIGO, i.SEQ
""", engine)

print(f"=== ABRIL 2025 — Itens de Pedidos (Peças) ===")
print(f"Total linhas (itens): {len(df)}")
print(f"Pedidos distintos:    {df['pedido'].nunique()}")
print(f"Valor total:          R$ {df['vlr_total'].sum():,.2f}")
print(f"\nPor status do pedido:")
print(df.groupby('status_pedido')['pedido'].nunique())
print(f"\nPor status do item:")
print(df['status_item'].value_counts())
print(f"\nPor vendedor:")
print(df.groupby('vendedor')['vlr_total'].sum().sort_values(ascending=False))
print(f"\nPor tipo de venda:")
print(df['tipo_venda'].value_counts())
