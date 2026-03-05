import pandas as pd
from src.conexao import get_engine

engine = get_engine()

sql = """
SELECT TOP 1 CODIGO, NRONOTA, TPDOCUMENTO, VEPEDIDO
FROM FN_NFS
WHERE NRONOTA = 41787
ORDER BY CODIGO DESC
"""
base = pd.read_sql(sql, engine)
print('FN_NFS:')
print(base.to_string(index=False))

if base.empty:
    raise SystemExit(0)

nfs_codigo = int(base.loc[0, 'CODIGO'])

sql_ad = f"""
SELECT NFS, VEPEDIDO
FROM FN_NFSADTPEDIDO
WHERE NFS = {nfs_codigo}
   OR NFS = 41787
"""
ad = pd.read_sql(sql_ad, engine)
print('\nFN_NFSADTPEDIDO:')
print(ad.to_string(index=False) if not ad.empty else 'sem linhas')

sql_it = f"""
SELECT i.NFS, i.CODIGO AS NFSITEM, i.PEDIDOITEM, pi.PEDIDO
FROM FN_NFSITENS i
LEFT JOIN VE_PEDIDOITENS pi ON pi.CODIGO = i.PEDIDOITEM
WHERE i.NFS = {nfs_codigo}
ORDER BY i.CODIGO
"""
it = pd.read_sql(sql_it, engine)
print('\nFN_NFSITENS -> VE_PEDIDOITENS:')
print(it.head(50).to_string(index=False) if not it.empty else 'sem linhas')
print('\nPedidos distintos pelos itens:', sorted([int(x) for x in it['PEDIDO'].dropna().unique().tolist()]))

if not it['PEDIDO'].dropna().empty:
    pedidos = ','.join(str(int(x)) for x in sorted(it['PEDIDO'].dropna().unique()))
    sql_v = f"""
    SELECT pv.PEDIDO, pv.CKPRINCIPAL, fv.TIPO, fv.RAZAO
    FROM VE_PEDIDOVENDEDOR pv
    JOIN FN_VENDEDORES fv ON fv.CODIGO = pv.VENDEDOR
    WHERE pv.PEDIDO IN ({pedidos})
    ORDER BY pv.PEDIDO, fv.TIPO, pv.CKPRINCIPAL DESC, fv.RAZAO
    """
    v = pd.read_sql(sql_v, engine)
    print('\nVendedores dos pedidos vinculados:')
    print(v.to_string(index=False) if not v.empty else 'sem linhas')
