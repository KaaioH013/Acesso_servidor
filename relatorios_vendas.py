"""
relatorios_vendas.py â€” Queries e exportaÃ§Ãµes prontas para Coordenador de Vendas
Banco: INDUSTRIAL | Sectra ERP

Execute: c:/python314/python.exe relatorios_vendas.py
"""
import os, sys
sys.path.insert(0, ".")
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv
load_dotenv()

import pyodbc
import pandas as pd
from sqlalchemy import create_engine, text
from urllib.parse import quote_plus

# â”€â”€ ConexÃ£o â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
SERVER   = os.getenv("DB_SERVER")
PORT     = os.getenv("DB_PORT", "1433")
DATABASE = os.getenv("DB_DATABASE")
USERNAME = os.getenv("DB_USERNAME")
PASSWORD = os.getenv("DB_PASSWORD")
DRIVER   = os.getenv("DB_DRIVER", "SQL Server")

pwd_escaped = "{" + PASSWORD.replace("}", "}}") + "}"
conn_str = (
    f"DRIVER={{{DRIVER}}};SERVER={SERVER},{PORT};"
    f"DATABASE={DATABASE};UID={USERNAME};PWD={pwd_escaped};"
    "TrustServerCertificate=Yes;"
)
params = quote_plus(conn_str)
engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}", connect_args={"timeout": 15})

OUTPUT = Path("exports")
OUTPUT.mkdir(exist_ok=True)
ts = datetime.now().strftime("%Y%m%d_%H%M%S")


def sql(query, params=None):
    with engine.connect() as conn:
        return pd.read_sql(text(query), conn, params=params or {})

def ajustar_colunas(writer, aba, df):
    ws = writer.sheets[aba]
    for i, col in enumerate(df.columns, 1):
        larg = max(len(str(col)), df[col].astype(str).str.len().max() if len(df) else 0)
        ws.column_dimensions[ws.cell(1, i).column_letter].width = min(larg + 4, 60)

# ============================================================
# FILTROS GLOBAIS â€” PEÃ‡AS (sem bombas, sem nÃ£o-vendas)
# ============================================================
#
# Tipos de venda EXCLUIDOS (garantia, remessa, transferÃªncia, retorno)
#   7  = REMESSA EM GARANTIA
#  21  = REMESSA EM GARANTIA
#   5  = REMESSA P/ DEMONSTRACAO
#  12  = REMESSA P/ DEMONSTRACAO EM FEIRA
#  24  = SIMPLES REMESSA
#  11  = REMESSA COM FIM ESPECIFICO DE EXPORTACAO (PROD)
#  26  = REMESSA COM FIM ESPECIFICO DE EXPORTACAO (REVENDA)
#   6  = REMESSA P/ BRINDE
#  15  = REMESSA P/ BONIF E BRINDE P/ CF
#  16  = REMESSA PARA AMOSTRA
#   8  = REMESSA P/ TROCA
#  19  = TRANSFERENCIA
#   9  = RETORNO CONSERTO
#  17  = RETORNO DE DEMONSTRACAO
#  53  = DEVOLUCAO DE MERCADORIA RECEBIDA EM TRANSFERENCIA
#  18  = PRESTACAO DE SERVICOS
#  65  = REMESSA EM GARANTIA SEM ST
#
# Materiais iniciados com "8" = bombas completas â†’ excluidos
#
TPVENDA_EXCLUIR = (7, 21, 5, 12, 24, 11, 26, 6, 15, 16, 8, 19, 9, 17, 53, 18, 65)

# Subquery reutilizada para excluir pedidos de clientes estrangeiros (UF = 'EX')
# Mais confiavel que filtrar por TPVENDA: cobre exportacoes nao fiscais tambem
EXCLUI_PEDIDOS_EXPORT = """p.CODIGO NOT IN (
    SELECT p2.CODIGO FROM VE_PEDIDO p2
    JOIN FN_FORNECEDORES f2 ON f2.CODIGO = p2.CLIENTE
    WHERE f2.UF = 'EX'
)"""
TPVENDA_SQL = ",".join(str(x) for x in TPVENDA_EXCLUIR)   # para usar no IN


# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
# QUERY 1 â€” Itens de pedidos filtrados (peÃ§as, vendas reais)
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
def itens_vendas_pecas(data_ini: str, data_fim: str) -> pd.DataFrame:
    """
    Itens de pedidos no perÃ­odo:
    - Somente tipo de venda = VENDA (exclui remessa, garantia, transferÃªncia, etc.)
    - Exclui materiais comeÃ§ando com '8' (bombas)
    - Exclui pedidos cancelados (STATUS = 'C')
    """
    return sql(f"""
        SELECT
            p.DTPEDIDO              AS dt_pedido,
            p.CODIGO                AS pedido,
            p.NUMINTERNO            AS num_interno,
            p.STATUS                AS status_pedido,
            c.RAZAO                 AS cliente,
            c.CIDADE                AS cidade_cliente,
            c.UF                    AS uf_cliente,
            v.RAZAO                 AS vendedor,
            i.SEQ                   AS seq,
            i.MATERIAL              AS cod_material,
            i.DESCRICAO             AS descricao,
            i.UNIDADE               AS unidade,
            t.DESCRICAO             AS tipo_venda,
            i.QTDE                  AS qtde_pedida,
            i.QTDEFAT               AS qtde_faturada,
            i.QTDE - i.QTDEFAT      AS qtde_saldo,
            i.VLRUNITARIO           AS vlr_unitario,
            i.VLRTOTAL              AS vlr_total,
            i.DTPRAZO               AS dt_prazo,
            i.STATUS                AS status_item
        FROM VE_PEDIDOITENS i
        JOIN VE_PEDIDO p              ON p.CODIGO = i.PEDIDO
        LEFT JOIN FN_FORNECEDORES c   ON c.CODIGO = p.CLIENTE
        LEFT JOIN FN_VENDEDORES v     ON v.CODIGO = p.VENDEDOR
        LEFT JOIN VE_TPVENDA t        ON t.CODIGO = i.TPVENDA
        WHERE p.DTPEDIDO BETWEEN :ini AND :fim
          AND p.STATUS <> 'C'
          AND i.TPVENDA NOT IN ({TPVENDA_SQL})
          AND i.MATERIAL NOT LIKE '8%'
          AND i.FLAGSUB <> 'S'
          AND i.STATUS <> 'C'
          AND {EXCLUI_PEDIDOS_EXPORT}
        ORDER BY p.DTPEDIDO DESC, p.CODIGO, i.SEQ
    """, {"ini": data_ini, "fim": data_fim})


# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
# QUERY 2 â€” Pedidos resumo (cabeÃ§alho, filtro aplicado)
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
def pedidos_resumo(data_ini: str, data_fim: str) -> pd.DataFrame:
    """
    CabeÃ§alho dos pedidos que tÃªm ao menos 1 item de peÃ§a (nÃ£o bomba, nÃ£o remessa).
    Exclui cancelados.
    """
    return sql(f"""
        SELECT DISTINCT
            p.DTPEDIDO              AS dt_pedido,
            p.CODIGO                AS pedido,
            p.NUMINTERNO            AS num_interno,
            p.PEDIDOCLI             AS pedido_cliente,
            p.STATUS                AS status,
            p.DTSTATUS              AS dt_status,
            c.RAZAO                 AS cliente,
            c.CIDADE                AS cidade_cliente,
            c.UF                    AS uf,
            v.RAZAO                 AS vendedor,
            p.DTENTREGA             AS dt_entrega,
            p.VLRBRUTO              AS vlr_bruto,
            p.VLRDESCONTO           AS vlr_desconto,
            p.VLRTOTAL              AS vlr_total,
            p.VLRCOMISSAO           AS vlr_comissao,
            p.PERCCOMISSAO          AS perc_comissao,
            o.NUMERO                AS num_orcamento,
            o.DTCADASTRO            AS dt_orcamento
        FROM VE_PEDIDO p
        JOIN VE_PEDIDOITENS i         ON i.PEDIDO = p.CODIGO
                                     AND i.TPVENDA NOT IN ({TPVENDA_SQL})
                                     AND i.MATERIAL NOT LIKE '8%'
                                     AND i.FLAGSUB <> 'S'
        LEFT JOIN FN_FORNECEDORES c   ON c.CODIGO = p.CLIENTE
        LEFT JOIN FN_VENDEDORES v     ON v.CODIGO = p.VENDEDOR
        LEFT JOIN VE_ORCAMENTOS o     ON o.CODIGO = p.PEDORIGEM
        WHERE p.DTPEDIDO BETWEEN :ini AND :fim
          AND p.STATUS <> 'C'
          AND {EXCLUI_PEDIDOS_EXPORT}
        ORDER BY p.DTPEDIDO DESC, p.CODIGO DESC
    """, {"ini": data_ini, "fim": data_fim})


# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
# QUERY 3 â€” Faturamento por vendedor (peÃ§as / nÃ£o-bomba)
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
def faturamento_por_vendedor(data_ini: str, data_fim: str) -> pd.DataFrame:
    return sql(f"""
        SELECT
            v.RAZAO                                 AS vendedor,
            MONTH(p.DTPEDIDO)                       AS mes,
            DATENAME(MONTH, p.DTPEDIDO)             AS mes_nome,
            COUNT(DISTINCT p.CODIGO)                AS qtd_pedidos,
            SUM(i.VLRTOTAL)                         AS faturamento_itens_peca,
            AVG(i.VLRUNITARIO)                      AS preco_medio,
            SUM(p.VLRCOMISSAO)                      AS vlr_comissao
        FROM VE_PEDIDOITENS i
        JOIN VE_PEDIDO p              ON p.CODIGO = i.PEDIDO
        LEFT JOIN FN_VENDEDORES v     ON v.CODIGO = p.VENDEDOR
        WHERE p.DTPEDIDO BETWEEN :ini AND :fim
          AND p.STATUS <> 'C'
          AND i.TPVENDA NOT IN ({TPVENDA_SQL})
          AND i.MATERIAL NOT LIKE '8%'
          AND i.FLAGSUB <> 'S'
          AND i.STATUS <> 'C'
          AND {EXCLUI_PEDIDOS_EXPORT}
        GROUP BY v.RAZAO, MONTH(p.DTPEDIDO), DATENAME(MONTH, p.DTPEDIDO)
        ORDER BY v.RAZAO, mes
    """, {"ini": data_ini, "fim": data_fim})


# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
# QUERY 4 â€” Ranking de clientes (peÃ§as)
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
def ranking_clientes(data_ini: str, data_fim: str) -> pd.DataFrame:
    return sql(f"""
        SELECT
            p.CLIENTE                               AS cod_cliente,
            c.RAZAO                                 AS cliente,
            c.CGC                                   AS cnpj_cpf,
            c.CIDADE                                AS cidade,
            c.UF                                    AS uf,
            COUNT(DISTINCT p.CODIGO)                AS qtd_pedidos,
            SUM(i.VLRTOTAL)                         AS faturamento_pecas,
            AVG(i.VLRUNITARIO)                      AS preco_medio,
            MIN(p.DTPEDIDO)                         AS primeiro_pedido,
            MAX(p.DTPEDIDO)                         AS ultimo_pedido
        FROM VE_PEDIDOITENS i
        JOIN VE_PEDIDO p              ON p.CODIGO = i.PEDIDO
        LEFT JOIN FN_FORNECEDORES c   ON c.CODIGO = p.CLIENTE
        WHERE p.DTPEDIDO BETWEEN :ini AND :fim
          AND p.STATUS <> 'C'
          AND i.TPVENDA NOT IN ({TPVENDA_SQL})
          AND i.MATERIAL NOT LIKE '8%'
          AND i.FLAGSUB <> 'S'
          AND i.STATUS <> 'C'
          AND {EXCLUI_PEDIDOS_EXPORT}
        GROUP BY p.CLIENTE, c.RAZAO, c.CGC, c.CIDADE, c.UF
        ORDER BY faturamento_pecas DESC
    """, {"ini": data_ini, "fim": data_fim})


# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
# QUERY 5 â€” Produtos mais vendidos (peÃ§as)
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
def produtos_mais_vendidos(data_ini: str, data_fim: str, top: int = 50) -> pd.DataFrame:
    return sql(f"""
        SELECT TOP {top}
            i.MATERIAL                              AS cod_material,
            i.DESCRICAO                             AS descricao,
            i.UNIDADE                               AS unidade,
            t.DESCRICAO                             AS tipo_venda,
            SUM(i.QTDE)                             AS qtde_total,
            SUM(i.VLRTOTAL)                         AS faturamento_total,
            AVG(i.VLRUNITARIO)                      AS preco_medio,
            COUNT(DISTINCT i.PEDIDO)                AS qtd_pedidos
        FROM VE_PEDIDOITENS i
        JOIN VE_PEDIDO p              ON p.CODIGO = i.PEDIDO
        LEFT JOIN VE_TPVENDA t        ON t.CODIGO = i.TPVENDA
        WHERE p.DTPEDIDO BETWEEN :ini AND :fim
          AND p.STATUS <> 'C'
          AND i.TPVENDA NOT IN ({TPVENDA_SQL})
          AND i.MATERIAL NOT LIKE '8%'
          AND i.FLAGSUB <> 'S'
          AND i.STATUS <> 'C'
          AND {EXCLUI_PEDIDOS_EXPORT}
        GROUP BY i.MATERIAL, i.DESCRICAO, i.UNIDADE, t.DESCRICAO
        ORDER BY faturamento_total DESC
    """, {"ini": data_ini, "fim": data_fim})


# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
# QUERY 6 â€” Carteira de pedidos em aberto (peÃ§as)
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
def carteira_abertos() -> pd.DataFrame:
    return sql(f"""
        SELECT
            p.CODIGO                                AS pedido,
            p.DTPEDIDO                              AS dt_pedido,
            p.DTENTREGA                             AS dt_entrega,
            DATEDIFF(DAY, p.DTENTREGA, GETDATE())   AS dias_atraso,
            c.RAZAO                                 AS cliente,
            c.CIDADE                                AS cidade,
            v.RAZAO                                 AS vendedor,
            p.STATUS                                AS status_pedido,
            p.NUMINTERNO                            AS num_interno,
            i.MATERIAL                              AS cod_material,
            i.DESCRICAO                             AS descricao,
            t.DESCRICAO                             AS tipo_venda,
            i.QTDE                                  AS qtde_pedida,
            i.QTDEFAT                               AS qtde_faturada,
            i.QTDE - i.QTDEFAT                      AS saldo_entregar,
            i.VLRUNITARIO                           AS vlr_unitario,
            (i.QTDE - i.QTDEFAT) * i.VLRUNITARIO   AS vlr_saldo,
            i.DTPRAZO                               AS dt_prazo_item
        FROM VE_PEDIDOITENS i
        JOIN VE_PEDIDO p              ON p.CODIGO = i.PEDIDO
        LEFT JOIN FN_FORNECEDORES c   ON c.CODIGO = p.CLIENTE
        LEFT JOIN FN_VENDEDORES v     ON v.CODIGO = p.VENDEDOR
        LEFT JOIN VE_TPVENDA t        ON t.CODIGO = i.TPVENDA
        WHERE p.STATUS NOT IN ('C', 'F')
          AND i.TPVENDA NOT IN ({TPVENDA_SQL})
          AND i.MATERIAL NOT LIKE '8%'
          AND i.FLAGSUB <> 'S'
          AND i.STATUS <> 'C'
          AND {EXCLUI_PEDIDOS_EXPORT}
          AND i.QTDE > i.QTDEFAT
        ORDER BY dias_atraso DESC, p.DTENTREGA, p.CODIGO, i.SEQ
    """)


# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
# QUERY 7 â€” de qual OrÃ§amento veio cada Pedido
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
def origem_orcamento(data_ini: str, data_fim: str) -> pd.DataFrame:
    """
    Mostra o orÃ§amento de origem de cada pedido do perÃ­odo.
    Apenas pedidos que vieram de um orÃ§amento (PEDORIGEM preenchido).
    """
    return sql(f"""
        SELECT
            p.DTPEDIDO              AS dt_pedido,
            p.CODIGO                AS pedido,
            p.NUMINTERNO            AS num_interno,
            p.STATUS                AS status_pedido,
            c.RAZAO                 AS cliente,
            v.RAZAO                 AS vendedor,
            p.VLRTOTAL              AS vlr_pedido,
            o.CODIGO                AS cod_orcamento,
            o.NUMERO                AS num_orcamento,
            o.DTCADASTRO            AS dt_orcamento,
            o.DTVALIDADE            AS dt_validade_orc,
            o.STATUS                AS status_orcamento,
            o.VLRORCADO             AS vlr_orcado,
            o.VLREFETIVO            AS vlr_efetivo,
            o.DESCRICAO             AS descricao_orc
        FROM VE_PEDIDO p
        JOIN VE_ORCAMENTOS o          ON o.CODIGO = p.PEDORIGEM
        LEFT JOIN FN_FORNECEDORES c   ON c.CODIGO = p.CLIENTE
        LEFT JOIN FN_VENDEDORES v     ON v.CODIGO = p.VENDEDOR
        WHERE p.DTPEDIDO BETWEEN :ini AND :fim
          AND p.STATUS <> 'C'          AND {EXCLUI_PEDIDOS_EXPORT}        ORDER BY p.DTPEDIDO DESC, p.CODIGO DESC
    """, {"ini": data_ini, "fim": data_fim})


# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
# QUERY 8 â€” NFs emitidas a partir dos pedidos do perÃ­odo
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
def nfs_emitidas(data_ini: str, data_fim: str) -> pd.DataFrame:
    """
    NFs de saÃ­da vinculadas a pedidos do perÃ­odo.
    Filtro de peÃ§as aplicado nos itens.
    """
    return sql(f"""
        SELECT
            p.DTPEDIDO              AS dt_pedido,
            p.CODIGO                AS pedido,
            p.NUMINTERNO            AS num_interno,
            c.RAZAO                 AS cliente,
            v.RAZAO                 AS vendedor,
            nf.NRONOTA              AS nro_nf,
            nf.DTEMISSAO            AS dt_emissao_nf,
            nf.DTSAIDA              AS dt_saida_nf,
            nf.STATUSNF             AS status_nf,
            nf.VLRPRODUTO           AS vlr_produto_nf,
            nf.VLRIPI               AS vlr_ipi_nf,
            nf.VLRICMSNF            AS vlr_icms_nf,
            nf.VLRTOTAL             AS vlr_total_nf
        FROM FN_NFS nf
        JOIN VE_PEDIDO p              ON p.CODIGO = nf.VEPEDIDO
        LEFT JOIN FN_FORNECEDORES c   ON c.CODIGO = p.CLIENTE
        LEFT JOIN FN_VENDEDORES v     ON v.CODIGO = p.VENDEDOR
        WHERE p.DTPEDIDO BETWEEN :ini AND :fim
          AND p.STATUS <> 'C'
          AND nf.STATUSNF <> 'C'
          AND EXISTS (
              SELECT 1 FROM VE_PEDIDOITENS i
              WHERE i.PEDIDO = p.CODIGO
                AND i.TPVENDA NOT IN ({TPVENDA_SQL})
                AND i.MATERIAL NOT LIKE '8%'
                AND i.FLAGSUB <> 'S'
                AND i.STATUS <> 'C'
          )
          AND {EXCLUI_PEDIDOS_EXPORT}
        ORDER BY nf.DTEMISSAO DESC, p.CODIGO
    """, {"ini": data_ini, "fim": data_fim})


# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
# QUERY 9 â€” Comparativo vs mÃªs anterior (peÃ§as)
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
def comparativo_mensal() -> pd.DataFrame:
    return sql(f"""
        SELECT
            YEAR(p.DTPEDIDO)            AS ano,
            MONTH(p.DTPEDIDO)           AS mes,
            DATENAME(MONTH, p.DTPEDIDO) AS mes_nome,
            COUNT(DISTINCT p.CODIGO)    AS qtd_pedidos,
            SUM(i.VLRTOTAL)             AS faturamento_pecas
        FROM VE_PEDIDOITENS i
        JOIN VE_PEDIDO p ON p.CODIGO = i.PEDIDO
        WHERE YEAR(p.DTPEDIDO) IN (YEAR(GETDATE()), YEAR(GETDATE())-1)
          AND p.STATUS <> 'C'
          AND i.TPVENDA NOT IN ({TPVENDA_SQL})
          AND i.MATERIAL NOT LIKE '8%'
          AND i.FLAGSUB <> 'S'
          AND i.STATUS <> 'C'
          AND {EXCLUI_PEDIDOS_EXPORT}
        GROUP BY YEAR(p.DTPEDIDO), MONTH(p.DTPEDIDO), DATENAME(MONTH, p.DTPEDIDO)
        ORDER BY ano, mes
    """)


# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
# EXECUCAO â€” gera relatorio Excel com todas as abas
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
if __name__ == "__main__":
    hoje = datetime.now()

    # Periodo: mes atual por padrao â€” mude aqui se quiser outro periodo
    DATA_INI = hoje.replace(day=1).strftime("%Y-%m-%d")
    DATA_FIM = hoje.strftime("%Y-%m-%d")
    MES_ANO  = hoje.strftime("%m-%Y")

    print(f"Relatorio de Vendas PECAS â€” {MES_ANO}")
    print(f"Periodo: {DATA_INI} a {DATA_FIM}")
    print(f"Filtros: sem bombas (mat. iniciados em '8') | apenas vendas reais | sem cancelados\n")

    abas = {
        "Itens Pedidos":     lambda: itens_vendas_pecas(DATA_INI, DATA_FIM),
        "Pedidos Resumo":    lambda: pedidos_resumo(DATA_INI, DATA_FIM),
        "Fat. por Vendedor": lambda: faturamento_por_vendedor(DATA_INI, DATA_FIM),
        "Ranking Clientes":  lambda: ranking_clientes(DATA_INI, DATA_FIM),
        "Prod. + Vendidos":  lambda: produtos_mais_vendidos(DATA_INI, DATA_FIM),
        "Carteira Abertos":  lambda: carteira_abertos(),
        "Origem Orcamento":  lambda: origem_orcamento(DATA_INI, DATA_FIM),
        "NFs Emitidas":      lambda: nfs_emitidas(DATA_INI, DATA_FIM),
        "Comp. Mensal":      lambda: comparativo_mensal(),
    }

    caminho = OUTPUT / f"relatorio_pecas_{MES_ANO}_{ts}.xlsx"

    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        for nome_aba, fn in abas.items():
            print(f"  Gerando '{nome_aba}'...", end=" ", flush=True)
            try:
                df = fn()
                df.to_excel(writer, sheet_name=nome_aba[:31], index=False)
                ajustar_colunas(writer, nome_aba[:31], df)
                print(f"{len(df):,} linhas")
            except Exception as e:
                print(f"ERRO: {e}")

    print(f"\nSalvo em: {caminho}")
