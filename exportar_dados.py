"""
exportar_dados.py — Exemplos prontos de exportação de dados do Sectra ERP.

Edite os blocos abaixo conforme suas necessidades e execute:
    python exportar_dados.py
"""

from src.exportar import (
    exportar_csv,
    exportar_excel,
    exportar_excel_multi_abas,
    exportar_tabela_completa,
    query_para_dataframe,
)


# ─────────────────────────────────────────────────────────────
# EXEMPLO 1 — Exportar uma tabela inteira para Excel
# ─────────────────────────────────────────────────────────────
def exemplo_tabela_completa():
    exportar_tabela_completa(
        tabela="NomeDaTabela",   # ← substitua pelo nome real
        schema="dbo",
        formato="excel",         # 'excel' ou 'csv'
    )


# ─────────────────────────────────────────────────────────────
# EXEMPLO 2 — Exportar com query personalizada → CSV
# ─────────────────────────────────────────────────────────────
def exemplo_query_csv():
    query = """
        SELECT TOP 1000
            id,
            nome,
            data_cadastro,
            status
        FROM dbo.Clientes
        WHERE status = 'ATIVO'
        ORDER BY data_cadastro DESC
    """
    exportar_csv(query, nome_arquivo="clientes_ativos")


# ─────────────────────────────────────────────────────────────
# EXEMPLO 3 — Exportar com parâmetros (evita SQL injection)
# ─────────────────────────────────────────────────────────────
def exemplo_com_parametros():
    query = """
        SELECT *
        FROM dbo.Pedidos
        WHERE YEAR(data_pedido) = :ano
          AND status = :status
    """
    exportar_excel(
        query,
        nome_arquivo="pedidos_2025",
        nome_aba="Pedidos 2025",
        params={"ano": 2025, "status": "FATURADO"},
    )


# ─────────────────────────────────────────────────────────────
# EXEMPLO 4 — Relatório multi-abas em um único Excel
# ─────────────────────────────────────────────────────────────
def exemplo_relatorio_multi_abas():
    exportar_excel_multi_abas(
        abas={
            "Clientes":  "SELECT * FROM dbo.Clientes",
            "Pedidos":   "SELECT TOP 5000 * FROM dbo.Pedidos ORDER BY data_pedido DESC",
            "Produtos":  "SELECT * FROM dbo.Produtos WHERE ativo = 1",
        },
        nome_arquivo="relatorio_geral_erp",
    )


# ─────────────────────────────────────────────────────────────
# EXEMPLO 5 — Só ler os dados (sem exportar) para analisar
# ─────────────────────────────────────────────────────────────
def exemplo_analisar_no_python():
    df = query_para_dataframe("SELECT * FROM dbo.Pedidos WHERE YEAR(data_pedido) = 2025")

    print(f"\nTotal de registros: {len(df):,}")
    print(f"\nColunas: {df.columns.tolist()}")
    print(f"\nTipos:\n{df.dtypes}")
    print(f"\nEstatísticas:\n{df.describe()}")

    # Exemplo de agrupamento
    # resumo = df.groupby('status')['valor_total'].agg(['sum', 'count', 'mean'])
    # print(resumo)


# ─────────────────────────────────────────────────────────────
# Execute o que precisar (descomente as linhas abaixo)
# ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    # exemplo_tabela_completa()
    # exemplo_query_csv()
    # exemplo_com_parametros()
    # exemplo_relatorio_multi_abas()
    # exemplo_analisar_no_python()

    print("Descomente um dos exemplos acima para executar.")
    print("Arquivos exportados ficam na pasta: exports/")
