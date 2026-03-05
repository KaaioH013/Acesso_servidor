"""
exportar.py — Funções para exportar dados para CSV, Excel e múltiplas abas.
"""

import os
import pandas as pd
from datetime import datetime
from pathlib import Path
from sqlalchemy import text
from src.conexao import get_engine


OUTPUT_DIR = Path("exports")
OUTPUT_DIR.mkdir(exist_ok=True)


def _timestamp():
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def query_para_dataframe(query: str, params: dict = None) -> pd.DataFrame:
    """Executa uma query SQL e retorna um DataFrame."""
    engine = get_engine()
    with engine.connect() as conn:
        return pd.read_sql(text(query), conn, params=params or {})


def exportar_csv(query: str, nome_arquivo: str, params: dict = None, sep: str = ";") -> Path:
    """
    Executa a query e salva o resultado como CSV.

    Parâmetros
    ----------
    query        : SQL SELECT
    nome_arquivo : Nome base do arquivo (sem extensão)
    params       : Dicionário de parâmetros para a query
    sep          : Separador (padrão ';' para abrir no Excel BR)
    """
    df = query_para_dataframe(query, params)
    caminho = OUTPUT_DIR / f"{nome_arquivo}_{_timestamp()}.csv"
    df.to_csv(caminho, index=False, sep=sep, encoding="utf-8-sig")
    print(f"✅ CSV exportado: {caminho}  ({len(df):,} linhas)")
    return caminho


def exportar_excel(query: str, nome_arquivo: str, nome_aba: str = "Dados", params: dict = None) -> Path:
    """
    Executa a query e salva o resultado como Excel (.xlsx).

    Parâmetros
    ----------
    query        : SQL SELECT
    nome_arquivo : Nome base do arquivo (sem extensão)
    nome_aba     : Nome da aba no Excel
    params       : Dicionário de parâmetros para a query
    """
    df = query_para_dataframe(query, params)
    caminho = OUTPUT_DIR / f"{nome_arquivo}_{_timestamp()}.xlsx"

    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=nome_aba, index=False)
        _auto_ajustar_colunas(writer, nome_aba, df)

    print(f"✅ Excel exportado: {caminho}  ({len(df):,} linhas)")
    return caminho


def exportar_excel_multi_abas(abas: dict[str, str], nome_arquivo: str) -> Path:
    """
    Exporta múltiplas queries para abas diferentes em um único Excel.

    Parâmetros
    ----------
    abas         : Dicionário { 'Nome da Aba': 'SELECT ...' }
    nome_arquivo : Nome base do arquivo (sem extensão)

    Exemplo
    -------
    exportar_excel_multi_abas(
        abas={
            'Clientes': 'SELECT * FROM dbo.Clientes',
            'Pedidos':  'SELECT * FROM dbo.Pedidos WHERE ano = 2025',
        },
        nome_arquivo='relatorio_geral'
    )
    """
    caminho = OUTPUT_DIR / f"{nome_arquivo}_{_timestamp()}.xlsx"
    engine = get_engine()

    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        for nome_aba, query in abas.items():
            with engine.connect() as conn:
                df = pd.read_sql(text(query), conn)
            df.to_excel(writer, sheet_name=nome_aba[:31], index=False)  # Excel limita 31 chars
            _auto_ajustar_colunas(writer, nome_aba[:31], df)
            print(f"   ✔ Aba '{nome_aba}': {len(df):,} linhas")

    print(f"✅ Excel multi-abas exportado: {caminho}")
    return caminho


def exportar_tabela_completa(tabela: str, schema: str = "dbo", formato: str = "excel") -> Path:
    """
    Exporta uma tabela inteira para CSV ou Excel.

    Parâmetros
    ----------
    tabela  : Nome da tabela
    schema  : Schema (padrão: dbo)
    formato : 'csv' ou 'excel'
    """
    query = f"SELECT * FROM [{schema}].[{tabela}]"
    nome = f"{schema}_{tabela}"
    if formato == "csv":
        return exportar_csv(query, nome)
    else:
        return exportar_excel(query, nome, nome_aba=tabela)


def _auto_ajustar_colunas(writer, nome_aba: str, df: pd.DataFrame):
    """Ajusta automaticamente a largura das colunas no Excel."""
    worksheet = writer.sheets[nome_aba]
    for i, col in enumerate(df.columns, 1):
        max_len = max(
            len(str(col)),
            df[col].astype(str).str.len().max() if len(df) > 0 else 0,
        )
        worksheet.column_dimensions[
            worksheet.cell(1, i).column_letter
        ].width = min(max_len + 4, 60)
