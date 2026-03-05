"""
explorar_banco.py — Ponto de entrada para exploração do banco Sectra ERP.

Execute: python explorar_banco.py
"""

from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich import print as rprint
import pyodbc
from src.conexao import testar_conexao, listar_drivers
from src.explorador import (
    listar_tabelas,
    descrever_tabela,
    listar_chaves_estrangeiras,
    listar_indices,
    resumo_banco,
    buscar_coluna,
    preview_tabela,
)

console = Console()


def imprimir_df(df, titulo: str = ""):
    """Imprime um DataFrame como tabela formatada no terminal."""
    if titulo:
        console.print(f"\n[bold cyan]{titulo}[/bold cyan]")
    if df.empty:
        console.print("[yellow]  (sem resultados)[/yellow]")
        return

    table = Table(show_header=True, header_style="bold magenta", border_style="dim")
    for col in df.columns:
        table.add_column(str(col), overflow="fold")
    for _, row in df.iterrows():
        table.add_row(*[str(v) if v is not None else "" for v in row])
    console.print(table)


def menu():
    console.print(Panel.fit(
        "[bold green]Explorador do Banco Sectra ERP[/bold green]\n"
        "Conectado via SQL Server (somente leitura)",
        border_style="green",
    ))

    opcoes = {
        "1": "Testar conexão",
        "2": "Listar drivers ODBC disponíveis",
        "3": "Resumo do banco (tipos de objetos)",
        "4": "Listar todas as tabelas",
        "5": "Descrever colunas de uma tabela",
        "6": "Preview de uma tabela (primeiras linhas)",
        "7": "Buscar coluna por nome",
        "8": "Listar chaves estrangeiras (relacionamentos)",
        "9": "Listar índices",
        "0": "Sair",
    }

    while True:
        console.print("\n[bold]Menu:[/bold]")
        for k, v in opcoes.items():
            console.print(f"  [cyan]{k}[/cyan] — {v}")

        escolha = console.input("\n[bold yellow]Opção:[/bold yellow] ").strip()

        if escolha == "0":
            break

        elif escolha == "1":
            testar_conexao()

        elif escolha == "2":
            drivers = listar_drivers()
            console.print("\n[bold]Drivers ODBC instalados:[/bold]")
            for d in drivers:
                console.print(f"  • {d}")

        elif escolha == "3":
            df = resumo_banco()
            imprimir_df(df, "Resumo do banco")

        elif escolha == "4":
            schema = console.input("Schema (Enter para todos): ").strip() or None
            df = listar_tabelas(schema)
            imprimir_df(df, f"Tabelas — {len(df)} objetos encontrados")

        elif escolha == "5":
            tabela = console.input("Nome da tabela: ").strip()
            schema = console.input("Schema (Enter para 'dbo'): ").strip() or "dbo"
            df = descrever_tabela(tabela, schema)
            imprimir_df(df, f"Colunas de [{schema}].[{tabela}]")

        elif escolha == "6":
            tabela = console.input("Nome da tabela: ").strip()
            schema = console.input("Schema (Enter para 'dbo'): ").strip() or "dbo"
            n = console.input("Quantas linhas? (Enter para 10): ").strip()
            n = int(n) if n.isdigit() else 10
            df = preview_tabela(tabela, schema, n)
            imprimir_df(df, f"Preview [{schema}].[{tabela}] — {len(df)} linhas")

        elif escolha == "7":
            nome = console.input("Parte do nome da coluna: ").strip()
            df = buscar_coluna(nome)
            imprimir_df(df, f"Tabelas com coluna contendo '{nome}'")

        elif escolha == "8":
            df = listar_chaves_estrangeiras()
            imprimir_df(df, f"Chaves Estrangeiras — {len(df)} relacionamentos")

        elif escolha == "9":
            tabela = console.input("Tabela específica (Enter para todas): ").strip() or None
            df = listar_indices(tabela)
            imprimir_df(df, "Índices")

        else:
            console.print("[red]Opção inválida.[/red]")

    console.print("\n[green]Até logo![/green]")


if __name__ == "__main__":
    menu()
