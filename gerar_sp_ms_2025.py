from datetime import datetime
from pathlib import Path

import pandas as pd


def main():
    src = Path("exports/relatorio_cobranca_vendedores_20260304_181943.xlsx")
    if not src.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {src}")

    df = pd.read_excel(src, sheet_name="Detalhado")
    df = df[df["UF"].isin(["SP", "MS"])].copy()

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = Path(f"exports/relatorio_cobranca_SP_MS_2025_{ts}.xlsx")

    resumo = (
        df.groupby(["UF"], dropna=False)
        .agg(Titulos=("NF", "count"), Clientes=("Cliente", "nunique"), Valor_Total=("Valor", "sum"))
        .reset_index()
    )

    resumo_cidade = (
        df.groupby(["UF", "Cidade"], dropna=False)
        .agg(Titulos=("NF", "count"), Clientes=("Cliente", "nunique"), Valor_Total=("Valor", "sum"))
        .reset_index()
        .sort_values(["UF", "Valor_Total"], ascending=[True, False])
    )

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        resumo.to_excel(writer, sheet_name="Resumo_UF", index=False)
        resumo_cidade.to_excel(writer, sheet_name="Resumo_Cidade", index=False)
        df.sort_values(["UF", "Cidade", "Vencimento", "Cliente"]).to_excel(
            writer, sheet_name="Detalhado", index=False
        )

    print(f"Arquivo: {out}")
    print(f"Linhas: {len(df)}")
    print(resumo.to_string(index=False))


if __name__ == "__main__":
    main()
