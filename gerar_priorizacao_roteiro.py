from datetime import datetime
from pathlib import Path

import pandas as pd

EXPORTS = Path("exports")


def get_input_file() -> Path:
    files = sorted(EXPORTS.glob("analise_roteiro_mp_*.xlsx"), key=lambda p: p.stat().st_mtime)
    if not files:
        raise FileNotFoundError("Nenhum arquivo analise_roteiro_mp_*.xlsx encontrado em exports/")
    return files[-1]


def main():
    input_file = get_input_file()

    df = pd.read_excel(input_file, sheet_name="Roteiro_<2025")

    if df.empty:
        print("Sem dados de roteiro para priorizar.")
        return

    df["Dt_Roteiro_Ref"] = pd.to_datetime(df["Dt_Roteiro_Ref"], errors="coerce")
    df["Dias_Desde_Roteiro"] = pd.to_numeric(df["Dias_Desde_Roteiro"], errors="coerce")

    detalhe = df.sort_values(
        by=["Dias_Desde_Roteiro", "Dt_Roteiro_Ref", "Desenho_OP", "OP"],
        ascending=[False, True, True, True],
    ).reset_index(drop=True)

    consolidado = (
        df.groupby(["Desenho_OP", "Descricao_OP"], dropna=False)
        .agg(
            Qtd_OPs=("OP", "nunique"),
            Menor_Dt_Roteiro=("Dt_Roteiro_Ref", "min"),
            Maior_Dias_Desde=("Dias_Desde_Roteiro", "max"),
            Menor_NROOP=("NROOP", "min"),
        )
        .reset_index()
        .sort_values(by=["Maior_Dias_Desde", "Qtd_OPs", "Desenho_OP"], ascending=[False, False, True])
        .reset_index(drop=True)
    )

    top_criticos = detalhe.head(100).copy()

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = EXPORTS / f"priorizacao_roteiros_{ts}.xlsx"

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        detalhe.to_excel(writer, sheet_name="Roteiro_Detalhe", index=False)
        consolidado.to_excel(writer, sheet_name="Roteiro_Por_Desenho", index=False)
        top_criticos.to_excel(writer, sheet_name="Top100_Criticos", index=False)

    print("Priorização gerada")
    print(f"Entrada: {input_file}")
    print(f"Saída: {output_file}")
    print(f"Detalhe (linhas): {len(detalhe)}")
    print(f"Consolidado desenhos: {len(consolidado)}")
    print(f"Top críticos: {len(top_criticos)}")


if __name__ == "__main__":
    main()
