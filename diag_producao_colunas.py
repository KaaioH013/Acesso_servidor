import pandas as pd
from src.conexao import get_engine

engine = get_engine()

tabelas = [
    'PR_OP',
    'PR_OPSTATUS',
    'PR_OPROTEIRO',
    'PR_OPROTEIROITEM',
    'PR_OPMATERIAIS',
    'PR_OPMATERIALITENS',
    'PR_DESENHOROTEIRO',
    'PR_DESENHOROTITEM',
    'PR_ROTEIROPADRAO',
    'PR_ROTEIROPADRAOITENS',
    'PR_MATERIALBASE',
    'MT_MOVIMENTACAO',
    'MT_MATERIAL',
]

for t in tabelas:
    print(f"\n=== {t} ===")
    q_cols = f"""
    SELECT COLUMN_NAME, DATA_TYPE
    FROM INFORMATION_SCHEMA.COLUMNS
    WHERE TABLE_NAME = '{t}'
    ORDER BY ORDINAL_POSITION
    """
    df = pd.read_sql(q_cols, engine)
    if df.empty:
        print('Tabela não encontrada')
        continue
    print(df.to_string(index=False))

    # preview simples para entender chaves e datas
    q_preview = f"SELECT TOP 3 * FROM {t}"
    try:
        d = pd.read_sql(q_preview, engine)
        print('\nPreview (3 linhas):')
        print(d.head(3).to_string(index=False, max_cols=12))
    except Exception as exc:
        print(f'Preview indisponível: {exc}')
