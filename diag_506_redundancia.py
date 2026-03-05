from pathlib import Path
import pandas as pd

exports = Path('exports')
files = sorted(exports.glob('relatorio_506_excel_*.xlsx'), key=lambda p: p.stat().st_mtime, reverse=True)
if not files:
    raise SystemExit('Nenhum relatorio_506_excel_*.xlsx encontrado em exports/')

f = files[0]
df = pd.read_excel(f, sheet_name='506_Melhorado')

ndf = df.copy()
for c in ndf.columns:
    if str(ndf[c].dtype).startswith('datetime'):
        ndf[c] = pd.to_datetime(ndf[c], errors='coerce').dt.strftime('%Y-%m-%d').fillna('')
    else:
        ndf[c] = ndf[c].fillna('').astype(str).str.strip()

cols = list(ndf.columns)
dups = []
for i, a in enumerate(cols):
    for b in cols[i + 1 :]:
        if (ndf[a] == ndf[b]).all():
            dups.append((a, b))

lines = []
lines.append(f'FILE={f}')
lines.append(f'ROWS={len(df)} COLS={len(df.columns)}')
lines.append(f'EXACT_DUP_PAIRS={dups}')

rule_p = ((df['P'].fillna('') == 'S') == (df['Status_Pagamento'].fillna('') == 'PAGA')).all()
rule_parcelas = (df['Parcelas_Total'].fillna(0) == (df['Parcelas_Pagas'].fillna(0) + df['Parcelas_Abertas'].fillna(0))).all()
rule_nfs = (df['NFS'].fillna('') == df['NRONOTA'].fillna('')).all()

lines.append(f'RULE_P_vs_STATUS={rule_p}')
lines.append(f'RULE_PARCELAS_TOTAL={rule_parcelas}')
lines.append(f'RULE_NFS_EQ_NRONOTA={rule_nfs}')

for c in [
    'Cod_Cliente',
    'Razao_Social',
    'Cod_Vendedor_NF',
    'Vendedor_NF',
    'NFS',
    'NRONOTA',
    'Receber',
    'NOTASEQ',
    'VEPEDIDO',
    'P',
    'Status_Pagamento',
]:
    if c in df.columns:
        lines.append(f'UNIQ::{c}={df[c].nunique(dropna=False)}/{len(df)}')

for c in ['Vendedor_Interno', 'Vendedor_Externo', 'VEPEDIDO', 'Dt_Pagto', 'Vlr_Desc', 'Qtd_Itens_Bomba', 'Qtd_Itens_Peca']:
    if c in df.columns:
        nr = float(df[c].isna().mean())
        lines.append(f'NULL::{c}={nr:.4f}')

out = exports / '_diag_506_cols.txt'
out.write_text('\n'.join(lines), encoding='utf-8')
print(str(out))
