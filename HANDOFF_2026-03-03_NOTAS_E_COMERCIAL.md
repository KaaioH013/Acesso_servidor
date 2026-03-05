# Handoff — 03/03/2026

## Status consolidado (encerramento frente de Notas)

A frente de **Notas / Relatório 506** foi encerrada para apresentação ao controller.

### Script oficial
- `relatorio_506_excel.py`

### Arquivo final de referência
- `exports/relatorio_506_excel_20260303_130626.xlsx`

### O que está garantido neste arquivo
- Aba principal: `506_Melhorado`
- Aba de auditoria: `Validacao`
- `Dt_Emissao` vem de `FN_NFS.DTEMISSAO` (emissão real da NF)
- `Dt_Emissao_Titulo` vem de `FN_RECEBER.DTEMISSAO` (lançamento financeiro)
- Quitação por NF:
  - `Status_NF_Quitada`
  - `Parcelas_Total`
  - `Parcelas_Pagas`
  - `Parcelas_Abertas`
  - `Dt_Ultimo_Venc_NF`
- Classificação de tipo na mesma planilha:
  - `Tipo_Produto = PECA | BOMBA | AMBAS`
  - sem duplicação de título (`Receber`)

### Validação cruzada (Excel vs SQL)
- A aba `Validacao` compara por tipo e total:
  - Títulos
  - NFs
  - Valor devido
  - Valor recebido
- Status esperado para apresentação: `OK` em todas as linhas.

## Comandos úteis

```powershell
# Gerar 506 validado
py -3 .\relatorio_506_excel.py --dt-ini 2024-01-01 --dt-fim 2026-03-03
```

## Caso crítico já validado
- NF `35851`
  - Emissão NF: `28/06/2024`
  - Último vencimento: `23/08/2024`
  - 8 parcelas, 1 em aberto
  - Status: `PENDENTE`

## Próxima frente (retomar agora)

Evolução da parte comercial com foco em **comissões**:
1. Parametrizar regras por estado/cidade/vendedor/tipo.
2. Aplicar regras sobre base financeira quitada.
3. Validar valores de comissão por competência e por vendedor.
4. Usar o 506 melhorado como base de conferência operacional.
