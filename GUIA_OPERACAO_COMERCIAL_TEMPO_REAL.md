# Guia rápido — Operação Comercial em Tempo Real

## Atualização do painel

```powershell
py -3 .\fase4_dashboard.py --no-open
```

Arquivo principal:
- `exports/dashboard.html`

## Como usar no dia a dia (coordenação)

1. Abra o bloco **🎯 Prioridades do Dia — Ação Comercial**.
2. Trabalhe de cima para baixo (ordem por criticidade + impacto financeiro).
3. Foco inicial:
   - `NF CRITICA` → cobrar emissão imediata
   - `PRAZO ATRASADO` → replanejar entrega + alinhamento com cliente
   - `NF ATENCAO` / `PRAZO URGENTE` → organizar fila da semana
4. Feche o dia revisando:
   - `🔴 Atrasados`
   - `🟡 Urgentes`
   - `📄 Aguardando NF`

## Regras da nova priorização

- **NF CRITICA**: item `STATUS='L'` com `Dias_Aguard_NF >= 10`
- **PRAZO ATRASADO**: semáforo da carteira em atraso
- **NF ATENCAO**: item `STATUS='L'` sem atingir crítico
- **PRAZO URGENTE**: prazo até 7 dias

## Rotina recomendada

- Gerar dashboard no início da manhã e após almoço.
- Usar `dashboard_base.html` e `dashboard_sem_contrato.html` para visão comparativa quando necessário.
