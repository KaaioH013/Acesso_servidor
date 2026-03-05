# Backlog Priorizado — Painel Comercial (Coordenação)

## Status validado hoje

- Dashboard gera normalmente: `exports/dashboard.html`
- Ranking de externos está territorial (internos fora da disputa)
- Exceções MG do Alexandre estão aplicadas via `Cidades_nao_atendidas_mg_alexandre.xlsx`
- Blocos operacionais ativos: Prioridades do Dia + Cotações Geradas

---

## Próximas evoluções lógicas (ordem sugerida)

### 1) Funil de cotações no mês fechado (alta prioridade)
**Objetivo:** enxergar conversão rapidamente (geradas → abertas → aprovadas → perdidas).

**Entregas:**
- KPI de taxa de conversão (qtde e valor)
- Gráfico funil por status
- Corte por vendedor externo territorial

### 2) SLA operacional diário (alta prioridade)
**Objetivo:** controlar execução da equipe ao longo do dia.

**Entregas:**
- Placar de prioridades: entrada do dia vs resolvidas do dia
- Tempo médio de permanência em `STATUS='L'`
- Tempo médio de atraso por vendedor/território

### 3) Mapa territorial auditável (alta prioridade)
**Objetivo:** reduzir discussão de “de quem é o cliente”.

**Entregas:**
- Aba/visão de auditoria territorial por pedido
- Motivo da regra aplicada (UF, cidade excluída, data de vigência)
- Export rápido para alinhamento com controller/comercial

### 4) Detecção de risco de meta (média prioridade)
**Objetivo:** antecipar mês ruim antes de fechar.

**Entregas:**
- Projeção simples de fechamento do mês (run-rate)
- Gap para meta por território
- Lista de top oportunidades de recuperação

### 5) Fila comercial por responsável (média prioridade)
**Objetivo:** transformar dashboard em rotina de gestão.

**Entregas:**
- Tabela “próximas 20 ações” por vendedor externo
- Priorização por valor + criticidade + prazo
- Filtro por território

---

## Melhorias UI/UX (ordem sugerida)

### A) Modo Operação (quick win)
- Reorganizar tela para colocar primeiro: KPIs críticos + Prioridades + Cotações
- Deixar gráficos estratégicos (ABC/YoY) abaixo

### B) Hierarquia visual de decisão (quick win)
- Cores mais consistentes por tipo de alerta
- Labels curtos e orientados a ação
- Destaque do “o que fazer agora” no topo

### C) Navegação por blocos (quick win)
- Índice no topo com âncoras: Prioridades, Cotações, Carteira, Ranking
- Facilita uso em reunião e operação diária

### D) Legibilidade de tabelas (quick win)
- Colunas fixas para Vendedor/Cliente
- Densidade de linhas equilibrada
- Formatação numérica padronizada (R$, dias, %)

### E) Modo Reunião x Modo Operação (médio prazo)
- Reunião: highlights e indicadores consolidados
- Operação: listas de ação e SLA

---

## Critério de sucesso para coordenação de vendas

1. Em menos de 3 minutos, identificar os 10 itens mais críticos do dia.
2. Saber exatamente quem acionar e por quê.
3. Medir no fim do dia quanto da fila crítica foi resolvida.
4. Reduzir itens `STATUS='L'` envelhecidos e atrasos recorrentes.
