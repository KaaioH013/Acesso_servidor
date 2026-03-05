# Roadmap — Sistema de Inteligência Comercial Helibombas

> Acesso leitura SQL Server • ERP Sectra • Coordenação de Vendas (Peças)
> Atualizado: Março/2026 — Fases 1, 2, 3 e 4 concluídas ✅ + Evoluções operacionais da coordenação

---

## Visão Geral

O ERP mostra **o que aconteceu**. Este sistema vai mostrar **o que está acontecendo e o que vai acontecer** — em tempo real, com contexto e com ação sugerida.

---

## ✅ FASE 1 — Relatórios Mensais Avançados — CONCLUÍDA
> Script: `fase1_comparativos.py` — `python fase1_comparativos.py --ano 2026`

### 1.1 Comparativos consolidados
- [x] Evolução mensal dos últimos 24 meses (YoY embutido)
- [x] Ranking de clientes: ano atual vs anterior com variação %
- [x] Ranking de vendedores: comparativo YoY por valor e pedidos

### 1.2 Taxa de conversão de orçamentos
- [x] Por mês: total orçamentos, aprovados, perdidos, % conversão
- [x] Valor orçado vs valor convertido por período

### 1.3 Curva ABC
- [x] Curva ABC clientes (A=80%, B=15%, C=5%) com coloração no Excel
- [x] Curva ABC materiais com mesmo critério
- [x] Análise de churn: clientes novos / retidos / perdidos entre dois anos

**Excel gerado:** `exports/fase1_analitico_AAAA_TIMESTAMP.xlsx` — 8 abas

> Resultado 2026 (jan-fev, **somente peças** — exclui bombas `MATERIAL LIKE '8%'`): 380 clientes, R$ 4,62M (= R$2,627M jan + R$1,993M fev validados), 100 clientes Curva A
> ⚠️ Baseado em **data do pedido** (não data NF). Relatório de faturamento real por NF emitida = trabalho futuro.

---

## ✅ FASE 2 — Gestão da Carteira em Tempo Real — CONCLUÍDA
> Script: `fase2_carteira.py` — `python fase2_carteira.py`

### 2.1 Painel de pedidos em aberto
- [x] Todos os pedidos com saldo a entregar, por semana de entrega prevista
- [x] Semáforo de prazo: 🟢 no prazo | 🟡 até 7 dias | 🔴 atrasado
- [x] Filtro por vendedor (externo e interno separados)

### 2.2 Itens com status "L" (Liberado)
- [x] Todos os itens liberados aguardando faturamento
- [x] **Há quanto tempo** estão liberados (cálculo de dias desde `DTALTERAFAT`)
- [x] Agrupado por responsável (vendedor interno) para cobrar
- [x] Valor total parado aguardando emissão de NF

### 2.3 OPs — descoberta importante
- [x] ⚠️ `PR_OP` só existe para **bombas completas** (`MATERIAL LIKE '8%'`)
- [x] Peças são compradas/estocadas — sem OP vinculada ao pedido
- [x] Campo: `DTALTERAFAT` = data de liberação para NF

**Excel gerado:** `exports/carteira_DD-MM-AAAA_TIMESTAMP.xlsx` — 5 abas

> Resultado 26/02/2026: **275 itens em aberto** | R$1,53M | 🔴 50 atrasados | 🟡 93 urgentes | 53 aguardando NF (R$167K)

---

## ✅ FASE 3 — Margens e Rentabilidade — CONCLUÍDA
> Script: `fase3_margens.py` — `python fase3_margens.py --mes 2 --ano 2026 [--margem-critica 20]`

### 3.1 Margem por pedido/item
- [x] `VLRCUSTO` vs preço de venda → margem % por item, com faixas de cor
- [x] Cobertura: ~81% dos itens têm `VLRCUSTO` preenchido
- [x] ⚠️ `VLRMATERIAL` **não é custo** — confirmado igual ao preço de venda
- [x] Margem por vendedor, por cliente

### 3.2 Custo de última compra
- [x] `MT_MOVIMENTACAO` com `EVENTO=3` = entrada de compra
- [x] Compara preço venda médio vs última compra → margem real
- [x] Flag 🔴 quando última compra > preço de venda

### 3.3 Desconto vs margem
- [x] `PERCDESCONTO` e `VLRDESCONTO` por vendedor → impacto na margem líquida estimada
- [x] Alerta threshold configurável (padrão: < 20%)

**Excel gerado:** `exports/margens_MM-AAAA_TIMESTAMP.xlsx` — 6 abas

> Resultado fev/2026: 669 itens | margem média **52,7%** | lucro bruto est. **R$974K** | 8 alertas críticos | cobertura VLRCUSTO 81,3%

---

## ✅ FASE 4 — Dashboard Visual Interativo — CONCLUÍDA
> Script: `fase4_dashboard.py` — `python fase4_dashboard.py`

### 4.1 Dashboard HTML local (sem servidor, sem instalação)
- [x] Cards KPI: **Vendas** mês (corrigido — não "Faturamento"), carteira, atrasados, urgentes, aguardando NF, margem média
- [x] Gráfico de barras agrupadas: evolução de vendas mensal YoY **+ linhas de meta** (R$2,2M/2025 · R$2,5M/2026)
- [x] Ranking **Representantes Externos** separado do ranking **Vendedores Internos** (TIPO = E/I)
- [x] Gráfico de **Performance Externos**: barras de faturamento + linha de margem % (duplo eixo)
- [x] Treemap: Curva ABC clientes Curva A/B/C (top 30)
- [x] Rosca: semáforo da carteira (atrasado / urgente / no prazo)
- [x] Tabela: top 10 itens aguardando NF — **corrigido**: usa `DtPedido` quando `DTALTERAFAT` é NULL
- [x] Tabela: alertas de margem crítica (<20%) **com explicação do cálculo** embutida
- [x] Arquivo único `.html` — abre diretamente no navegador

**Metas mensais configuráveis** em `META_MENSAL` no topo de `fase4_dashboard.py`.

### 4.2 Atualização automática
- [x] `--no-open` para rodar sem abrir navegador (Task Scheduler)
- [ ] Windows Task Scheduler: configurar `fase4_dashboard.py --no-open` às 07h00
- [ ] Atalho na área de trabalho para `exports/dashboard.html`

### 4.3 Evoluções operacionais entregues em 03/03/2026
- [x] Modo Operação no topo (ação diária antes dos gráficos estratégicos)
- [x] Bloco Prioridades do Dia (NF crítica, prazo atrasado, urgentes)
- [x] Bloco Cotações Geradas (KPI + tabela operacional)
- [x] Funil comercial de cotações e conversão por território
- [x] Territorialização explícita de externos por UF/cidade/data
- [x] Exceções MG por arquivo externo (`Cidades_nao_atendidas_mg_alexandre.xlsx`)
- [x] Navegação por âncoras no topo (Operação, Cotações, Estratégico)

**HTML gerado:** `exports/dashboard.html` (sempre sobrescreve — versão mais recente)

> Resultado fev/2026 (v2): Dashboard com 6 KPIs, 6 gráficos interativos e 2 tabelas de alertas.
> Ranking externo/interno separados. Meta visualizada no gráfico de evolução. Bug 0 dias NF corrigido.

---

## FASE 5 — Relatório PDF Mensal Automatizado
> *O relatório executivo que vai pro e-mail*

### 5.1 Layout com identidade Helibombas
- [ ] Template com logo, cores e fonte da empresa
- [ ] Capa com mês/ano e resumo executivo (3-4 KPIs principais)
- [ ] Seções: faturamento, carteira, vendedores, clientes, margem

### 5.2 Conteúdo automático
- [ ] Comparativo do mês atual vs mesmo mês ano anterior
- [ ] Top 5 clientes e top 5 materiais do mês
- [ ] Gráficos embutidos no PDF (não é print de tela — gerado programaticamente)
- [ ] Destaques automáticos: "Cliente X cresceu 40% vs jan/25"

### 5.3 Envio automático por e-mail
- [ ] Script envia o PDF todo dia 1 do mês às 08h00
- [ ] Lista de destinatários configurável (gerência, diretoria, vendedores)
- [ ] Assunto e corpo do e-mail gerados automaticamente com os KPIs

**Entrega:** `relatorio_pdf.py` + configuração de Task Scheduler + template Word/PDF.

---

## FASE 6 — Inteligência Comercial
> *Ir além do histórico — prever e recomendar*

### 6.1 Alertas automáticos (diários)
- [ ] Pedidos que vencem essa semana sem NF emitida
- [ ] Orçamentos vencendo em 7 dias sem resposta
- [ ] Clientes sem pedido há X dias (reativação)
- [ ] Itens com status L há mais de N dias

### 6.2 Sazonalidade
- [ ] Análise de quais meses/trimestres historicamente vendem mais por cliente
- [ ] Identificar clientes com pedidos sazonais previsíveis
- [ ] Sugestão: "cliente X normalmente pede em março — ainda não pediu"

### 6.3 Churn e retenção
- [ ] Clientes ativos em 2024 que não compraram em 2025
- [ ] Valor perdido de clientes inativos
- [ ] Clientes com queda de volume > 30% vs ano anterior

**Entrega:** `alertas_diarios.py` → e-mail matinal com lista de ações.

---

## FASE 7 — ERP 360 (Painel Executivo-Operacional Integrado)
> *Dashboard único para saber a empresa inteira, com profundidade por módulo*

### 7.1 Cockpit executivo (1 tela de comando)
- [ ] Receita/Vendas: mês, YoY, run-rate, gap para meta
- [ ] Carteira: valor, atraso, urgência, aging por faixa
- [ ] Financeiro: a receber vencido/a vencer, inadimplência, DSO simplificado
- [ ] Margem: média do mês, risco de erosão, top desvios
- [ ] Operação: fila crítica do dia e SLA de resolução

### 7.2 Drill-down por área (navegação consistente)
- [ ] Comercial (pedidos, cotações, conversão, território)
- [ ] Financeiro (títulos, quitação, atraso, concentração por cliente)
- [ ] Fiscal/Notas (emissão, pendências e reconciliação NF x título)
- [ ] Suprimentos/Estoque (ruptura, cobertura e giro)

### 7.3 Camada de governança de dados
- [ ] Dicionário único de KPIs (definição, fórmula, fonte)
- [ ] Regras de filtro centralizadas (peças, exportação, cancelados, substituição)
- [ ] Trilhas de auditoria por KPI crítico (consulta de conferência pronta)

**Entrega esperada:** `dashboard_erp_360.html` + dicionário de KPIs + pacote de auditoria.

---

## FASE 8 — SLA e Gestão de Execução (Coordenação)
> *Transformar painel em rotina diária com medição de fechamento de fila*

### 8.1 SLA diário (prioridade máxima)
- [x] Entrada do dia x resolvidas do dia (placar v1 — snapshot operacional)
- [ ] Backlog inicial x backlog final
- [ ] Tempo médio de resolução por tipo de prioridade

### 8.1.1 Status de implementação em 03/03/2026
- [x] Bloco `⏱️ SLA Diário — Entrada x Resolução` no dashboard
- [x] Entradas do dia por itens em `STATUS='L'` com liberação no dia
- [x] Resolvidas do dia por itens com NF emitida no dia
- [x] Saldo do dia, backlog atual e taxa de saída
- [ ] Evoluir para SLA completo com histórico diário persistido (entrada real x saída real)

### 8.1.2 Evolução de visualização (03/03/2026)
- [x] Duas abas no mesmo dashboard: `Mês Fechado` e `MTD Hoje`
- [x] Aba MTD com leitura operacional do mês em andamento (hoje, acumulado, run-rate, projeção, gap)
- [ ] Expandir aba MTD com carteira do dia, cotações do dia e ranking intramês em tempo quase real

### 8.2 Gestão por responsável
- [ ] Fila de ação por vendedor/território (top 20 acionáveis)
- [ ] Motivo da prioridade e próximo passo sugerido
- [ ] Export rápido para ritual diário de acompanhamento

### 8.3 Quick wins (alto impacto, baixo esforço)
- [ ] Pipeline de ação diária (Top 20) com ranking por criticidade + valor
- [ ] Cliente em risco (queda relevante vs média 3-6 meses)
- [ ] Concentração de receita (dependência por cliente/território)
- [ ] Card de saúde de dados (nulos críticos e divergências)

### 8.4 Eficiência operacional comercial
- [ ] Tempo médio por etapa da cotação (onde trava)
- [ ] Tempo de primeira resposta comercial
- [ ] Tempo até aprovação/perda
- [ ] SLA por responsável (entrada, resolvido, backlog e aging)

**Entrega esperada:** bloco SLA no dashboard + relatório diário de execução.

---

## FASE 9 — Previsão e Risco de Meta
> *Antecipar fechamento do mês e orientar recuperação*

### 9.1 Forecast simples e robusto
- [ ] Run-rate diário do mês
- [ ] Cenários conservador/base/agressivo
- [ ] Gap por território e por carteira de clientes

### 9.2 Alavancas de recuperação
- [ ] Top oportunidades por valor x probabilidade
- [ ] Lista priorizada de ações para recuperar meta

### 9.3 Risco de meta (camada executiva)
- [ ] Probabilidade de atingir meta por território
- [ ] Gap para meta com explicação de alavancas
- [ ] Cenário semanal atualizado automaticamente

**Entrega esperada:** painel de previsão + plano de recuperação semanal.

---

## FASE 10 — Distribuição e Operação Contínua
> *Garantir que o dashboard rode sozinho, com confiabilidade e histórico*

### 10.1 Automação e publicação
- [x] Agendamento diário (Task Scheduler)
- [x] Agendamento intradiário otimizado (a cada 2h no comercial)
- [ ] Publicação em pasta/rede compartilhada
- [ ] Versionamento diário de snapshots HTML/Excel

### 10.2 Qualidade operacional
- [x] Healthcheck básico por log de execução (OK/FALHA)
- [ ] Log de execução e alerta de erro
- [x] Checklist de validação rápida de KPIs críticos (script `validar_dashboard.py`)

### 10.3 Governança e auditoria de KPI
- [ ] Dicionário de KPI (definição, fórmula, fonte e periodicidade)
- [ ] Consulta de auditoria por KPI sensível (1 clique)
- [ ] Trilhas de divergência e histórico de correções

### Prioridade de execução recomendada (próximas semanas)
1. [ ] **Semana 1**: Pipeline Top 20 + export diário + saúde de dados
2. [ ] **Semana 2**: Cliente em risco + concentração de receita
3. [ ] **Semana 3**: Tempo por etapa da cotação + eficiência comercial
4. [ ] **Semana 4**: Risco de meta por território + alavancas semanais

**Entrega esperada:** operação autônoma diária com observabilidade mínima.

---

## Sequência Sugerida

```
JAN  FEV  MAR  ABR  MAI  JUN
├──────────┤
FASE 1     │ já temos base, 1-2 semanas
           ├──────────┤
           FASE 2     │ carteira em tempo real, 2-3 semanas
                      ├──────────┤
                      FASE 3     │ margens, depende de validar colunas de custo
                                 ├──────────────┤
                                 FASE 4+5       │ dashboard + PDF, 3-4 semanas
                                                ├──────────┤
                                                FASE 6     │ alertas e inteligência
```

---

## O que precisamos confirmar antes de avançar

| Item | Status | Detalhe |
|---|---|---|
| Tabela de saldo de estoque | ✅ Confirmado | `MT_ESTOQUE` — usar `QTDEREAL`. Sem coluna `QTDERESERVADA`. |
| Como `PR_OP` vincula com pedido | ✅ Confirmado | `PR_OP.PEDIDO` → `VE_PEDIDO.CODIGO` (FK direta) |
| OPSTATUS da PR_OP | ✅ Confirmado | 1=Aguardando, 2=Em Produção, 3=Encerrada, 4=Cancelada |
| `VE_PEDIDOITENS.VLRCUSTO` preenchido | ✅ Confirmado | 64% dos itens; usar `VLRMATERIAL` (99%) como fallback |
| Última compra de materiais | ✅ Confirmado | `MT_MOVIMENTACAO` — `EVENTO=3` = entrada de compra, `DATA` = data movimento |
| Credenciais SMTP para e-mail | ⏳ Pendente | Fase 5.3 e 6 |
| Logo + cores oficiais da Helibombas | ⏳ Pendente | Fase 5.1 — template PDF |

---

## KPIs que teremos no final

- Faturamento do mês (peças) — real-time
- Comparativo YoY e MoM
- Pipeline de orçamentos (R$ e %)
- Taxa de conversão por vendedor
- Carteira pendente por semana de entrega
- Valor parado aguardando faturamento (status L)
- Margem bruta média por vendedor/cliente
- Pedidos com disponibilidade parcial para decisão comercial
