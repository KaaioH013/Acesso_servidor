# Checkpoint de Continuidade — 03/03/2026

## Estado atual (congelado)

Este arquivo registra o estado consolidado do projeto após as evoluções de 03/03/2026, para retomada sem depender de histórico de conversa.

## O que está pronto

- Frente financeira de notas (`relatorio_506_excel.py`) encerrada e validada com controller.
- Dashboard comercial (`fase4_dashboard.py`) evoluído para operação em tempo real.
- Ranking de externos territorial corrigido por regras explícitas.
- Internos removidos da disputa de ranking externo.
- Exceções MG aplicadas por arquivo externo.
- Modo Operação no topo do dashboard.
- Navegação rápida por âncoras: Operação, Cotações, Estratégico.

## Artefatos de saída válidos

- `exports/dashboard.html`
- `exports/dashboard_base.html`
- `exports/dashboard_sem_contrato.html`
- `exports/dashboard_reuniao.html`
- `exports/relatorio_506_excel_20260303_130626.xlsx`

## Regras críticas vigentes

1. Em 506, emissão de NF usa `FN_NFS.DTEMISSAO`; emissão de título fica separada.
2. Quitação de NF é calculada por parcelas (total, pagas, abertas, vencimento final).
3. Classificação de tipo de produto: `PECA`, `BOMBA`, `AMBAS`.
4. Territorialização comercial de externos por UF/cidade/data, com exceções MG.
5. Itens internos não entram no ranking territorial de externos.

## Arquivos de referência para continuidade

- `ROADMAP.md` (estratégia atualizada para ERP 360)
- `HISTORICO_PROJETO.md` (linha do tempo de decisões e correções)
- `HANDOFF_2026-03-03_NOTAS_E_COMERCIAL.md` (encerramento da frente 506)
- `BACKLOG_PAINEL_COMERCIAL_UIUX.md` (fila de evolução do painel)
- `GUIA_OPERACAO_COMERCIAL_TEMPO_REAL.md` (uso diário da coordenação)

## Próxima evolução recomendada (imediata)

- Implementar placar de SLA diário no dashboard:
  - entrada do dia,
  - resolvidas no dia,
  - backlog inicial/final,
  - tempo médio de resolução.

## Incremento aplicado (03/03/2026 — etapa 1)

- SLA diário v1 implementado em `fase4_dashboard.py` e publicado nos 4 HTMLs.
- Métricas já ativas no bloco `⏱️ SLA Diário — Entrada x Resolução`:
  - Entradas (itens em `L` com liberação no dia)
  - Resolvidas (itens com NF emitida no dia)
  - Saldo do dia
  - Backlog atual
  - Taxa de saída
- Validação executada com geração completa:
  - `python fase4_dashboard.py --no-open`
  - Saída OK para `dashboard.html`, `dashboard_base.html`, `dashboard_sem_contrato.html`, `dashboard_reuniao.html`

## Incremento aplicado (03/03/2026 — etapa 2)

- Correção da conversão de cotações no `fase4_dashboard.py`.
- Status de orçamentos normalizados para leitura operacional:
  - `A` = aprovada
  - `P/X/C` = perdida
  - demais = aberta
- Conversão agora exibida em duas lentes:
  - **Geral**: aprovadas ÷ geradas no mês
  - **Finalizadas**: aprovadas ÷ (aprovadas + perdidas)
- Conversão por valor baseada em `Vlr_Orcado` (evita distorção por `VLREFETIVO`).
- Resultado validado no HTML:
  - `Conv. geral qtd/valor: 2,9% · 2,9%`
  - `Conv. finalizadas qtd/valor: 47,4% · 54,8%`

## Incremento aplicado (03/03/2026 — etapa 3)

- Dashboard com duas abas implementado:
  - `📅 Mês Fechado` (visão histórica consolidada)
  - `⚡ MTD Hoje` (visão do mês em andamento)
- Nova coleta MTD adicionada (`get_faturamento_mtd_atual`):
  - faturamento/pedidos de hoje,
  - faturamento/pedidos/itens MTD,
  - run-rate diário,
  - projeção de fechamento,
  - gap para meta e atingimento projetado.
- Geração validada com sucesso para os 4 HTMLs.

## Incremento aplicado (03/03/2026 — etapa 4)

- Automação Windows Task Scheduler implantada e testada.
- Scripts criados:
  - `scripts/run_dashboard_refresh.ps1`
  - `scripts/register_dashboard_tasks.ps1`
  - `scripts/unregister_dashboard_tasks.ps1`
- Tarefas registradas:
  - `ERP_Dashboard_Daily_0700`
  - `ERP_Dashboard_Hourly_0800_1800`
- Runner com log + retentativa (3 tentativas) e ambiente UTF-8 para estabilidade.
- Log de execução: `logs/dashboard_refresh.log`.
- Evidência de sucesso no agendador: último resultado `0` na tarefa diária após ajuste.

## Incremento aplicado (03/03/2026 — etapa 5)

- Script de validação de consistência criado: `validar_dashboard.py`.
- Escopo de validação:
  - carteira e semáforo,
  - SLA x carteira,
  - métricas MTD,
  - consistência de status e conversão de cotações,
  - conferência de aprovadas (SQL independente vs dashboard).
- Execução validada:
  - comando: `py -3 validar_dashboard.py --somente ambos`
  - resultado: `OK: 20 | ALERTA: 0 | FALHA: 0`
  - evidência: `exports/validacao_dashboard_20260303_172248.xlsx`

## Incremento aplicado (03/03/2026 — etapa 6)

- Cadência de automação ajustada para menor carga:
  - refresh diário: `07:00`
  - refresh intradiário: a cada `2 horas` entre `08:00` e `18:00`
  - validação diária: `07:10`
- Tarefas ativas:
  - `ERP_Dashboard_Daily_0700`
  - `ERP_Dashboard_Intraday_2h_0800_1800`
  - `ERP_Dashboard_Validation_0710`
- Validação automática testada via scheduler com sucesso:
  - `OK: 20 | ALERTA: 0 | FALHA: 0`
  - evidência: `exports/validacao_dashboard_20260303_172959.xlsx`

## Incremento aplicado (03/03/2026 — etapa 7)

- Roadmap expandido com novos blocos de evolução comercial (quick wins, eficiência, risco de meta e governança).
- Primeira evolução prática iniciada:
  - export diário da fila Top 20 acionável para ritual comercial.
- Implementação técnica em `fase4_dashboard.py`:
  - função reutilizável `construir_prioridades_operacionais`;
  - geração automática de:
    - `exports/fila_acao_diaria.xlsx`
    - `exports/fila_acao_diaria_YYYYMMDD_HHMMSS.xlsx`
- Validação da geração concluída com sucesso junto dos 4 dashboards HTML.

## Incremento aplicado (03/03/2026 — etapa 8)

- Correções críticas de confiabilidade aplicadas no `fase4_dashboard.py`:
  - cotações/funil recalibrados com base em `STATUS='E'` como encerrada,
  - conversão territorial calculada por `UF` (estado do cliente),
  - `Prioridades do Dia` e `Top NF` usando somente `DTALTERAFAT` para medir tempo em `L`,
  - `SLA Diário` redefinido para fila de faturamento (itens `L` liberados no dia),
  - ranking interno removido da interface,
  - títulos ajustados para deixar claro que a base é vendas de peças.
- Verificação de veracidade YoY por vendedor executada: total da tabela YoY bate com card YoY.
- Nova validação pós-correção:
  - `exports/validacao_dashboard_20260303_175520.xlsx`
  - `OK: 20 | ALERTA: 0 | FALHA: 0`

## Incremento aplicado (03/03/2026 — etapa 9)

- Alinhamento da consulta de cotações com a tela oficial do ERP (`VE_ORCAMENTOS`):
  - filtro `FILIAL = 1` aplicado no dashboard.
- Números de fevereiro/2026 passaram a refletir a referência operacional:
  - criadas: `299`
  - encerradas (`STATUS='E'`): `181`
- Conversão e funil recalculados com essa base.
- Validação final da versão:
  - `exports/validacao_dashboard_20260303_180012.xlsx`
  - `OK: 20 | ALERTA: 0 | FALHA: 0`

## Comando de atualização do dashboard

```powershell
py -3 .\fase4_dashboard.py --no-open
```
