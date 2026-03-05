# Histórico do Projeto — Inteligência Comercial Helibombas

> Documento de backup e referência histórica.
> Registra todas as descobertas, decisões, correções e resultados das sessões de desenvolvimento.
> Última atualização: 04/03/2026

---

## Checkpoint — Encerramento do dia (04/03/2026)

### Produção/PCP (pausado com entregáveis prontos)
- Validado vínculo correto da data de roteiro pela tela do ERP:
    - `PR_OP -> PR_DESENHO -> PR_LISTAENG (POSICAO='00') -> PR_DESENHOROTEIROOPCAO.DTPROJETADO`
- Confirmado no desenho `3809029008`: data de roteiro atualizada para `2026-03-03`.
- Extração completa executada sem limite:
    - Roteiro desatualizado: **465 OPs**
    - MP sem compra há 6+ meses: **126 linhas**
- Priorização operacional gerada:
    - `exports/priorizacao_roteiros_20260304_155927.xlsx`
    - abas: detalhe, consolidado por desenho e top 100 críticos.

### Comercial (próxima retomada)
- Módulo comercial permanece como próxima frente ativa.
- Ponto de retomada sugerido: validação fina dos números no `exports/dashboard.html` e sequência das evoluções operacionais já mapeadas no roadmap.

### Financeiro / RH / Cobrança (entregas adicionais do dia)
- RH salários (diagnóstico de viabilidade):
    - Base acessível contém cadastro/contratação e histórico salarial, mas tabelas de folha mensal estão vazias/quase vazias para este usuário.
    - Relatórios gerados:
        - `exports/rh_salarios_20260304_172137.xlsx`
        - `exports/rh_salarios_20260304_172321.xlsx` (com alertas: queda, fora da curva, sem reajuste)
- Validação pontual salário (caso Caio):
    - Encontrado em `RH_FUNCIONARIO`/`RH_CONTRATACAO`, mas sem eventos recentes de reajuste na base `INDUSTRIAL`.
    - Investigação multi-base indicou restrição de acesso em outros bancos (erro SQL 916), limitando descoberta de possível base alternativa de folha.
- Relatório automático de cobrança por vendedor (Excel):
    - Script principal criado/ajustado: `relatorio_vendedor_cobranca.py`.
    - Regras finais aplicadas:
        - escopo apenas peças (filtros padrão validados do projeto),
        - cotação com fallback para `NUMINTERNO` do pedido,
        - vencimento usando **última parcela** da NF,
        - UF e cidade no detalhado,
        - regra de cidades de MG não atendidas pelo Alexandre.
    - Arquivos gerados (principal):
        - `exports/relatorio_cobranca_vendedores_20260304_181235.xlsx`
    - Caso específico SP+MS (2025):
        - `exports/relatorio_cobranca_SP_MS_2025_20260304_182156.xlsx`

---

## Sessão 1 — Conexão e Relatório Base

### Conquistas
- Estabelecida conexão com SQL Server ERP Sectra (banco `INDUSTRIAL`)
  - Host: `192.168.0.5:1433` | Driver: `SQL Server` (ODBC legado — não usar "ODBC Driver 17")
  - Usuário: `ConsultaHelicoidal` (somente leitura)
- Criado `src/conexao.py` com `get_engine()` — padrão adotado em TODOS os scripts
- Criado `relatorios_vendas.py` — relatório mensal Excel 9 abas

### Correções de filtros (críticas)
Após validação com o ERP, os filtros corretos para "peças" são:

```python
TPVENDA_EXCLUIR = (7, 21, 5, 12, 24, 11, 26, 6, 15, 16, 8, 19, 9, 17, 53, 18, 65, 23)

AND p.STATUS <> 'C'
AND i.STATUS <> 'C'
AND i.TPVENDA NOT IN (...)      -- excluir tipos especiais
AND i.MATERIAL NOT LIKE '8%'   -- excluir bombas completas
AND i.FLAGSUB <> 'S'           -- excluir itens substituídos (duplicados)
AND p.CODIGO NOT IN (           -- excluir clientes estrangeiros
    SELECT p2.CODIGO FROM VE_PEDIDO p2
    JOIN FN_FORNECEDORES f2 ON f2.CODIGO = p2.CLIENTE
    WHERE f2.UF = 'EX'
)
```

### Resultado validado
| Mês | Linhas script | Linhas ERP | Status |
|---|---|---|---|
| Fev/2026 | 663 | 663 | ✅ |
| Jan/2026 | 891 | 891 | ✅ |
| Abr/2025 | 747 | 747 | ✅ |

---

## Sessão 2 — Fase 1: Analytics Estratégico

### Script: `fase1_comparativos.py`
- **Execução**: `python fase1_comparativos.py --ano 2026`
- **Saída**: `exports/fase1_analitico_AAAA_TIMESTAMP.xlsx` — 8 abas

#### Resultado 2026 (jan+fev)
| Métrica | Valor |
|---|---|
| Clientes únicos | 380 |
| Faturamento total | R$4,62M |
| Clientes Curva A | 100 (80% do faturamento) |
| Retenção vs 2025 | 26,3% |

#### Descoberta importante: VLRMATERIAL ≠ Custo
- **`VLRMATERIAL`** = preço de venda (igual a `VLRUNITARIO`) — **não é custo**
- **`VLRCUSTO`** = único campo de custo real (~81% cobertura)
- Essa correção levou à revisão de todos os cálculos de margem

---

## Sessão 3 — Fase 2: Carteira em Tempo Real

### Script: `fase2_carteira.py`
- **Execução**: `python fase2_carteira.py`
- **Saída**: `exports/carteira_DD-MM-AAAA_TIMESTAMP.xlsx` — 5 abas

#### Resultado 26/02/2026
| Métrica | Valor |
|---|---|
| Itens em aberto | 275 |
| Valor total carteira | R$1,53M |
| 🔴 Atrasados | 50 itens |
| 🟡 Urgentes (≤7 dias) | 93 itens |
| 📄 Aguardando NF (status L) | 53 itens · R$167K |

#### Descoberta: OPs só existem para bombas
- `PR_OP` retorna 0 resultados quando filtrado por peças
- OPs (`PR_OP`) só vinculadas a `MATERIAL LIKE '8%'` (bombas completas)
- Peças são compradas/estocadas — sem OP vinculada ao pedido
- Campo `PR_OPSTATUS`: 1=Aguardando, 2=Em produção, 3=Fechada, 4=Cancelada

#### Bug descoberto na Sessão de Dashboard (26/02)
- `DTALTERAFAT` está **NULL em todos os itens com STATUS='L'**
- Solução: usar `ISNULL(DTALTERAFAT, p.DTPEDIDO)` como fallback
- Implementado em `fase4_dashboard.py`

---

## Sessão 4 — Fase 3: Margens e Rentabilidade

### Script: `fase3_margens.py`
- **Execução**: `python fase3_margens.py --mes 2 --ano 2026 [--margem-critica 20]`
- **Saída**: `exports/margens_MM-AAAA_TIMESTAMP.xlsx` — 6 abas

#### Resultado Fev/2026
| Métrica | Valor |
|---|---|
| Itens analisados | 669 |
| Margem média | 52,7% |
| Lucro bruto estimado | R$974K |
| Alertas críticos (<20%) | 8 itens |
| Cobertura VLRCUSTO | 81,3% |

#### Fórmula de margem
```
Margem % = (VLRUNITARIO - VLRCUSTO) / VLRUNITARIO × 100
```
- Calculada somente onde `VLRCUSTO > 0`
- `VLRCUSTO` = custo unitário cadastrado no ERP (pode estar desatualizado)
- Para margem real usar também última compra via `MT_MOVIMENTACAO (EVENTO=3)`

#### Distribuição de margens Fev/2026
| Faixa | Qtd itens |
|---|---|
| Negativa (abaixo do custo) | 152 |
| 40–60% | 436 |
| ≥ 60% | 85 |

---

## Sessão 5 — Fase 4: Dashboard HTML (26/02/2026)

### Instalação
- Plotly 6.5.2 instalado em `c:/python314/` via pip

### Script: `fase4_dashboard.py`
- **Execução**: `python fase4_dashboard.py` (abre navegador)
- **Execução silenciosa**: `python fase4_dashboard.py --no-open` (Task Scheduler)
- **Saída**: `exports/dashboard.html` (sobrescreve sempre)

#### Estrutura do Dashboard (v2 — versão final desta sessão)

| Seção | Conteúdo |
|---|---|
| **KPI Cards** (6) | Vendas mês, Carteira total, 🔴 Atrasados, 🟡 Urgentes, 📄 Aguardando NF, Margem média |
| **Evolução de Vendas** | Barras agrupadas YoY + linhas de meta (R$2,2M/2025 · R$2,5M/2026) |
| **Ranking Externos** | Representantes com `TIPO='E'` — faturamento YoY |
| **Ranking Internos** | Vendedores com `TIPO='I'` — faturamento YoY |
| **Treemap ABC** | Curva A/B/C clientes (top 30, ano atual) |
| **Performance Externos** | Duplo eixo: barras faturamento + linha margem % média |
| **Tabela NF** | Top 10 itens status L mais antigos (dias desde pedido) |
| **Alertas Margem** | Todos os itens <20% do mês atual, com explicação da fórmula |

#### Metas configuráveis
```python
# No topo de fase4_dashboard.py:
META_MENSAL = {
    2025: 2_200_000,   # meta mensal 2025
    2026: 2_500_000,   # meta mensal 2026
}
```

#### Refinamentos aplicados na v2 (26/02/2026)
| Problema | Causa | Solução |
|---|---|---|
| "Faturamento" no KPI | Nome incorreto | Alterado para "Vendas" |
| Gráfico evolução pequeno | height=320, sem meta | height=420 + linhas META_MENSAL |
| Ranking único misturado | Sem separação TIPO | Dois gráficos: externos (E) e internos (I) |
| "Margem por vendedor" obscuro | Gráfico pouco claro | Substituído por Performance Externos (duplo eixo) |
| Tabela NF mostrava 0 dias | `DTALTERAFAT` = NULL | Fallback para `p.DTPEDIDO` |
| Alertas sem explicação | Sem contexto | Adicionada nota com fórmula e regras |

---

## Mapa Técnico do Banco de Dados

### Tabelas principais
| Tabela | Uso |
|---|---|
| `VE_PEDIDOITENS` | Itens de pedido — base de tudo |
| `VE_PEDIDO` | Cabeçalho do pedido |
| `FN_FORNECEDORES` | Clientes e fornecedores (UF='EX' = estrangeiro) |
| `FN_VENDEDORES` | Vendedores (TIPO='I'=interno · 'E'=externo · ATIVO='S') |
| `MT_ESTOQUE` | Estoque atual (TPESTOQUE='AL' para estoque de peças) |
| `MT_MOVIMENTACAO` | Movimentações (EVENTO=3 = entrada de compra) |
| `PR_OP` | Ordens de produção (só para MATERIAL LIKE '8%') |
| `FN_NFS` | Notas fiscais emitidas |
| `VE_ORCAMENTOS` | Orçamentos (STATUS: A=aprovado, P=perdido, E=em aberto) |

### Colunas "armadilha" — erros comuns evitados
| Coluna | Armadilha | Correto |
|---|---|---|
| `VLRMATERIAL` | Parece custo — é venda | `VLRCUSTO` é o custo |
| `FN_VENDEDORES.ATIVO` | Parece boolean — é varchar | `WHERE ATIVO = 'S'` |
| `MT_MOVIMENTACAO.DATA` | Parece `DTMOVIMENTO` | O nome da coluna é `DATA` |
| `VE_PEDIDOITENS.DTALTERAFAT` | Parece sempre preenchido | NULL em muitos itens L |
| `FN_FORNECEDORES.RAZAO` | Não é `RAZAOSOCIAL` | Coluna se chama `RAZAO` |
| `FN_VENDEDORES.RAZAO` | Não é `NOME` | Coluna se chama `RAZAO` |

### Filtros obrigatórios para relatório de peças
```sql
AND p.STATUS <> 'C'
AND i.STATUS <> 'C'                          -- ou NOT IN ('C','F') para carteira
AND i.TPVENDA NOT IN (7,21,5,12,24,11,26,6,15,16,8,19,9,17,53,18,65,23)
AND i.MATERIAL NOT LIKE '8%'                 -- exclui bombas completas
AND i.FLAGSUB <> 'S'                         -- exclui itens substituídos
AND p.CODIGO NOT IN (                        -- exclui clientes estrangeiros
    SELECT p2.CODIGO FROM VE_PEDIDO p2
    JOIN FN_FORNECEDORES f2 ON f2.CODIGO = p2.CLIENTE
    WHERE f2.UF = 'EX'
)
```

### PR_OPSTATUS — códigos
| Valor | Significado |
|---|---|
| 1 | Aguardando |
| 2 | Em produção |
| 3 | Fechada/Concluída |
| 4 | Cancelada |

### MT_MOVIMENTACAO — EVENTO
| Valor | Significado |
|---|---|
| 3 | Entrada de compra (para calcular última compra) |
| Outros | Saídas, ajustes, etc |

---

## Ambiente de Desenvolvimento

| Item | Valor |
|---|---|
| Python | 3.14.2 em `c:/python314/python.exe` |
| SQL Server Driver | `SQL Server` (ODBC legado) |
| Pandas | Disponível |
| OpenPyXL | Disponível |
| SQLAlchemy | Disponível |
| Python-dotenv | Disponível |
| **Plotly** | **6.5.2** (instalado 26/02/2026) |
| Terminal encoding | `$env:PYTHONIOENCODING="utf-8"` antes de rodar |
| Diretório de trabalho | `C:\Dev\Acesso_servidor\` |

### Comandos de execução (PowerShell)
```powershell
# Sempre usar antes de rodar qualquer script:
$env:PYTHONIOENCODING="utf-8"

# Relatório mensal
c:/python314/python.exe relatorios_vendas.py

# Fase 1 — Analytics
c:/python314/python.exe fase1_comparativos.py --ano 2026

# Fase 2 — Carteira
c:/python314/python.exe fase2_carteira.py

# Fase 3 — Margens
c:/python314/python.exe fase3_margens.py --mes 2 --ano 2026

# Fase 4 — Dashboard (abre navegador)
c:/python314/python.exe fase4_dashboard.py

# Fase 4 — Dashboard (Task Scheduler)
c:/python314/python.exe fase4_dashboard.py --no-open
```

---

## Estado do Projeto em 26/02/2026

### Fases
| Fase | Script | Status | Resultado chave |
|---|---|---|---|
| 1 — Analytics estratégico | `fase1_comparativos.py` | ✅ | R$4,62M jan+fev, 100 clientes A |
| 2 — Carteira em tempo real | `fase2_carteira.py` | ✅ | 275 itens, R$1,53M, 50 atrasados |
| 3 — Margens e rentabilidade | `fase3_margens.py` | ✅ | Margem 52,7%, R$974K bruto |
| 4 — Dashboard HTML | `fase4_dashboard.py` | ✅ | 6 KPIs, 6 gráficos, 2 tabelas |
| 5 — PDF mensal automatizado | `relatorio_pdf.py` | ⏳ | Aguarda logo + SMTP |
| 6 — Alertas diários | `alertas_diarios.py` | ⏳ | Aguarda SMTP |

### Próximas evoluções identificadas
1. **Task Scheduler**: agendar `fase4_dashboard.py --no-open` às 07h00 todos os dias
2. **Atalho na área de trabalho**: apontar para `exports/dashboard.html`
3. **Melhorar cálculo de dias NF**: investigar se há outra tabela com o log de mudança de status
4. **PDF mensal**: requer logo em PNG e configuração de servidor SMTP
5. **Alertas diários**: e-mail automático com pedidos vencendo + itens L > 10 dias

### Perguntas para o time respondidas neste projeto
- ✅ Qual é a meta mensal? → R$2,5M/mês em 2026 (R$2,2M em 2025)
- ✅ VLRMATERIAL é custo? → Não. É preço de venda = VLRUNITARIO
- ✅ Bombas têm OP? → Sim, MATERIAL LIKE '8%'. Peças existem sem OP
- ✅ Como separar interno de externo? → FN_VENDEDORES.TIPO = 'I' ou 'E'
- ✅ Por que itens aparecem duplicados? → FLAGSUB = 'S' — filtrar sempre
- ✅ O que é DTALTERAFAT? → Data de alteração para faturamento (muito vezes NULL)

---

## Sessão 6 — Fechamento frente de Notas (03/03/2026)

### Objetivo da sessão
- Consolidar um **Relatório 506 Melhorado** confiável para apresentação ao controller.
- Corrigir inconsistências de data (emissão da NF vs emissão do título) e quitação por parcelas.
- Deixar trilha de auditoria para não depender do histórico de chat.

### Entregas concluídas
- Script evoluído: `relatorio_506_excel.py`
- Saída principal: planilha com aba `506_Melhorado`
- Colunas chave adicionadas/ajustadas:
    - `Dt_Emissao` (origem correta: `FN_NFS.DTEMISSAO`)
    - `Dt_Emissao_Titulo` (origem: `FN_RECEBER.DTEMISSAO`)
    - `Dt_Ultimo_Venc_NF`, `Status_NF_Quitada`, `Parcelas_Total`, `Parcelas_Pagas`, `Parcelas_Abertas`
    - `Tipo_Produto` = `PECA` / `BOMBA` / `AMBAS`
    - `Qtd_Itens_Peca`, `Qtd_Itens_Bomba` (auditoria de classificação)

### Correções críticas aplicadas
- **Data de emissão** corrigida para emissão da NF (não do título financeiro).
- **Sem duplicação**: validação obrigatória de unicidade por `Receber`.
- **Classificação por tipo** por composição da NF:
    - só item `MATERIAL LIKE '8%'` → `BOMBA`
    - só itens diferentes de `8%` → `PECA`
    - mistura dos dois na mesma NF → `AMBAS`

### Validação técnica implementada (nível controller)
- Adicionada aba `Validacao` no Excel, comparando:
    - agregados do relatório (Excel)
    - consulta SQL independente de controle
- Campos comparados por tipo e total: títulos, NFs, valor devido, valor recebido.
- Status final por linha: `OK` ou `DIVERGENTE`.

### Arquivo de referência final validado
- `exports/relatorio_506_excel_20260303_130626.xlsx`
    - Aba `506_Melhorado`
    - Aba `Validacao` (todos os tipos e total com `OK`)

### Caso de teste validado em reunião (NF 35851)
- Emissão NF: `28/06/2024`
- Último vencimento NF: `23/08/2024`
- Parcelamento: 8 total / 7 pagas / 1 em aberto
- Status: `PENDENTE`

### Encaminhamento decidido
- Frente de **Notas/506** considerada fechada para esta etapa.
- Próximo foco: **evolução comercial e regras de comissionamento** (microregras por estado/cidade/vendedor), usando o 506 melhorado como base de conferência.

---

## Sessão 7 — Evolução Comercial em Tempo Real (03/03/2026)

### Objetivo da sessão
- Sair do foco financeiro e colocar o dashboard para uso de coordenação comercial diária.
- Corrigir ranking de externos para refletir território real por regra de negócio.
- Tornar a interface mais rápida para decisão durante a operação e reuniões.

### Entregas concluídas no `fase4_dashboard.py`
- Bloco **Prioridades do Dia** (itens críticos para ação imediata).
- Bloco **Cotações Geradas — Controle Comercial** com KPI e tabela operacional.
- Funil de cotações e conversão por território.
- Mapeamento territorial explícito de representantes externos por UF/cidade/data.
- Exclusão de vendedores internos dos rankings de externos.
- Fallback territorial para conta Helibombas quando necessário.
- Leitura automática de exceções MG por arquivo externo (`.xlsx`/`.xlsm`).
- Reorganização visual para **Modo Operação** no topo (ação antes de estratégia).
- Navegação por âncoras no topo: **Operação**, **Cotações**, **Estratégico**.

### Correções técnicas aplicadas
- Join de cotações ajustado para coluna correta de cliente em `VE_ORCAMENTOS`.
- Numeração de cotação tratada como texto para evitar erro de conversão.
- Geração final validada para 4 artefatos HTML:
    - `exports/dashboard.html`
    - `exports/dashboard_base.html`
    - `exports/dashboard_sem_contrato.html`
    - `exports/dashboard_reuniao.html`

### Resultado de negócio
- Dashboard passa a responder ao ritual de coordenação do dia:
    1. ver fila crítica,
    2. atacar cotações e conversão,
    3. consultar contexto estratégico sem perder foco operacional.

### Próximo passo recomendado
- Implementar placar de SLA diário (entrada x resolvidas no dia) para medir execução da equipe.
