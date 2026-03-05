# Documentação — Acesso ERP Sectra (SQL Server)

> Coordenador de Vendas — área de **peças** (não cobre bombas completas nem exportação)

---

## 1. Ambiente Python

| Item | Valor |
|---|---|
| Python | 3.14.2 em `c:/python314/python.exe` (instalação global, não venv) |
| Pacotes instalados | pyodbc, pandas, sqlalchemy, openpyxl, python-dotenv, rich, tabulate |
| Encoding no terminal | `$env:PYTHONIOENCODING="utf-8"` antes de rodar scripts |

**Como rodar qualquer script:**
```powershell
$env:PYTHONIOENCODING="utf-8"; c:/python314/python.exe nome_do_script.py
```

---

## 2. Conexão com o Banco

**Driver instalado na máquina:** `SQL Server` (legado — NÃO é "ODBC Driver 17")

**Arquivo `.env`:**
```
DB_SERVER=192.168.0.5
DB_PORT=1433
DB_DATABASE=INDUSTRIAL
DB_USERNAME=ConsultaHelicoidal
DB_PASSWORD=Consulta@;
DB_DRIVER=SQL Server
```

> ⚠️ A senha contém `;` — por isso é escapada com `{password}` na connection string (ex: `PWD={Consulta@;}`)

**Trecho de conexão (usado em todos os scripts):**
```python
from dotenv import load_dotenv
import os, pandas as pd
from sqlalchemy import create_engine, text
from urllib.parse import quote_plus

load_dotenv()
pwd = os.getenv("DB_PASSWORD")
pwd_escaped = "{" + pwd.replace("}", "}}") + "}"
conn_str = (
    f"DRIVER={{{os.getenv('DB_DRIVER')}}};SERVER={os.getenv('DB_SERVER')},{os.getenv('DB_PORT')};"
    f"DATABASE={os.getenv('DB_DATABASE')};UID={os.getenv('DB_USERNAME')};PWD={pwd_escaped};"
    "TrustServerCertificate=Yes;"
)
engine = create_engine(f"mssql+pyodbc:///?odbc_connect={quote_plus(conn_str)}")
```

---

## 3. Banco de Dados: INDUSTRIAL

- **1.579 tabelas**, 161 views, 3.059 FKs
- ERP: **Sectra**
- Prefixos dos módulos:

| Prefixo | Módulo |
|---|---|
| `VE_` | Vendas |
| `MT_` | Materiais / Estoque |
| `FN_` | Financeiro / Fiscal |
| `PR_` | Produção |
| `CO_` | Compras |
| `RH_` | Recursos Humanos |

---

## 4. Tabelas-Chave do Módulo de Vendas

### `VE_PEDIDO` — Cabeçalho do Pedido
| Coluna | Descrição |
|---|---|
| `CODIGO` | PK |
| `DTPEDIDO` | Data do pedido |
| `CLIENTE` | FK → `FN_FORNECEDORES.CODIGO` |
| `VENDEDOR` | FK → `FN_VENDEDORES.CODIGO` |
| `STATUS` | `V`=em aberto, `E`=entregue/faturado, `C`=cancelado, `F`=faturado |
| `VLRTOTAL` | Valor total do pedido |
| `PEDORIGEM` | FK → `VE_ORCAMENTOS.CODIGO` (orçamento de origem) |
| `NUMINTERNO` | Número interno do pedido |
| `PEDIDOCLI` | Número do pedido do cliente |

### `VE_PEDIDOITENS` — Itens do Pedido
| Coluna | Descrição |
|---|---|
| `CODIGO` | PK |
| `PEDIDO` | FK → `VE_PEDIDO.CODIGO` |
| `SEQ` | Sequência do item no pedido |
| `MATERIAL` | Código do material |
| `DESCRICAO` | Descrição do item |
| `TPVENDA` | FK → `VE_TPVENDA.CODIGO` (tipo de venda — ver seção 5) |
| `QTDE` | Quantidade pedida |
| `QTDEFAT` | Quantidade faturada |
| `VLRUNITARIO` | Valor unitário |
| `VLRTOTAL` | Valor total do item |
| `STATUS` | `V`=em aberto, `F`=faturado, `L`=liberado, `C`=cancelado |
| `FLAGSUB` | `S`=item substituído (foi dividido em dois) → **excluir sempre** |
| `PAI` | FK para item original quando `FLAGSUB='S'` |
| `NFEITEM` | FK → `FN_NFEITENS.CODIGO` |
| `DTPRAZO` | Data de prazo do item |
| `DTALTERAFAT` | Data de alteração/liberação para faturamento — **pode ser NULL** mesmo em itens `STATUS='L'` |
| `VLRCUSTO` | Custo unitário cadastrado (~81% cobertura) |
| `PERCDESCONTO` | Percentual de desconto concedido |

> ⚠️ **`FLAGSUB = 'S'`**: quando um item é dividido em partes (ex: 30 peças viram 28+2), o item original fica com `FLAGSUB='S'` e os dois filhos apontam para ele via `PAI`. O ERP exibe só os filhos — sempre filtrar `AND i.FLAGSUB <> 'S'`
> 📌 **`DTALTERAFAT` NULL**: itens em `STATUS='L'` frequentemente têm essa coluna vazia. Para calcular "há quantos dias aguarda NF", usar `ISNULL(DTALTERAFAT, p.DTPEDIDO)` ou `DtPedido` como fallback.

### `VE_ORCAMENTOS` — Orçamentos
| Coluna | Descrição |
|---|---|
| `CODIGO` | PK |
| `NUMERO` | Número do orçamento (exibição) |
| `DTCADASTRO` | Data de cadastro (NÃO é `DTORCAMENTO`) |
| `DTVALIDADE` | Data de validade |
| `STATUS` | `A`=aprovado, `P`=perdido, `E`=em aberto |
| `VLRORCADO` | Valor orçado |
| `VLREFETIVO` | Valor efetivo (convertido) |
| `VENDEDOR` | FK → `FN_VENDEDORES.CODIGO` |
| `PEDORIGEM` | FK → `VE_PEDIDO.CODIGO` |
| `DESCRICAO` | Descrição do orçamento |

### `FN_FORNECEDORES` — Clientes e Fornecedores
| Coluna | Descrição |
|---|---|
| `CODIGO` | PK |
| `RAZAO` | Nome / Razão social (NÃO é `RAZAOSOCIAL`) |
| `FANTASIA` | Nome fantasia |
| `CGC` | CNPJ/CPF |
| `CIDADE` | Cidade |
| `UF` | Estado — **`'EX'` = cliente estrangeiro (exportação)** |
| `PAIS` | Código do país |

> ⚠️ Essa tabela serve tanto para **clientes** quanto para **transportadoras**

### `FN_VENDEDORES` — Vendedores
| Coluna | Descrição |
|---|---|
| `CODIGO` | PK |
| `RAZAO` | Nome do vendedor (NÃO é `NOME`) |
| `TIPO` | **`'I'` = vendedor interno · `'E'` = representante externo** |
| `ATIVO` | `'S'` = ativo (varchar, não bit) |
| `CKREPRESENTANTE` | Flag adicional de representante (nem sempre preenchido) |
| `GERENTE` | Gerente responsável |

> ⚠️ `ATIVO` é varchar → usar `WHERE ATIVO = 'S'` (não `ATIVO = 1`)
> ⚠️ `TIPO` deve entrar no GROUP BY quando usado junto com agregações

### `FN_NFS` — Notas Fiscais
| Coluna | Descrição |
|---|---|
| `CODIGO` | PK |
| `NRONOTA` | Número da NF |
| `DTEMISSAO` | Data de emissão |
| `DTSAIDA` | Data de saída |
| `STATUSNF` | Status (`C`=cancelada) |
| `VEPEDIDO` | FK → `VE_PEDIDO.CODIGO` |
| `VLRPRODUTO` | Valor dos produtos |
| `VLRIPI` | Valor do IPI |
| `VLRICMSNF` | Valor do ICMS |
| `VLRTOTAL` | Valor total |

### `FN_NFEITENS` — Itens das Notas Fiscais
| Coluna | Descrição |
|---|---|
| `CODIGO` | PK |
| `NFE` | FK → `FN_NFS.CODIGO` |
| `PEDIDOITEM` | FK → `VE_PEDIDOITENS.CODIGO` |
| `MATERIAL` | Código do material |

### `VE_TPVENDA` — Tipos de Venda
| Coluna | Descrição |
|---|---|
| `CODIGO` | PK |
| `DESCRICAO` | Descrição do tipo |

---

## 5. Filtros do Relatório de Peças

### Regra de negócio geral
- Área cobre apenas **peças e componentes** — não bombas completas, não exportação
- Materiais cujo código começa com `'8'` = bombas completas → **excluir**
- Clientes com `UF = 'EX'` = estrangeiros (exportação) → **excluir pedido inteiro**

### Tipos de venda EXCLUÍDOS (não são vendas reais da área)

| Código | Descrição |
|---|---|
| 5 | REMESSA P/ DEMONSTRACAO |
| 6 | REMESSA P/ BRINDE |
| 7 | REMESSA EM GARANTIA |
| 8 | REMESSA P/ TROCA |
| 9 | RETORNO CONSERTO |
| 11 | REMESSA COM FIM ESPECIFICO DE EXPORTACAO (PROD) |
| 12 | REMESSA P/ DEMONSTRACAO EM FEIRA |
| 15 | REMESSA P/ BONIF E BRINDE P/ CF |
| 16 | REMESSA PARA AMOSTRA |
| 17 | RETORNO DE DEMONSTRACAO |
| 18 | PRESTACAO DE SERVICOS |
| 19 | TRANSFERENCIA |
| 21 | REMESSA EM GARANTIA (variante) |
| 23 | **VENDA EXPORTAÇÃO** |
| 24 | SIMPLES REMESSA |
| 26 | REMESSA COM FIM ESPECIFICO DE EXPORTACAO (REVENDA) |
| 53 | DEVOLUCAO DE MERCADORIA RECEBIDA EM TRANSFERENCIA |
| 65 | REMESSA EM GARANTIA SEM ST |

### Tipos de venda INCLUÍDOS (vendas reais de peças)
- VENDA DE PRODUTOS E SERVIÇOS P/ CF
- VENDA DE PRODUTOS E SERV P/ CF INTERESTADUAL
- VENDA DE PRODUTOS
- VENDA DE PRODUTOS SEM ST INDUSTRIALIZACAO
- REVENDA DE PRODUTOS P/ CF
- REVENDA DE PRODUTOS
- REVENDA DE PRODUTOS SEM ST
- VENDA POR CONTA E ORDEM - PEÇAS - CONS.FINAL
- VENDA DE PROD E SERVIÇOS P/ CF-ISENTO PIS/COFINS
- VENDA COM SUSPENSÃO IPI (USO E CONSUMO)
- REVENDA - SAÍDA COM SUSPENSÃO IPI (USO E CONSUMO)
- (e outros similares que apareçam)

### WHERE completo padrão (copiar/colar em qualquer query)
```sql
WHERE p.DTPEDIDO BETWEEN :ini AND :fim
  AND p.STATUS <> 'C'                          -- sem pedidos cancelados
  AND i.STATUS <> 'C'                          -- sem itens cancelados
  AND i.TPVENDA NOT IN (7,21,5,12,24,11,26,6,15,16,8,19,9,17,53,18,65,23)
  AND i.MATERIAL NOT LIKE '8%'                 -- sem bombas completas
  AND i.FLAGSUB <> 'S'                         -- sem itens originais divididos
  AND p.CODIGO NOT IN (                        -- sem clientes estrangeiros
      SELECT p2.CODIGO FROM VE_PEDIDO p2
      JOIN FN_FORNECEDORES f ON f.CODIGO = p2.CLIENTE
      WHERE f.UF = 'EX'
  )
```

---

## 6. Scripts Disponíveis

| Script | Finalidade |
|---|---|
| `listar_bancos.py` | Lista todos os bancos do servidor (usado para descobrir o nome do DB) |
| `inspecionar_industrial.py` | Mostra todas as tabelas do banco INDUSTRIAL com contagem de linhas |
| `inspecionar_vendas.py` | Mapeia todas as tabelas VE_ com colunas e FK |
| `relatorios_vendas.py` | **Principal mensal** — gera Excel com 9 abas para o mês atual |
| `fase1_comparativos.py` | **Análises estratégicas** — Evolução YoY, Curva ABC, Churn, Conversão, Ranking Vendedores |
| `fase2_carteira.py` | **Carteira em tempo real** — semáforo de prazo, aguardando NF, estoque disponível, OPs |
| `fase3_margens.py` | **Margens e rentabilidade** — VLRCUSTO vs venda, última compra, alertas críticos, desconto vs margem |
| `fase4_dashboard.py` | **Dashboard HTML interativo** — KPI cards, gráficos Plotly, treemap ABC, tabelas de alertas; gera `exports/dashboard.html`. Configurar **metas mensais** em `META_MENSAL` no topo do script. Ranking separado externo (`TIPO='E'`) e interno (`TIPO='I'`). |
| `consulta_janeiro.py` | Consulta rápida de um mês específico para conferência |
| `src/conexao.py` | Módulo de conexão compartilhado |
| `src/explorador.py` | Funções para explorar o schema (listar tabelas, colunas, FKs) |
| `src/exportar.py` | Utilitários de exportação CSV/Excel |
| `analise_erp.ipynb` | Notebook Jupyter para análises interativas |

### Gerar relatório do mês atual
```powershell
$env:PYTHONIOENCODING="utf-8"; c:/python314/python.exe relatorios_vendas.py
```
Arquivo salvo em `exports/relatorio_pecas_MM-AAAA_TIMESTAMP.xlsx`

---

## 7. Abas do Relatório Excel

| Aba | Conteúdo |
|---|---|
| **Itens Pedidos** | Todos os itens de pedidos do período, com cliente, vendedor, material, qtde, valor |
| **Pedidos Resumo** | Cabeçalho dos pedidos (1 linha por pedido) com link para orçamento de origem |
| **Fat. por Vendedor** | Faturamento agrupado por vendedor e mês |
| **Ranking Clientes** | Clientes ordenados por valor total de peças |
| **Prod. + Vendidos** | Top 50 materiais mais vendidos por faturamento |
| **Carteira Abertos** | Itens com saldo a entregar (qtde > qtde faturada), ordenados por atraso |
| **Origem Orcamento** | Pedidos do período vinculados ao orçamento de origem via `PEDORIGEM` |
| **NFs Emitidas** | Notas fiscais emitidas vinculadas a pedidos do período |
| **Comp. Mensal** | Comparativo mês a mês dos últimos 2 anos |

---

## 8. Validações Realizadas

| Mês | Linhas (itens) | Pedidos | Valor Total | Status |
|---|---|---|---|---|
| Fevereiro/2026 | 663 | 252 | R$ 1.992.992,21 | ✅ Bateu com sistema |
| Janeiro/2026 | 891 | 373 | R$ 2.627.055,36 | ✅ Bateu com sistema |
| Abril/2025 | 747 | 328 | R$ 2.028.871,69 | ✅ Bateu com sistema |

---

## 9. Relacionamentos Principais

```
VE_ORCAMENTOS ──────────────────────────────────────────────────┐
     CODIGO                                                       │
        ↑                                                         │
VE_PEDIDO.PEDORIGEM                                               │
                                                                  │
VE_PEDIDO ──────────────────────────────────────────────────────┐│
    CODIGO                                                       ││
    CLIENTE ──→ FN_FORNECEDORES.CODIGO (RAZAO, UF, CGC, CIDADE) ││
    VENDEDOR ──→ FN_VENDEDORES.CODIGO  (RAZAO)                  ││
        ↑                                ↑                       ││
VE_PEDIDOITENS.PEDIDO          FN_NFS.VEPEDIDO                   ││
                                    ↑                            ││
VE_PEDIDOITENS ────────────────────────────────────────────────  ││
    CODIGO                                                       ││
    TPVENDA ──→ VE_TPVENDA.CODIGO  (DESCRICAO)                  ││
    NFEITEM ──→ FN_NFEITENS.CODIGO                               ││
        ↑                                                        ││
FN_NFEITENS.PEDIDOITEM          FN_NFS ──────────────────────────┘│
                                    NFE ──→ FN_NFEITENS          │
                                    VEPEDIDO ──→ VE_PEDIDO ───────┘
```

---

## 10. Módulo de Produção (PR_)

### `PR_OP` — Ordens de Produção
| Coluna | Tipo | Descrição |
|---|---|---|
| `CODIGO` | numeric | PK |
| `NROOP` | int | Número da OP (sequencial exibido ao usuário) |
| `OPSTATUS` | numeric | Status **→ ver tabela abaixo** |
| `DTCADASTRO` | datetime | Data de abertura da OP (NÃO é `DTABERTURA`) |
| `DTPRAZO` | datetime | Data de prazo de entrega |
| `DTEFETIVA` | datetime | Data de encerramento efetiva (`NULL` se ainda aberta) |
| `MATERIAL` | varchar | Código do material a produzir |
| `QTDE` | numeric | Quantidade pedida |
| `QTDEPRODUZIDA` | numeric | Quantidade já produzida |
| `PEDIDO` | numeric | FK → `VE_PEDIDO.CODIGO` (e `VE_PEDIDOITENS.CODIGO`) |
| `TPOP` | numeric | Tipo de OP |

**Tabela `PR_OPSTATUS` — decode de status:**
| CODIGO | DESCRICAO |
|---|---|
| 1 | AGUARDANDO LIBERAÇÃO |
| 2 | EM PRODUÇÃO |
| 3 | ENCERRADA |
| 4 | CANCELADA |

**Para carteira em produção:** `WHERE o.OPSTATUS IN (1, 2)` — ordens abertas/em andamento.

**JOIN com pedido:**
```sql
JOIN PR_OP o ON o.PEDIDO = p.CODIGO
```

---

## 11. Módulo de Materiais / Estoque (MT_)

### `MT_ESTOQUE` — Saldo de Estoque
| Coluna | Tipo | Descrição |
|---|---|---|
| `FILIAL` | numeric | Filial |
| `MATERIAL` | varchar | Código do material |
| `TPESTOQUE` | varchar | Tipo de estoque (`AL`=Almoxarifado, outros) |
| `MATERIALLOTE` | numeric | Lote |
| `QTDEREAL` | numeric | Quantidade real em estoque ✅ usar este |
| `QTDEFISCAL` | int | Quantidade fiscal |

> ⚠️ **Não existe `QTDERESERVADA`** nesta tabela. Reservas precisam ser calculadas via `MT_MOVIMENTACAO` ou `PR_OPMATERIAIS`.

### `MT_MOVIMENTACAO` — Movimentações de Estoque (1.28M linhas)
Colunas relevantes:

| Coluna | Tipo | Descrição |
|---|---|---|
| `CODIGO` | numeric | PK |
| `FILIAL` | numeric | Filial |
| `MATERIAL` | varchar | Código do material |
| `TPESTOQUE` | varchar | Tipo de estoque |
| `DATA` | datetime | Data da movimentação (NÃO é `DTMOVIMENTO`) |
| `EVENTO` | numeric | Código do tipo de movimento — **ver tabela abaixo** |
| `QTDE` | numeric | Quantidade movimentada |
| `VLRTOTAL` | numeric | Valor total da movimentação |
| `DESCRICAO` | varchar | Descrição textual |
| `NFEITENS` | numeric | FK → `FN_NFEITENS.CODIGO` (presente em compras) |
| `PEDIDOITEM` | numeric | FK → `VE_PEDIDOITENS.CODIGO` |
| `OPMOVIMENTACAO` | numeric | FK para movimentação de OP |
| `STATUS` | varchar | Status da movimentação |

**Principais tipos de `EVENTO`:**
| EVENTO | Qtd linhas | Descrição / Uso |
|---|---|---|
| 10 | 594.725 | Requisição de estoque (consumo interno/OP) |
| 1 | 235.540 | Saída (ex: vulcanização) |
| 3 | 138.155 | **ENTRADA DE COMPRA** ← usar para última compra |
| 9 | 125.541 | Consumo em OP de produção |
| 2 | 96.608 | Venda/saída exportação |
| 16 | 60.339 | Estorno de inventário |
| 21 | 17.012 | Transferência entre depósitos |

**Para obter última compra de um material:**
```sql
SELECT TOP 1
    m.MATERIAL, m.DATA, m.QTDE,
    m.VLRTOTAL / NULLIF(m.QTDE, 0) AS vlr_unitario_compra
FROM MT_MOVIMENTACAO m
WHERE m.MATERIAL = :codigo_material
  AND m.EVENTO = 3       -- entrada de compra
  AND m.STATUS <> 'C'    -- não estornada
ORDER BY m.DATA DESC
```

---

## 12. Módulo de Vendas — Margens (VLRCUSTO)

`VE_PEDIDOITENS` possui campos de custo:

| Coluna | Preenchimento | Descrição |
|---|---|---|
| `VLRCUSTO` | ~81% dos itens | **Custo real do item** — usar para cálculo de margem |
| `VLRMATERIAL` | ~99% dos itens | ⚠️ **NÃO é custo** — confirmado igual ao `VLRUNITARIO` (preço de venda). Ignorar para margem. |
| `VLRDESCONTO` | variável | Valor de desconto concedido por item |
| `PERCDESCONTO` | variável | Percentual de desconto |

**Cálculo de margem:**
```sql
(i.VLRUNITARIO - i.VLRCUSTO) / NULLIF(i.VLRUNITARIO, 0) * 100 AS margem_pct
```

> Usar `VLRCUSTO` apenas quando `> 0`. Itens sem custo = margem indisponível.

**Última compra via `MT_MOVIMENTACAO`:**
```sql
SELECT TOP 1
    m.DATA, ROUND(m.VLRTOTAL / NULLIF(m.QTDE, 0), 4) AS vlr_unit_compra
FROM MT_MOVIMENTACAO m
WHERE m.MATERIAL = :codigo AND m.EVENTO = 3 AND m.QTDE > 0 AND m.STATUS <> 'C'
ORDER BY m.DATA DESC
```

**Margem típica fev/2026:** 52,7% média | 8 alertas críticos (< 20%) | 81,3% cobertura

---

## 13. Problemas Encontrados e Soluções

| Problema | Causa | Solução |
|---|---|---|
| Erro de driver na conexão | Máquina tem `SQL Server` (legado), não `ODBC Driver 17` | Usar `DRIVER=SQL Server` no `.env` |
| Senha com `;` quebrava conexão | `;` é separador da connection string | Escapar com `{Consulta@;}` |
| Coluna `RAZAOSOCIAL` não existe | Tabela usa `RAZAO` | Usar `c.RAZAO` e `v.RAZAO` |
| Coluna `DTORCAMENTO` não existe | Nome real é `DTCADASTRO` | Usar `o.DTCADASTRO` |
| Duplicatas nos resultados (FLAGSUB) | Item dividido gera o original + os filhos | `AND i.FLAGSUB <> 'S'` |
| Pedidos de exportação aparecendo | TPVENDA=23 não cobre todos os casos (ex: NOV com TPVENDA doméstico) | Excluir via `FN_FORNECEDORES.UF = 'EX'` |
| Itens cancelados aparecendo | Filtro só em nível de pedido | `AND i.STATUS <> 'C'` nos itens também |
| Encoding no terminal PowerShell | Caracteres especiais em português | `$env:PYTHONIOENCODING="utf-8"` |
