# Acesso_servidor

Sistema de apoio para analise comercial e operacional em cima do ERP (SQL Server), com foco em pecas.

## Objetivo

Consolidar consultas e relatorios para:

- painel comercial diario
- carteira e prioridades de faturamento
- cotações e conversao
- relatorios especificos por vendedor/UF
- validacoes de consistencia

## Stack

- Python 3
- pandas
- SQLAlchemy + pyodbc
- openpyxl
- plotly
- SQL Server (acesso leitura)

## Estrutura principal

- `src/conexao.py`: conexao com o banco
- `fase4_dashboard.py`: gera `exports/dashboard.html`
- `validar_dashboard.py`: validacao do dashboard
- `relatorio_vendedor_cobranca.py`: relatorio de faturado por vendedor (UF/cidade/cotacao/vencimento)
- `relatorio_pagas_estado.py`: pagamentos por estado
- `relatorio_rh_salarios.py`: diagnostico inicial de salarios RH
- `scripts/*.ps1`: automacao no Windows (refresh/validacao/task scheduler)

## Configuracao

1. Crie o arquivo `.env` com os dados de conexao:

```env
DB_SERVER=192.168.0.5
DB_PORT=1433
DB_DATABASE=INDUSTRIAL
DB_USERNAME=SEU_USUARIO
DB_PASSWORD=SUA_SENHA
DB_DRIVER=SQL Server
```

2. Instale dependencias:

```bash
pip install -r requirements.txt
```

## Como rodar

### Dashboard

```bash
python fase4_dashboard.py
```

Somente gerar HTML (sem abrir navegador):

```bash
python fase4_dashboard.py --no-open
```

### Validacao

```bash
python validar_dashboard.py
```

### Relatorio de cobranca por vendedor

Padrao: inicio do ano ate hoje

```bash
python relatorio_vendedor_cobranca.py
```

Exemplo com periodo fechado:

```bash
python relatorio_vendedor_cobranca.py --inicio 2025-01-01 --fim 2025-12-31
```

## Exemplos prontos

Comandos mais usados na rotina:

1. Atualizar dashboard sem abrir browser:

```bash
python fase4_dashboard.py --no-open
```

2. Validar consistencia apos atualizar dashboard:

```bash
python validar_dashboard.py
```

3. Gerar cobranca do ano atual ate hoje:

```bash
python relatorio_vendedor_cobranca.py
```

4. Gerar cobranca da ultima semana:

```bash
python relatorio_vendedor_cobranca.py --inicio 2026-03-01 --fim 2026-03-07
```

5. Gerar cobranca por UF (exemplo MG):

```bash
python relatorio_vendedor_cobranca.py --uf MG --inicio 2026-01-01 --fim 2026-12-31
```

6. Gerar apenas vencidos no periodo:

```bash
python relatorio_vendedor_cobranca.py --inicio 2026-01-01 --fim 2026-12-31 --somente-vencidos
```

7. Relatorio RH com alertas:

```bash
python relatorio_rh_salarios.py --meses 12 --meses-sem-reajuste 12 --limite-delta-pct 30
```

## Saidas

Arquivos gerados ficam em `exports/` (dashboard, planilhas e relatorios).

## Observacoes

- As consultas comerciais usam filtros de pecas (exclui bombas `MATERIAL LIKE '8%'`).
- O projeto esta preparado para uso em rotina operacional com agendamento no Windows.
