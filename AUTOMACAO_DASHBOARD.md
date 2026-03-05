# Automação do Dashboard (Windows Task Scheduler)

## Como ficou configurado

Foram criadas duas tarefas no Agendador do Windows:

- `ERP_Dashboard_Daily_0700`
  - Executa todos os dias às 07:00.
- `ERP_Dashboard_Intraday_2h_0800_1800`
  - Executa diariamente a partir de 08:00, repetindo a cada 120 minutos por 11 horas.
  - Cobertura prática: 08h, 10h, 12h, 14h, 16h, 18h.
- `ERP_Dashboard_Validation_0710`
  - Executa todos os dias às 07:10 para validar os dados após o refresh das 07:00.

Script de execução usado pelas tarefas:
- `scripts/run_dashboard_refresh.ps1`

Esse runner:
- chama `py -3 fase4_dashboard.py --no-open`;
- grava log em `logs/dashboard_refresh.log`;
- faz até 3 tentativas em caso de falha transitória.

## Comandos de operação

Registrar tarefas:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\register_dashboard_tasks.ps1
```

Remover tarefas:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\unregister_dashboard_tasks.ps1
```

Rodar atualização manual (mesmo caminho da automação):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run_dashboard_refresh.ps1
```

Rodar validação de consistência (base e sem contrato):

```powershell
py -3 .\validar_dashboard.py --somente ambos
```

## Como validar se está funcionando

1. Ver tarefas:

```powershell
schtasks /Query /TN ERP_Dashboard_Daily_0700 /V /FO LIST
schtasks /Query /TN ERP_Dashboard_Intraday_2h_0800_1800 /V /FO LIST
schtasks /Query /TN ERP_Dashboard_Validation_0710 /V /FO LIST
```

2. Disparar uma execução de teste:

```powershell
schtasks /Run /TN ERP_Dashboard_Daily_0700
```

3. Conferir log:

```powershell
Get-Content .\logs\dashboard_refresh.log -Tail 80
```

4. Abrir saída:

- `exports/dashboard.html`

5. Conferir relatório de validação:

- `exports/validacao_dashboard_YYYYMMDD_HHMMSS.xlsx`

## Modelo de arquitetura atual

- Banco SQL Server (ERP Sectra)
- Script Python (`fase4_dashboard.py`)
- Arquivo HTML atualizado automaticamente (`exports/dashboard.html`)
- Task Scheduler chama o processo sem intervenção manual

## HTML ou EXE?

- **Agora**: HTML gerado automaticamente é a opção mais simples e estável.
- **EXE**: opcional (empacotar com PyInstaller), útil para distribuição sem Python.
- **Futuro (estável)**: migrar para Streamlit com autenticação, mantendo as consultas e regras já validadas.
