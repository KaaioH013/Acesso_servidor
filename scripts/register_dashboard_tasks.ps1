$ErrorActionPreference = 'Stop'

$workspace = "C:\DEV\Acesso_servidor"
$runner = Join-Path $workspace "scripts\run_dashboard_refresh.ps1"
$runnerValidacao = Join-Path $workspace "scripts\run_dashboard_validation.ps1"

if (-not (Test-Path $runner)) {
    throw "Runner não encontrado: $runner"
}
if (-not (Test-Path $runnerValidacao)) {
    throw "Runner de validação não encontrado: $runnerValidacao"
}

$taskDaily = "ERP_Dashboard_Daily_0700"
$taskIntraday = "ERP_Dashboard_Intraday_2h_0800_1800"
$taskValidacao = "ERP_Dashboard_Validation_0710"

$action = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$runner`""
$actionValidacao = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$runnerValidacao`""

schtasks /Create /TN $taskDaily /TR $action /SC DAILY /ST 07:00 /F | Out-Null
schtasks /Create /TN $taskIntraday /TR $action /SC DAILY /ST 08:00 /RI 120 /DU 11:00 /F | Out-Null
schtasks /Create /TN $taskValidacao /TR $actionValidacao /SC DAILY /ST 07:10 /F | Out-Null

Write-Host "Tarefas registradas com sucesso:"
Write-Host " - $taskDaily"
Write-Host " - $taskIntraday"
Write-Host " - $taskValidacao"
