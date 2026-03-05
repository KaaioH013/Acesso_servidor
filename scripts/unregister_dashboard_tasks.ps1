$ErrorActionPreference = 'SilentlyContinue'

$taskDaily = "ERP_Dashboard_Daily_0700"
$taskIntraday = "ERP_Dashboard_Intraday_2h_0800_1800"
$taskValidacao = "ERP_Dashboard_Validation_0710"

schtasks /Delete /TN $taskDaily /F | Out-Null
schtasks /Delete /TN $taskIntraday /F | Out-Null
schtasks /Delete /TN $taskValidacao /F | Out-Null

# compatibilidade com nome antigo
schtasks /Delete /TN "ERP_Dashboard_Hourly_0800_1800" /F | Out-Null

Write-Host "Tarefas removidas (se existiam):"
Write-Host " - $taskDaily"
Write-Host " - $taskIntraday"
Write-Host " - $taskValidacao"
