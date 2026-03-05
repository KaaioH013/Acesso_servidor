$ErrorActionPreference = 'Stop'

$workspace = "C:\DEV\Acesso_servidor"
$scriptPath = Join-Path $workspace "validar_dashboard.py"
$logsDir = Join-Path $workspace "logs"

if (-not (Test-Path $logsDir)) {
    New-Item -Path $logsDir -ItemType Directory | Out-Null
}

$logFile = Join-Path $logsDir "dashboard_validation.log"

try {
    Set-Location -Path $workspace
    $env:PYTHONIOENCODING = "utf-8"
    $env:PYTHONUTF8 = "1"

    $start = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[$start] START validation" | Out-File -FilePath $logFile -Append -Encoding utf8

    py -3 $scriptPath --somente ambos 2>&1 | Out-File -FilePath $logFile -Append -Encoding utf8
    $exitCode = $LASTEXITCODE

    if ($exitCode -ne 0) {
        throw "Validação falhou com exit code $exitCode"
    }

    $ok = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[$ok] OK validation" | Out-File -FilePath $logFile -Append -Encoding utf8
}
catch {
    $err = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[$err] ERROR validation: $($_.Exception.Message)" | Out-File -FilePath $logFile -Append -Encoding utf8
    throw
}
