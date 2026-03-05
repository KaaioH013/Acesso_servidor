$ErrorActionPreference = 'Stop'

$workspace = "C:\DEV\Acesso_servidor"
$scriptPath = Join-Path $workspace "fase4_dashboard.py"
$logsDir = Join-Path $workspace "logs"

if (-not (Test-Path $logsDir)) {
    New-Item -Path $logsDir -ItemType Directory | Out-Null
}

$logFile = Join-Path $logsDir "dashboard_refresh.log"

$maxTentativas = 3
$esperaSegundos = 10

try {
    Set-Location -Path $workspace
    $env:PYTHONIOENCODING = "utf-8"
    $env:PYTHONUTF8 = "1"

    for ($tentativa = 1; $tentativa -le $maxTentativas; $tentativa++) {
        $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        "[$stamp] START refresh (tentativa $tentativa/$maxTentativas)" | Out-File -FilePath $logFile -Append -Encoding utf8

        py -3 $scriptPath --no-open 2>&1 | Out-File -FilePath $logFile -Append -Encoding utf8
        $exitCode = $LASTEXITCODE

        if ($exitCode -eq 0) {
            $ok = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            "[$ok] OK refresh" | Out-File -FilePath $logFile -Append -Encoding utf8
            break
        }

        $fail = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        "[$fail] WARN refresh falhou com exit code $exitCode" | Out-File -FilePath $logFile -Append -Encoding utf8

        if ($tentativa -lt $maxTentativas) {
            Start-Sleep -Seconds $esperaSegundos
        }
        else {
            throw "Falha após $maxTentativas tentativas (exit code $exitCode)."
        }
    }
}
catch {
    $errStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[$errStamp] ERROR refresh: $($_.Exception.Message)" | Out-File -FilePath $logFile -Append -Encoding utf8
    throw
}
