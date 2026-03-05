@echo off
setlocal
cd /d "%~dp0"

echo Atualizando dashboards...
python fase4_dashboard.py --no-open
if errorlevel 1 (
  echo.
  echo Falha ao gerar dashboards. Verifique o Python e as dependencias.
  pause
  exit /b 1
)

echo.
echo Abrindo dashboards para comparacao...
start "Dashboard Base" "exports\dashboard_base.html"
start "Dashboard Sem Contrato" "exports\dashboard_sem_contrato.html"

echo.
echo Pronto: Base e Sem Contrato abertos no navegador.
endlocal
